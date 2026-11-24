import sqlite3
import csv
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch

from openpyxl import load_workbook
from selenium.common.exceptions import TimeoutException

import main
from config import ANALISTA_DESTINO, NOME_ANALISTA, TIPO_ACAO, URL_CAIXA_GERAL
from distribuicao import CaixaGeral, DiarioTentativas, ErroConsulta, _redistribuir_um, redistribuir_lote, revisar_tentativa
from planilha import Leitura, Registro, Resultado, exportar_relatorios


NUMERO = "08442/24"


class TestTentativas(unittest.TestCase):
    def setUp(self):
        pasta = self.enterContext(TemporaryDirectory())
        self.banco = Path(pasta) / "tentativas.sqlite3"
        self.conexao = sqlite3.connect(self.banco)
        self.addCleanup(self.conexao.close)
        self.diario = DiarioTentativas(self.conexao)
        self.caixa = Mock(spec=CaixaGeral)
        self.caixa.filtrar.return_value = (object(), "-")
        self.registro = Registro("Entrada", 2, "0844224", NUMERO)

    def executar(self):
        return _redistribuir_um(NUMERO, self.caixa, self.diario)

    def test_timeout_antes_do_registro(self):
        self.caixa.preparar.side_effect = TimeoutException()
        self.assertEqual(self.executar().status, "ERRO")
        self.assertIsNone(self.diario.consultar(NUMERO))
        self.caixa.confirmar.assert_not_called()

    def test_timeout_depois_do_clique(self):
        self.caixa.confirmar.side_effect = TimeoutException()
        self.assertEqual(self.executar().status, "INDETERMINADO")
        self.assertEqual(self.diario.consultar(NUMERO), ("-", None))
        self.assertEqual(self.executar().status, "INDETERMINADO")
        self.caixa.confirmar.assert_called_once_with(NUMERO)
        self.assertEqual(self.caixa.filtrar.call_count, 2)

    def test_erro_apos_commit_do_registro(self):
        gravar = self.diario.registrar_antes_do_clique

        def gravar_e_falhar(*args):
            self.assertTrue(gravar(*args))
            raise RuntimeError("Falha após o commit")

        with patch.object(self.diario, "registrar_antes_do_clique", side_effect=gravar_e_falhar):
            self.assertEqual(self.executar().status, "INDETERMINADO")
        self.assertEqual(self.diario.consultar(NUMERO), ("-", None))
        self.caixa.confirmar.assert_not_called()
        self.caixa.filtrar.return_value = (object(), ANALISTA_DESTINO)
        self.assertEqual(self.executar().status, "REDISTRIBUIDO")

    def test_interrupcao_e_retomada_reabrem_diario(self):
        self.caixa.confirmar.side_effect = KeyboardInterrupt()
        with self.assertRaises(KeyboardInterrupt):
            self.executar()
        self.conexao.close()
        nova_conexao = sqlite3.connect(self.banco)
        self.addCleanup(nova_conexao.close)
        self.diario = DiarioTentativas(nova_conexao)
        leitura = Leitura(Path("entrada.xlsx"), (self.registro,), ("Entrada",))
        resultados = {}
        self.diario.completar_relatorio(leitura, resultados)
        self.assertEqual(resultados[NUMERO].status, "INDETERMINADO")
        self.caixa.filtrar.return_value = (object(), ANALISTA_DESTINO)
        self.assertEqual(self.executar().status, "REDISTRIBUIDO")
        self.caixa.confirmar.assert_called_once_with(NUMERO)

    def test_tentativa_confirmada_diverge_do_sch(self):
        self.diario.registrar_antes_do_clique(NUMERO, "-")
        self.diario.registrar_verificacao(NUMERO, "data original")
        self.caixa.filtrar.return_value = (object(), "Outro Analista")
        resultado = self.executar()
        self.assertEqual(resultado.status, "PENDENTE_REVISAO")
        self.assertIn("Outro Analista", resultado.observacao)
        self.assertEqual(self.diario.consultar(NUMERO), ("-", "data original"))
        self.caixa.confirmar.assert_not_called()

    def test_tentativa_confirmada_revalidada(self):
        self.diario.registrar_antes_do_clique(NUMERO, "-")
        self.diario.registrar_verificacao(NUMERO, "data original")
        self.caixa.filtrar.return_value = (object(), ANALISTA_DESTINO)
        self.assertEqual(self.executar().status, "REDISTRIBUIDO")
        self.caixa.filtrar.assert_called_once_with(NUMERO)
        self.caixa.confirmar.assert_not_called()

    def test_status_sem_tentativa_anterior(self):
        casos = (
            ((object(), ANALISTA_DESTINO), "JA_ATRIBUIDO"),
            ((object(), "Outro Analista"), "ATRIBUIDO_A_TERCEIRO"),
            (None, "NAO_ENCONTRADO"),
        )
        for retorno, status in casos:
            with self.subTest(status=status):
                self.caixa.filtrar.return_value = retorno
                self.assertEqual(self.executar().status, status)
        self.caixa.preparar.assert_not_called()
        self.caixa.confirmar.assert_not_called()
        self.assertIsNone(self.diario.consultar(NUMERO))

    def test_clique_so_classificado_apos_nova_consulta(self):
        self.caixa.filtrar.side_effect = [(object(), "-"), (object(), ANALISTA_DESTINO)]

        def confirmar(numero):
            self.assertEqual(self.diario.consultar(numero), ("-", None))

        self.caixa.confirmar.side_effect = confirmar
        self.assertEqual(self.executar().status, "REDISTRIBUIDO")
        self.assertEqual(self.caixa.filtrar.call_count, 2)
        self.assertIsNotNone(self.diario.consultar(NUMERO)[1])

    def test_piloto_limita_a_consulta_e_pode_ser_cancelado(self):
        outro = Registro("Entrada", 3, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (outro, self.registro), ("Entrada",))
        resultados = {}
        autorizar = Mock(return_value=False)
        redistribuir_lote(leitura, self.caixa, self.diario, resultados, NUMERO, autorizar)
        self.caixa.filtrar.assert_called_once_with(NUMERO)
        autorizar.assert_called_once_with(NUMERO)
        self.assertEqual(resultados[NUMERO].status, "CANCELADO")
        self.caixa.confirmar.assert_not_called()
        self.assertIsNone(self.diario.consultar(NUMERO))

    def test_lotes_mantem_ordem_e_pulam_piloto_e_resultados_gravados(self):
        numeros = ("04700/24", NUMERO, "00001/24", "04000/24", "00700/24")
        leitura = Leitura(Path("entrada.xlsx"), tuple(
            Registro("Entrada", linha, numero, numero) for linha, numero in enumerate(numeros, start=2)
        ), ("Entrada",))
        self.diario.registrar_antes_do_clique(NUMERO, "-")
        self.diario.registrar_verificacao(NUMERO, "data do piloto")
        self.caixa.filtrar.return_value = (object(), "Outro Analista")
        primeiro, segundo = {}, {}

        redistribuir_lote(leitura, self.caixa, self.diario, primeiro, limite=2)
        redistribuir_lote(leitura, self.caixa, self.diario, segundo, limite=2)

        self.assertEqual(tuple(primeiro), ("04700/24", "00001/24"))
        self.assertEqual(tuple(segundo), ("04000/24", "00700/24"))
        self.assertEqual([chamada.args[0] for chamada in self.caixa.filtrar.call_args_list],
                         ["04700/24", "00001/24", "04000/24", "00700/24"])
        self.caixa.confirmar.assert_not_called()
        acumulados = {}
        self.diario.completar_relatorio(leitura, acumulados)
        self.assertEqual(set(acumulados), set(numeros))
        self.assertEqual(acumulados[NUMERO].status, "REDISTRIBUIDO")

    def test_resultado_visivel_no_sqlite_antes_da_proxima_consulta(self):
        outro = Registro("Entrada", 3, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (self.registro, outro), ("Entrada",))
        observador = sqlite3.connect(self.banco)
        self.addCleanup(observador.close)

        def consultar(numero):
            if numero == NUMERO:
                return None
            self.assertEqual(observador.execute(
                "SELECT status FROM resultados WHERE numero = ?", (NUMERO,)
            ).fetchone(), ("NAO_ENCONTRADO",))
            raise KeyboardInterrupt()

        self.caixa.filtrar.side_effect = consultar
        with self.assertRaises(KeyboardInterrupt):
            redistribuir_lote(leitura, self.caixa, self.diario, {}, limite=2)
        self.caixa.filtrar.reset_mock(side_effect=True)
        self.caixa.filtrar.return_value = None
        redistribuir_lote(leitura, self.caixa, self.diario, {}, limite=2)
        self.caixa.filtrar.assert_called_once_with(outro.numero)

    def test_pendencias_gravadas_e_tentativa_interrompida_ficam_fora_do_lote(self):
        numeros = (NUMERO, "01234/24", "04321/24", "00123/24")
        leitura = Leitura(Path("entrada.xlsx"), tuple(
            Registro("Entrada", linha, numero, numero) for linha, numero in enumerate(numeros, start=2)
        ), ("Entrada",))
        for numero in numeros[:3]:
            self.diario.registrar_antes_do_clique(numero, "-")
        self.diario.registrar_resultado(numeros[0], Resultado("INDETERMINADO", "data", "timeout"))
        self.diario.registrar_resultado(numeros[1], Resultado("PENDENTE_REVISAO", "data", "divergência"))
        self.caixa.filtrar.return_value = None
        resultados = {}

        redistribuir_lote(leitura, self.caixa, self.diario, resultados, limite=10)
        self.diario.completar_relatorio(leitura, resultados)

        self.caixa.filtrar.assert_called_once_with(numeros[3])
        self.caixa.confirmar.assert_not_called()
        self.assertEqual(resultados[numeros[0]].status, "INDETERMINADO")
        self.assertEqual(resultados[numeros[1]].status, "PENDENTE_REVISAO")
        self.assertEqual(resultados[numeros[2]].status, "INDETERMINADO")

    def test_cancelamento_supervisionado_pausa_lote_sem_gravar_tentativa(self):
        outro = Registro("Entrada", 3, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (self.registro, outro), ("Entrada",))
        autorizar = Mock(return_value=False)
        redistribuir_lote(leitura, self.caixa, self.diario, {}, autorizar=autorizar, limite=10)
        autorizar.assert_called_once_with(NUMERO)
        self.caixa.filtrar.assert_called_once_with(NUMERO)
        self.caixa.confirmar.assert_not_called()
        self.assertIsNone(self.diario.consultar(NUMERO))

    def test_automatico_pula_11_confirmados_com_ou_sem_limite(self):
        numeros = tuple(f"{indice:05d}/24" for indice in range(1, 15))
        leitura = Leitura(Path("entrada.xlsx"), tuple(
            Registro("Entrada", linha, numero, numero) for linha, numero in enumerate(numeros, start=2)
        ), ("Entrada",))
        for numero in numeros[:11]:
            self.diario.registrar_antes_do_clique(numero, "-")
            self.diario.registrar_verificacao(numero, "data histórica")
        antes = self.conexao.execute("SELECT * FROM tentativas").fetchall()
        self.caixa.filtrar.side_effect = [
            (object(), "-"), (object(), ANALISTA_DESTINO),
            (object(), ANALISTA_DESTINO), (object(), "Outra Pessoa"),
        ]
        autorizar = Mock(return_value=True)
        primeiro, segundo = {}, {}

        self.assertTrue(redistribuir_lote(leitura, self.caixa, self.diario, primeiro,
                                         limite=1, autorizar_lote=autorizar))
        autorizar.assert_called_once_with(1, 11)
        autorizar.reset_mock()
        self.assertTrue(redistribuir_lote(leitura, self.caixa, self.diario, segundo, autorizar_lote=autorizar))
        autorizar.assert_called_once_with(2, 12)

        self.caixa.confirmar.assert_called_once_with(numeros[11])
        self.assertEqual(tuple(segundo), numeros[12:])
        self.assertEqual([r.status for r in segundo.values()], ["JA_ATRIBUIDO", "ATRIBUIDO_A_TERCEIRO"])
        self.diario.completar_relatorio(leitura, segundo)
        self.assertEqual(sum(r.status == "REDISTRIBUIDO" for r in segundo.values()), 12)
        self.assertEqual(self.conexao.execute("SELECT * FROM tentativas").fetchall()[:11], antes)
        self.assertEqual(segundo[numeros[-1]].analista_atual, "Outra Pessoa")

    def test_automatico_tentativa_duplicada_nao_confirma(self):
        leitura = Leitura(Path("entrada.xlsx"), (self.registro,), ("Entrada",))
        self.caixa.preparar.side_effect = lambda numero: self.diario.registrar_antes_do_clique(numero, "-")
        resultados = {}

        self.assertFalse(redistribuir_lote(leitura, self.caixa, self.diario, resultados,
                                          autorizar_lote=Mock(return_value=True)))

        self.caixa.confirmar.assert_not_called()
        self.assertEqual(resultados[NUMERO].status, "INDETERMINADO")
        self.assertEqual(self.conexao.execute("SELECT COUNT(*) FROM tentativas").fetchone(), (1,))

    def test_automatico_grava_erro_e_para_antes_do_proximo_requerimento(self):
        outro = Registro("Entrada", 3, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (self.registro, outro), ("Entrada",))
        for etapa, erro, status, cliques in (
            ("preparar", ValueError("Analista selecionado incorreto"), "ERRO", 0),
            ("filtrar", ValueError("Sessão expirada"), "ERRO_CONSULTA", 0),
            ("confirmar", TimeoutException(), "INDETERMINADO", 1),
        ):
            with self.subTest(etapa=etapa):
                caixa = Mock(spec=CaixaGeral)
                caixa.filtrar.return_value = (object(), "-")
                getattr(caixa, etapa).side_effect = erro
                resultados = {}
                self.assertFalse(redistribuir_lote(leitura, caixa, self.diario, resultados,
                                                  autorizar_lote=Mock(return_value=True)))
                caixa.filtrar.assert_called_once_with(NUMERO)
                self.assertEqual(caixa.confirmar.call_count, cliques)
                self.assertEqual(resultados[NUMERO].status, status)
                self.assertEqual(self.conexao.execute(
                    "SELECT status, analista_atual FROM resultados WHERE numero = ?", (NUMERO,)
                ).fetchone(), (status, None))

        caixa = Mock(spec=CaixaGeral)
        caixa.filtrar.return_value = None
        self.assertTrue(redistribuir_lote(leitura, caixa, self.diario, {},
                                         autorizar_lote=Mock(return_value=True)))
        caixa.filtrar.assert_called_once_with(outro.numero)
        caixa.confirmar.assert_not_called()

    def test_timeout_da_consulta_grava_erro_e_exige_restabelecimento_na_retomada(self):
        outro = Registro("Entrada", 3, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (outro, self.registro), ("Entrada",))
        self.diario.registrar_antes_do_clique(outro.numero, "-")
        self.diario.registrar_verificacao(outro.numero, "data histórica")
        self.caixa.filtrar.side_effect = TimeoutException()
        resultados = {}

        self.assertFalse(redistribuir_lote(leitura, self.caixa, self.diario, resultados, limite=2))
        self.assertEqual(resultados[NUMERO].status, "ERRO_CONSULTA")
        self.assertIsNone(self.diario.consultar(NUMERO))
        self.caixa.preparar.assert_not_called()
        self.caixa.confirmar.assert_not_called()

        self.caixa.reset_mock(side_effect=True)
        self.caixa.restabelecer_caixa_geral.side_effect = ValueError("Login necessário")
        self.assertFalse(redistribuir_lote(leitura, self.caixa, self.diario, {}, limite=2))
        self.caixa.filtrar.assert_not_called()

        self.caixa.reset_mock(side_effect=True)
        self.caixa.filtrar.return_value = (object(), "Outro Analista")
        self.assertTrue(redistribuir_lote(leitura, self.caixa, self.diario, {}, limite=2))
        self.assertEqual([chamada[0] for chamada in self.caixa.method_calls],
                         ["restabelecer_caixa_geral", "filtrar"])
        self.caixa.filtrar.assert_called_once_with(NUMERO)
        self.assertEqual(self.diario.consultar(outro.numero), ("-", "data histórica"))

    def test_revisao_por_leitura_preserva_diario_e_resultado_original(self):
        leitura = Leitura(Path("entrada.xlsx"), (self.registro,), ("Entrada",))
        self.diario.registrar_antes_do_clique(NUMERO, "-")
        self.diario.registrar_resultado(NUMERO, Resultado("INDETERMINADO", "2026-01-01", "Sessão recusada", "-"))
        tentativa_original = self.conexao.execute("SELECT * FROM tentativas").fetchone()
        resultado_original = self.conexao.execute("SELECT * FROM resultados").fetchone()

        for analista, status in (("-", "INDETERMINADO"), ("Outro Analista", "PENDENTE_REVISAO"),
                                 (ANALISTA_DESTINO, "REDISTRIBUIDO")):
            with self.subTest(status=status):
                self.caixa.filtrar.return_value = (object(), analista)
                revisao = revisar_tentativa(NUMERO, self.caixa, self.diario)
                self.assertEqual(revisao.status, status)
                self.assertEqual(revisao.analista_atual, analista)
                relatorio = {}
                self.diario.completar_relatorio(leitura, relatorio)
                self.assertEqual(relatorio[NUMERO], revisao)

        self.assertEqual(self.conexao.execute("SELECT COUNT(*) FROM revisoes_tentativas").fetchone(), (3,))
        self.assertEqual(self.conexao.execute("SELECT * FROM tentativas").fetchone(), tentativa_original)
        self.assertEqual(self.conexao.execute("SELECT * FROM resultados").fetchone(), resultado_original)
        self.assertIsNotNone(self.diario.consultar(NUMERO)[1])
        self.caixa.preparar.assert_not_called()
        self.caixa.confirmar.assert_not_called()

    def test_perda_de_sessao_apos_clique_interrompe_lote_e_preserva_tentativas_anteriores(self):
        confirmado = Registro("Entrada", 2, "1234524", "12345/24")
        restante = Registro("Entrada", 4, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (confirmado, self.registro, restante), ("Entrada",))
        self.diario.registrar_antes_do_clique(confirmado.numero, "-")
        self.diario.registrar_verificacao(confirmado.numero, "data histórica")
        snapshot = self.conexao.execute("SELECT * FROM tentativas WHERE numero=?", (confirmado.numero,)).fetchone()
        self.caixa.confirmar.side_effect = ErroConsulta("O SCH recusou a sessão")
        resultados = {}

        self.assertFalse(redistribuir_lote(leitura, self.caixa, self.diario, resultados,
                                          limite=12, autorizar=Mock(return_value=True)))
        self.caixa.filtrar.assert_called_once_with(NUMERO)
        self.assertEqual(resultados[NUMERO].status, "INDETERMINADO")
        self.assertEqual(self.diario.consultar(NUMERO), ("-", None))
        self.assertEqual(self.conexao.execute("SELECT * FROM tentativas WHERE numero=?",
                                             (confirmado.numero,)).fetchone(), snapshot)

        caixa = Mock(spec=CaixaGeral)
        caixa.filtrar.return_value = (object(), "Outro Analista")
        self.assertTrue(redistribuir_lote(leitura, caixa, self.diario, {}, limite=12,
                                         autorizar_lote=Mock(return_value=True)))
        caixa.filtrar.assert_called_once_with(restante.numero)
        caixa.confirmar.assert_not_called()

    def test_relatorios_incluem_analista_atual_e_preservam_historico(self):
        outro = Registro("Entrada", 3, "9999924", "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), (self.registro, outro), ("Entrada",))
        self.diario.registrar_antes_do_clique(NUMERO, "-")
        self.diario.registrar_verificacao(NUMERO, "data histórica")
        self.diario.registrar_resultado(outro.numero, Resultado(
            "PENDENTE_REVISAO", "data atual", "divergência", "-", "Outra Pessoa"
        ))
        resultados = {}
        self.diario.completar_relatorio(leitura, resultados)
        excel, csv_path = exportar_relatorios(leitura, resultados, self.banco.parent)

        with csv_path.open(encoding="utf-8-sig", newline="") as arquivo:
            linhas = list(csv.DictReader(arquivo, delimiter=";"))
        self.assertEqual([r["Analista anterior"] for r in linhas], ["-", "-"])
        self.assertEqual([r["Analista atual"] for r in linhas], [ANALISTA_DESTINO, "Outra Pessoa"])
        self.assertEqual(linhas[0]["Data e hora (São Paulo)"], "data histórica")
        livro = load_workbook(excel, read_only=True)
        self.addCleanup(livro.close)
        cabecalho, *linhas_excel = livro["Requerimentos"].values
        indice = cabecalho.index("Analista atual")
        self.assertEqual([r[indice] for r in linhas_excel], [ANALISTA_DESTINO, "Outra Pessoa"])


class TestModoAutomatico(unittest.TestCase):
    def test_cli_automatico_aceita_confirmacao_unica_para_dois_cliques(self):
        numeros = (NUMERO, "99999/24")
        leitura = Leitura(Path("entrada.xlsx"), tuple(
            Registro("Entrada", linha, numero, numero) for linha, numero in enumerate(numeros, start=2)
        ), ("Entrada",))
        caixa = Mock(spec=CaixaGeral)
        caixa.filtrar.side_effect = [(object(), "-"), (object(), ANALISTA_DESTINO)] * 2
        pasta = Path(self.enterContext(TemporaryDirectory()))
        with patch("sys.argv", ["main.py", "--redistribuir", "--automatico", "--limite", "20"]), \
             patch.object(main, "ler_planilha", return_value=leitura), \
             patch.object(main, "PASTA_RELATORIOS", pasta), patch.object(main, "_conectar"), \
             patch.object(main, "CaixaGeral", return_value=caixa), patch.object(main, "gerar_relatorios"), \
             patch("builtins.input", return_value="SIM") as entrada:
            self.assertEqual(main.main(), 0)

        entrada.assert_called_once()
        self.assertEqual([chamada.args[0] for chamada in caixa.confirmar.call_args_list], list(numeros))


class TestCaixaGeral(unittest.TestCase):
    def setUp(self):
        self.navegador = Mock(current_url=URL_CAIXA_GERAL)
        self.caixa = CaixaGeral(self.navegador)
        self.campo = Mock()
        self.campo.get_property.return_value = NUMERO
        self.botao = Mock()
        self.botao.get_dom_attribute.return_value = "form:botao"

    def test_ajax_sem_substituicao_e_resultado_estavel(self):
        linha = Mock(id="mesmo elemento")
        self.caixa.resultado_atual = Mock(side_effect=[(linha, "-"), (linha, ANALISTA_DESTINO),
                                                      (linha, ANALISTA_DESTINO)])
        self.caixa.espera = Mock()
        vistos = []

        def aguardar(condicao):
            for _ in range(5):
                valor = condicao(self.navegador)
                vistos.append(valor)
                if valor:
                    return valor
            raise TimeoutException()

        self.caixa.espera.until.side_effect = aguardar
        eventos = iter([None, "OK", "OK", "OK", "OK"])
        fila = iter([False, True, True, True])
        def executar(script):
            if script == "return window.redReqAjax":
                return next(eventos)
            if "Queue" in script:
                return next(fila)
            return True

        self.navegador.execute_script.side_effect = executar
        with patch("distribuicao._unico", return_value=self.campo):
            self.assertEqual(self.caixa._aguardar_resultado(NUMERO), (linha, ANALISTA_DESTINO))
        self.assertEqual(vistos[:-1], [False, False, False, False])
        self.assertEqual(self.caixa.resultado_atual.call_count, 3)

    def test_resposta_concluida_com_grade_antiga_expira_sem_ler_linha(self):
        self.caixa.resultado_atual = Mock()
        self.caixa.espera.until = Mock()

        def aguardar(condicao):
            self.assertFalse(condicao(self.navegador))
            self.assertFalse(condicao(self.navegador))
            raise TimeoutException()

        self.caixa.espera.until.side_effect = aguardar
        self.navegador.execute_script.side_effect = lambda script: (
            "OK" if script.endswith("redReqAjax") else "Queue" in script
        )
        with self.assertRaises(TimeoutException):
            self.caixa._aguardar_resultado(NUMERO)
        self.caixa.resultado_atual.assert_not_called()

    def test_sessao_expirada_bloqueia_outra_consulta_ate_reabrir_caixa(self):
        self.caixa.espera.until = Mock(side_effect=lambda condicao: condicao(self.navegador))
        self.navegador.execute_script.return_value = "SESSAO"
        with patch.object(self.caixa, "_filtrar", side_effect=self.caixa._aguardar_resultado) as filtrar, \
             patch.object(self.caixa, "_verificar_sessao", side_effect=ValueError("Login necessário")) as verificar:
            with self.assertRaisesRegex(ErroConsulta, "Confira o login"):
                self.caixa.filtrar(NUMERO)
            verificar.assert_called_once()
            with self.assertRaisesRegex(ErroConsulta, "Reabra a Caixa Geral"):
                self.caixa.filtrar(NUMERO)
            filtrar.assert_called_once_with(NUMERO)
        with patch.object(self.caixa, "_verificar_sessao"), patch.object(self.caixa, "_aguardar_ajax"), \
             patch.object(self.caixa, "_selecao", return_value=""), patch.object(self.caixa, "_filtrar") as filtrar:
            self.caixa.restabelecer_caixa_geral()
            self.caixa.filtrar(NUMERO)
            self.navegador.get.assert_called_once_with(URL_CAIXA_GERAL)
            filtrar.assert_called_once_with(NUMERO)

    def test_grade_renovada_exige_numero_exato_ou_zero_resultados(self):
        tabela = Mock()
        linha = Mock()
        linha.find_elements.return_value = [Mock(text="99999/24"), Mock(text="-")]
        cabecalhos = [Mock(text="Nº Req."), Mock(text="Analista")]
        for quantidade, linhas in ((1, [linha]), (0, [])):
            with self.subTest(quantidade=quantidade):
                paginador = Mock(text=f"(Registro: {quantidade} - {quantidade} de {quantidade}, Página: 1 de 1)")
                tabela.find_elements.side_effect = [cabecalhos, [paginador], linhas]
                with patch("distribuicao._unico", return_value=tabela):
                    if quantidade:
                        with self.assertRaisesRegex(ValueError, "não corresponde"):
                            self.caixa.resultado_atual(NUMERO)
                    else:
                        self.assertIsNone(self.caixa.resultado_atual(NUMERO))

    def _montar_selecao(self, selecao, analista_destino, segunda_linha=False):
        linha = Mock()
        linha.get_dom_attribute.return_value = "chave-correta"
        checkbox = Mock()
        tabela = Mock()
        tabela.find_elements.return_value = [checkbox, Mock()] if segunda_linha else [checkbox]
        self.caixa.resultado_atual = Mock(return_value=(linha, "-"))
        self.caixa._selecao = Mock(return_value=selecao)
        opcoes = {TIPO_ACAO: [("Redistribuir", "REDISTRIBUIR", True)],
                  NOME_ANALISTA: [(analista_destino, "id-analista", True)]}
        self.caixa._opcoes = Mock(side_effect=opcoes.__getitem__)
        self.caixa._aguardar_ajax = Mock()
        controles = {"Nº do Requerimento": self.campo, "seleção da linha": checkbox,
                     "tabela de requerimentos": tabela, "Confirmar": self.botao}
        return patch("distribuicao._unico", side_effect=lambda ctx, seletor, rotulo: controles[rotulo])

    def test_impede_confirmar_linha_incorreta_ou_multiplas_marcadas(self):
        for chave, segunda in (("chave-errada", False), ("chave-correta", True)):
            with self.subTest(chave=chave, segunda=segunda):
                with self._montar_selecao(chave, ANALISTA_DESTINO, segunda):
                    with self.assertRaisesRegex(ValueError, "marcação exclusiva"):
                        self.caixa.confirmar(NUMERO)
        self.botao.click.assert_not_called()

    def test_impede_confirmar_destino_incorreto(self):
        with self._montar_selecao("chave-correta", "Outro Analista"):
            with self.assertRaisesRegex(ValueError, "opção selecionada difere"):
                self.caixa.confirmar(NUMERO)
        self.botao.click.assert_not_called()

    def test_impede_confirmar_se_filtro_mudou(self):
        self.campo.get_property.return_value = "99999/24"
        with self._montar_selecao("chave-correta", ANALISTA_DESTINO):
            with self.assertRaisesRegex(ValueError, "difere do requerimento"):
                self.caixa.confirmar(NUMERO)
        self.botao.click.assert_not_called()


if __name__ == "__main__":
    unittest.main()
