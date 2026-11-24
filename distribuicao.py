import re
import sqlite3
from collections import Counter

from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from config import (
    ANALISTA_DESTINO, BOTAO_CONFIRMAR, BOTAO_FILTRAR, CAMPO_REQUERIMENTO, CHECKBOX_LINHA,
    CONTAINER_RESULTADOS, LIMITE_SIMULACAO, NOME_ANALISTA, PAGINACAO_RESULTADOS,
    SELECAO_TABELA, TABELA_RESULTADOS, TEMPO_SCH, TIPO_ACAO, URL_CAIXA_GERAL,
)
from navegador import verificar_url_caixa_geral
from planilha import Resultado, horario_local


PAGINACAO = re.compile(
    r"\(Registro:\s*(\d+)\s*-\s*(\d+)\s+de\s+(\d+),\s*Página:\s*(\d+)\s+de\s+(\d+)\)"
)

MONITOR_AJAX = """
const [origem, alvos, seletorGrade] = arguments;
const eventos = jQuery(document);
window.redReqLimparAjax?.();
window.redReqAjax = null;
window.redReqAjaxDetalhe = "";
const gradeAnterior = document.querySelector(seletorGrade);
const corpoAnterior = gradeAnterior?.tBodies[0];
if (!corpoAnterior) return false;

const alvosGrade = new Set(alvos);
for (let no = gradeAnterior; no && no !== document.body; no = no.parentElement) {
    if (no.id) alvosGrade.add(no.id);
}
for (const no of gradeAnterior.querySelectorAll("tbody[id], tbody [id]")) alvosGrade.add(no.id);

let corpoAlterado = false;
const observador = new MutationObserver(() => { corpoAlterado = true; });
observador.observe(corpoAnterior, {childList: true, characterData: true, subtree: true});
window.redReqGradeAtualizada = () => {
    const grade = document.querySelector(seletorGrade);
    const corpo = grade?.tBodies[0];
    corpoAlterado ||= observador.takeRecords().length > 0;
    const renovada = Boolean(corpo && (grade !== gradeAnterior || corpo !== corpoAnterior || corpoAlterado));
    if (renovada) observador.disconnect();
    return renovada;
};
window.redReqLimparAjax = () => {
    observador.disconnect();
    eventos.off("ajaxComplete.redReq");
};
eventos.off("ajaxComplete.redReq").on("ajaxComplete.redReq", function (_, xhr, config) {
    const dados = new URLSearchParams(config.data || "");
    if (dados.get("javax.faces.source") !== origem) return;
    const resposta = xhr.responseXML;
    const atualizacoes = Array.from(resposta?.querySelectorAll("update") || [], no => no.id);
    if ([401, 403].includes(xhr.status) || resposta?.querySelector("redirect")) {
        window.redReqAjax = "SESSAO";
    } else {
        window.redReqAjax = xhr.status >= 200 && xhr.status < 300
        && resposta?.documentElement?.localName === "partial-response"
        && !resposta.querySelector("error") && !xhr.pfArgs?.validationFailed
        && atualizacoes.some(alvo => alvosGrade.has(alvo)) ? "OK" : "FALHA";
    }
    window.redReqAjaxDetalhe = `HTTP ${xhr.status}. Resposta ${resposta?.documentElement?.localName ?? "sem XML"}. `
        + `Alvos retornados ${atualizacoes.join(", ") || "nenhum"}.`;
    eventos.off("ajaxComplete.redReq");
});
return true;
"""


class ErroConsulta(ValueError):
    pass


def _unico(contexto, seletor, rotulo):
    elementos = contexto.find_elements(*seletor)
    if len(elementos) != 1:
        raise ValueError(f"Esperado um único controle de {rotulo}. Encontrados {len(elementos)}.")
    return elementos[0]


def _status_analista(analista):
    if analista in ("", "-", "–", "—"):
        return None
    if " ".join(analista.split()).casefold() == ANALISTA_DESTINO.casefold():
        return "JA_ATRIBUIDO"
    return "ATRIBUIDO_A_TERCEIRO"


class CaixaGeral:
    def __init__(self, navegador):
        self.navegador = navegador
        self.espera = WebDriverWait(navegador, TEMPO_SCH)
        self.consulta_interrompida = False

    def _aguardar_ajax(self):
        self.espera.until(self._fila_ajax_vazia)

    @staticmethod
    def _fila_ajax_vazia(navegador):
        return navegador.execute_script(
            "return window.PrimeFaces?.ajax?.Queue?.isEmpty() === true"
        )

    def _monitorar_atualizacao(self, botao, *alvos):
        origem = botao.get_dom_attribute("id")
        if not origem:
            raise ValueError("O botão não tem um ID que identifique sua requisição AJAX.")
        if not self.navegador.execute_script(MONITOR_AJAX, origem, alvos, TABELA_RESULTADOS[1]):
            raise ValueError("A grade do SCH está indisponível para acompanhar a atualização.")

    def _aguardar_resultado(self, numero):
        assinatura_anterior = None

        def conferir(navegador):
            nonlocal assinatura_anterior
            verificar_url_caixa_geral(navegador)
            estado = navegador.execute_script("return window.redReqAjax")
            if estado == "SESSAO":
                raise ValueError("O SCH recusou a sessão ou redirecionou a consulta. Confira o login institucional.")
            if estado == "FALHA":
                detalhe = navegador.execute_script("return window.redReqAjaxDetalhe")
                raise ValueError(f"A resposta AJAX do SCH não confirmou a atualização da grade. {detalhe}")
            if (estado != "OK" or not self._fila_ajax_vazia(navegador)
                    or not navegador.execute_script("return window.redReqGradeAtualizada?.() === true")):
                assinatura_anterior = None
                return False
            try:
                campo = _unico(navegador, CAMPO_REQUERIMENTO, "Nº do Requerimento")
                if campo.get_property("value") != numero:
                    raise ValueError("O filtro mudou durante a atualização da grade.")
                resultado = self.resultado_atual(numero)
            except StaleElementReferenceException:
                assinatura_anterior = None
                return False
            assinatura = () if resultado is None else (resultado[0].id, resultado[1])
            estavel = assinatura == assinatura_anterior
            assinatura_anterior = assinatura
            return (resultado,) if estavel else False

        return self.espera.until(conferir)[0]

    def _verificar_sessao(self):
        verificar_url_caixa_geral(self.navegador)
        _unico(self.navegador, CAMPO_REQUERIMENTO, "Nº do Requerimento")
        _unico(self.navegador, BOTAO_FILTRAR, "Filtrar")

    def restabelecer_caixa_geral(self):
        self.consulta_interrompida = True
        self.navegador.get(URL_CAIXA_GERAL)
        self._verificar_sessao()
        self._aguardar_ajax()
        if self._selecao():
            raise ErroConsulta("A Caixa Geral ainda tem uma seleção ativa.")
        self.consulta_interrompida = False

    def filtrar(self, numero):
        if self.consulta_interrompida:
            raise ErroConsulta("Reabra a Caixa Geral e valide a sessão antes de consultar novamente.")
        try:
            return self._filtrar(numero)
        except (ValueError, WebDriverException) as erro:
            self.consulta_interrompida = True
            detalhe = _erro_sch(erro).observacao
            try:
                self._verificar_sessao()
            except (ValueError, WebDriverException):
                detalhe += " A sessão da Caixa Geral não pôde ser validada. Confira o login institucional."
            raise ErroConsulta(detalhe) from erro

    def _filtrar(self, numero):
        verificar_url_caixa_geral(self.navegador)
        self._aguardar_ajax()
        campo = _unico(self.navegador, CAMPO_REQUERIMENTO, "Nº do Requerimento")
        botao = _unico(self.navegador, BOTAO_FILTRAR, "Filtrar")
        if not campo.is_enabled() or not botao.is_enabled():
            raise ValueError("O campo de requerimento ou o botão Filtrar está desabilitado.")

        # A máscara do SCH insere a barra a partir dos sete dígitos.
        campo.click()
        campo.send_keys(Keys.CONTROL, "a")
        campo.send_keys(numero.replace("/", ""))
        if campo.get_property("value") != numero:
            raise ValueError("O valor exibido no filtro difere do requerimento da planilha.")
        self._monitorar_atualizacao(botao, CONTAINER_RESULTADOS[1])
        botao.send_keys(Keys.SPACE)
        return self._aguardar_resultado(numero)

    def resultado_atual(self, numero):
        tabela = _unico(self.navegador, TABELA_RESULTADOS, "tabela de requerimentos")
        cabecalhos = [th.text.strip() for th in tabela.find_elements(By.CSS_SELECTOR, "thead tr[role='row'] th")]
        if cabecalhos.count("Nº Req.") != 1 or cabecalhos.count("Analista") != 1:
            raise ValueError("As colunas Nº Req. e Analista não foram identificadas unicamente.")
        indice_numero = cabecalhos.index("Nº Req.")
        indice_analista = cabecalhos.index("Analista")

        paginadores = tabela.find_elements(*PAGINACAO_RESULTADOS)
        paginacao = [PAGINACAO.fullmatch(elemento.text.strip()) for elemento in paginadores]
        if not paginacao or any(item is None for item in paginacao):
            raise ValueError("A quantidade de resultados exibida pelo SCH não pôde ser validada.")
        contagens = {tuple(map(int, item.groups())) for item in paginacao}
        if len(contagens) != 1:
            raise ValueError("Os paginadores da tabela exibem contagens diferentes.")

        linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody > tr[role='row']")
        if contagens == {(0, 0, 0, 1, 1)} and not linhas:
            return None
        if contagens != {(1, 1, 1, 1, 1)} or len(linhas) != 1:
            raise ValueError("O SCH não retornou exatamente uma linha de resultado.")
        linha = linhas[0]
        celulas = linha.find_elements(By.CSS_SELECTOR, "td")
        if len(celulas) != len(cabecalhos) or celulas[indice_numero].text.strip() != numero:
            raise ValueError("A linha retornada pelo SCH não corresponde ao número pesquisado.")
        return linha, celulas[indice_analista].text.strip()

    def _selecao(self):
        return _unico(self.navegador, SELECAO_TABELA, "seleção de requerimentos").get_property("value")

    def _opcoes(self, controle):
        select = _unico(self.navegador, (By.ID, f"{controle}_input"), controle)
        return [(opcao.get_property("textContent").strip(), opcao.get_property("value"),
                 opcao.get_property("selected")) for opcao in select.find_elements(By.TAG_NAME, "option")]

    def _selecionar_opcao(self, controle, rotulo):
        opcoes = [opcao for opcao in self._opcoes(controle) if opcao[0] == rotulo]
        if len(opcoes) != 1 or not opcoes[0][1]:
            raise ValueError(f"A opção {rotulo} não foi encontrada unicamente.")
        if not opcoes[0][2]:
            gatilho = (By.CSS_SELECTOR, f"[id='{controle}'] .ui-selectonemenu-trigger")
            _unico(self.navegador, gatilho, controle).click()
            painel = (By.ID, f"{controle}_panel")
            self.espera.until(EC.visibility_of_element_located(painel))
            menu = _unico(self.navegador, painel, controle)
            itens = [item for item in menu.find_elements(By.TAG_NAME, "li")
                     if item.is_displayed() and item.text.strip() == rotulo]
            if len(itens) != 1:
                raise ValueError(f"O menu não apresenta uma única opção {rotulo}.")
            itens[0].click()
            self._aguardar_ajax()
        self._conferir_opcao(controle, rotulo)

    def _conferir_opcao(self, controle, rotulo):
        opcoes = self._opcoes(controle)
        if [texto for texto, _, selecionado in opcoes if selecionado] != [rotulo]:
            raise ValueError(f"A opção selecionada difere de {rotulo}.")

    def preparar(self, numero):
        verificar_url_caixa_geral(self.navegador)
        self._aguardar_ajax()
        encontrado = self.resultado_atual(numero)
        if encontrado is None or _status_analista(encontrado[1]) or self._selecao():
            raise ValueError("A linha está ausente, atribuída ou já há uma seleção na tabela.")
        linha, _ = encontrado
        if not linha.get_dom_attribute("data-rk"):
            raise ValueError("A linha não possui uma chave de seleção do SCH.")
        _unico(linha, CHECKBOX_LINHA, "checkbox do requerimento").click()
        self._aguardar_ajax()
        self._selecionar_opcao(TIPO_ACAO, "Redistribuir")
        self._selecionar_opcao(NOME_ANALISTA, ANALISTA_DESTINO)
        self._conferir_acao(numero)

    def _conferir_acao(self, numero):
        verificar_url_caixa_geral(self.navegador)
        campo = _unico(self.navegador, CAMPO_REQUERIMENTO, "Nº do Requerimento")
        encontrado = self.resultado_atual(numero)
        if campo.get_property("value") != numero or encontrado is None:
            raise ValueError("O requerimento selecionado difere do requerimento da planilha.")
        linha, analista = encontrado
        chave = linha.get_dom_attribute("data-rk")
        checkbox = _unico(linha, (By.CSS_SELECTOR, "input[type='checkbox']"), "seleção da linha")
        tabela = _unico(self.navegador, TABELA_RESULTADOS, "tabela de requerimentos")
        marcados = tabela.find_elements(By.CSS_SELECTOR, "tbody > tr[role='row'] input[type='checkbox']:checked")
        if _status_analista(analista) or not chave or self._selecao() != chave or marcados != [checkbox]:
            raise ValueError("A marcação exclusiva da linha ou a ausência de analista não foi confirmada.")
        self._conferir_opcao(TIPO_ACAO, "Redistribuir")
        self._conferir_opcao(NOME_ANALISTA, ANALISTA_DESTINO)
        botao = _unico(self.navegador, BOTAO_CONFIRMAR, "Confirmar")
        if not botao.is_displayed() or not botao.is_enabled():
            raise ValueError("O botão Confirmar não está disponível.")
        return botao

    def confirmar(self, numero):
        self._aguardar_ajax()
        botao = self._conferir_acao(numero)
        self._monitorar_atualizacao(botao, "form:tableCaixaGeral", CONTAINER_RESULTADOS[1])
        botao.click()
        self._aguardar_resultado(numero)


def simular(leitura, caixa, registrar=print):
    resultados = {}
    for registro in leitura.validos[:LIMITE_SIMULACAO]:
        try:
            encontrado = caixa.filtrar(registro.numero)
            if encontrado is None:
                resultado = Resultado("NAO_ENCONTRADO", horario_local(), "Consulta retornou zero registros.")
            else:
                _, analista = encontrado
                resultado = Resultado(
                    _status_analista(analista) or "SIMULADO_SEM_ANALISTA", horario_local(),
                    "Analista conferido na tabela do SCH em consulta de leitura.", analista,
                )
        except (ValueError, WebDriverException) as erro:
            resultado = _erro_sch(erro)
        resultados[registro.numero] = resultado
        registrar(f"{registro.numero}  {resultado.status}")
        if resultado.status == "ERRO_CONSULTA":
            break
    return resultados


def _erro_sch(erro):
    if isinstance(erro, TimeoutException):
        detalhe = "O SCH não atualizou a tabela ou o formulário dentro do tempo de espera."
    elif isinstance(erro, WebDriverException):
        detalhe = f"O Chrome não concluiu a interação ({type(erro).__name__})."
    else:
        detalhe = str(erro)
    return Resultado("ERRO_CONSULTA" if isinstance(erro, ErroConsulta) else "ERRO", horario_local(), detalhe)


class DiarioTentativas:
    def __init__(self, conexao: sqlite3.Connection):
        self.conexao = conexao
        self.conexao.execute("PRAGMA synchronous = FULL")
        self.conexao.execute("""
            CREATE TABLE IF NOT EXISTS tentativas (
                numero TEXT PRIMARY KEY,
                analista_anterior TEXT NOT NULL,
                analista_destino TEXT NOT NULL,
                tentativa_em TEXT NOT NULL,
                verificado_em TEXT
            )
        """)
        self.conexao.execute("""
            CREATE TABLE IF NOT EXISTS resultados (
                numero TEXT PRIMARY KEY,
                status TEXT NOT NULL,
                data_hora TEXT NOT NULL,
                observacao TEXT NOT NULL,
                analista_anterior TEXT NOT NULL,
                analista_atual TEXT
            )
        """)
        self.conexao.execute("""
            CREATE TABLE IF NOT EXISTS revisoes_tentativas (
                id INTEGER PRIMARY KEY,
                numero TEXT NOT NULL,
                status TEXT NOT NULL,
                data_hora TEXT NOT NULL,
                observacao TEXT NOT NULL,
                analista_atual TEXT
            )
        """)
        with self.conexao:
            self.conexao.execute("BEGIN IMMEDIATE")
            colunas = {coluna[1] for coluna in self.conexao.execute("PRAGMA table_info(resultados)")}
            if "analista_atual" not in colunas:
                self.conexao.execute("ALTER TABLE resultados ADD COLUMN analista_atual TEXT")

    def consultar(self, numero):
        return self.conexao.execute("""
            SELECT t.analista_anterior, COALESCE(t.verificado_em, (
                SELECT r.data_hora FROM revisoes_tentativas r
                WHERE r.numero = t.numero AND r.status = 'REDISTRIBUIDO'
                ORDER BY r.id LIMIT 1
            ))
            FROM tentativas t WHERE t.numero = ?
        """, (numero,)).fetchone()

    def registrar_antes_do_clique(self, numero, analista_anterior):
        with self.conexao:
            cursor = self.conexao.execute("""
                INSERT INTO tentativas
                    (numero, analista_anterior, analista_destino, tentativa_em) VALUES (?, ?, ?, ?)
                ON CONFLICT(numero) DO NOTHING
            """, (numero, analista_anterior, ANALISTA_DESTINO, horario_local()))
        return cursor.rowcount == 1

    def registrar_verificacao(self, numero, data_hora):
        with self.conexao:
            self.conexao.execute(
                "UPDATE tentativas SET verificado_em = ? WHERE numero = ? AND verificado_em IS NULL",
                (data_hora, numero),
            )

    def registrar_resultado(self, numero, resultado):
        with self.conexao:
            self.conexao.execute("""
                INSERT INTO resultados (numero, status, data_hora, observacao, analista_anterior, analista_atual)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(numero) DO UPDATE SET
                    status = excluded.status,
                    data_hora = excluded.data_hora,
                    observacao = excluded.observacao,
                    analista_anterior = excluded.analista_anterior,
                    analista_atual = excluded.analista_atual
            """, (numero, resultado.status, resultado.data_hora, resultado.observacao,
                  resultado.analista_anterior, resultado.analista_atual))

    def registrar_revisao(self, numero, resultado):
        with self.conexao:
            self.conexao.execute("""
                INSERT INTO revisoes_tentativas (numero, status, data_hora, observacao, analista_atual)
                VALUES (?, ?, ?, ?, ?)
            """, (numero, resultado.status, resultado.data_hora, resultado.observacao, resultado.analista_atual))

    def completar_relatorio(self, leitura, resultados, numero_alvo=None):
        numeros = {numero_alvo} if numero_alvo else {registro.numero for registro in leitura.validos}
        for numero, status, data_hora, observacao, anterior, atual in self.conexao.execute("""
            SELECT numero, status, data_hora, observacao, analista_anterior, analista_atual FROM resultados
        """):
            if numero in numeros:
                resultados.setdefault(numero, Resultado(status, data_hora, observacao, anterior, atual))
        for numero, anterior, tentativa_em, verificado_em in self.conexao.execute("""
            SELECT numero, analista_anterior, tentativa_em, verificado_em FROM tentativas
        """):
            if numero not in numeros:
                continue
            salvo = resultados.get(numero)
            status_tentativa = "REDISTRIBUIDO" if verificado_em else "INDETERMINADO"
            if salvo and salvo.status in ("PENDENTE_REVISAO", status_tentativa):
                continue
            if verificado_em:
                resultado = Resultado("REDISTRIBUIDO", verificado_em, "Destino já verificado no SCH.", anterior)
            else:
                resultado = Resultado(
                    "INDETERMINADO", tentativa_em, "Tentativa registrada. Aguardando verificação no SCH.", anterior
                )
            resultados[numero] = resultado

        for numero, status, data_hora, observacao, anterior, atual in self.conexao.execute("""
            SELECT r.numero, r.status, r.data_hora, r.observacao, t.analista_anterior, r.analista_atual
            FROM revisoes_tentativas r JOIN tentativas t ON t.numero = r.numero
            ORDER BY r.id
        """):
            salvo = resultados.get(numero)
            if numero in numeros and (salvo is None or data_hora >= salvo.data_hora):
                resultados[numero] = Resultado(status, data_hora, observacao, anterior, atual)


def revisar_tentativa(numero, caixa, diario):
    tentativa = diario.consultar(numero)
    if tentativa is None or tentativa[1] is not None:
        raise ValueError("A revisão de leitura exige uma tentativa registrada sem confirmação.")

    encontrado = caixa.filtrar(numero)
    analista_atual = (encontrado[1] or "-") if encontrado else None
    if analista_atual is None:
        status = "PENDENTE_REVISAO"
        evidencia = "O filtro retornou zero registros."
    elif _status_analista(analista_atual) == "JA_ATRIBUIDO":
        status = "REDISTRIBUIDO"
        evidencia = f"O filtro retornou uma linha exata com {analista_atual} na coluna Analista."
    elif _status_analista(analista_atual) == "ATRIBUIDO_A_TERCEIRO":
        status = "PENDENTE_REVISAO"
        evidencia = f"O filtro retornou uma linha exata atribuída a {analista_atual}, diferente do destino da tentativa."
    else:
        status = "INDETERMINADO"
        evidencia = "O filtro retornou uma linha exata sem analista. O caso aguarda decisão do operador."

    resultado = Resultado(status, horario_local(), f"Revisão de leitura no SCH. {evidencia}",
                          tentativa[0], analista_atual)
    diario.registrar_revisao(numero, resultado)
    return resultado


def _verificar_tentativa(numero, tentativa, encontrado, diario):
    analista_anterior, verificado_em = tentativa
    data_hora = horario_local()
    analista_atual = (encontrado[1] or "-") if encontrado else None
    status_atual = _status_analista(analista_atual) if analista_atual is not None else None
    if status_atual == "JA_ATRIBUIDO":
        if verificado_em:
            return Resultado(
                "REDISTRIBUIDO", verificado_em, "Destino anterior reconfirmado no SCH.", analista_anterior, analista_atual
            )
        diario.registrar_verificacao(numero, data_hora)
        return Resultado(
            "REDISTRIBUIDO", data_hora, "Destino verificado em nova consulta ao SCH.", analista_anterior, analista_atual
        )
    if verificado_em or encontrado is None or status_atual == "ATRIBUIDO_A_TERCEIRO":
        estado_atual = f"Analista atual no SCH: {encontrado[1] or 'sem analista'}." if encontrado else "Requerimento ausente no SCH."
        return Resultado(
            "PENDENTE_REVISAO", data_hora, f"Histórico de tentativa para Leonardo. {estado_atual}",
            analista_anterior, analista_atual,
        )
    return Resultado(
        "INDETERMINADO", data_hora,
        "Tentativa registrada. A consulta ao SCH ainda não confirmou o destino.", analista_anterior, analista_atual,
    )


def _redistribuir_um(numero, caixa, diario, autorizar=None, retomar_consulta=False):
    tentativa = diario.consultar(numero)
    consulta_inicial = True
    try:
        if retomar_consulta and tentativa is None:
            caixa.restabelecer_caixa_geral()
        encontrado = caixa.filtrar(numero)
        consulta_inicial = False
        if tentativa:
            return _verificar_tentativa(numero, tentativa, encontrado, diario)
        if encontrado is None:
            return Resultado("NAO_ENCONTRADO", horario_local(), "Consulta retornou zero registros.")
        _, analista = encontrado
        status = _status_analista(analista)
        if status:
            return Resultado(status, horario_local(), "Analista atual conferido.", analista, analista)
        caixa.preparar(numero)
        if autorizar is not None and not autorizar(numero):
            return Resultado("CANCELADO", horario_local(), "Operação cancelada antes do registro da tentativa.",
                             analista, analista or "-")
        if not diario.registrar_antes_do_clique(numero, analista):
            tentativa = diario.consultar(numero)
            return _verificar_tentativa(numero, tentativa, caixa.filtrar(numero), diario)
        caixa.confirmar(numero)
        return _verificar_tentativa(numero, (analista, None), caixa.filtrar(numero), diario)
    except Exception as erro:
        # A gravação pode ter sido concluída mesmo que a chamada tenha levantado uma exceção.
        tentativa = diario.consultar(numero)
        if not tentativa and not isinstance(erro, (ValueError, WebDriverException)):
            raise
        resultado = _erro_sch(erro)
        if not tentativa:
            if consulta_inicial:
                return Resultado("ERRO_CONSULTA", resultado.data_hora, resultado.observacao)
            return resultado
        status = "PENDENTE_REVISAO" if tentativa[1] else "INDETERMINADO"
        return Resultado(status, resultado.data_hora, resultado.observacao, tentativa[0])


def redistribuir_lote(leitura, caixa, diario, resultados, numero_alvo=None, autorizar=None, limite=None,
                     autorizar_lote=None):
    acumulados = {}
    diario.completar_relatorio(leitura, acumulados, numero_alvo)
    registros = tuple(r for r in leitura.validos if numero_alvo is None or r.numero == numero_alvo)
    total = len(registros)
    if numero_alvo is None:
        registros = tuple(
            r for r in registros
            if r.numero not in acumulados or acumulados[r.numero].status in ("ERRO", "ERRO_CONSULTA", "CANCELADO")
        )[:limite]
    print(f"Lote com {len(registros)} requerimentos. Histórico disponível para {len(acumulados)}/{total}.", flush=True)
    if registros and autorizar_lote is not None and not autorizar_lote(len(registros), len(acumulados)):
        print("Lote cancelado antes da primeira consulta.", flush=True)
        return True
    for indice, registro in enumerate(registros, start=1):
        numero = registro.numero
        anterior = acumulados.get(numero)
        retomar_consulta = anterior is not None and anterior.status in ("ERRO", "ERRO_CONSULTA")
        resultado = _redistribuir_um(numero, caixa, diario, autorizar, retomar_consulta)
        diario.registrar_resultado(numero, resultado)
        resultados[numero] = acumulados[numero] = resultado
        estados = Counter(item.status for item in acumulados.values())
        totais = "  |  ".join(f"{status} {quantidade}" for status, quantidade in sorted(estados.items()))
        print(
            f"Lote {indice}/{len(registros)}  {numero}  {resultado.status}  "
            f"Acumulado {len(acumulados)}/{total}  |  {totais}",
            flush=True,
        )
        if resultado.status in ("ERRO_CONSULTA", "INDETERMINADO") or (
            autorizar_lote is not None and resultado.status in ("ERRO", "PENDENTE_REVISAO")
        ):
            print(f"Lote interrompido para revisão de {numero}. {resultado.observacao}", flush=True)
            return False
        if resultado.status == "CANCELADO":
            break
    return True
