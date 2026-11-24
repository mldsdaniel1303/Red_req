import argparse
import json
import sqlite3
from contextlib import closing
from pathlib import Path

from openpyxl.utils.exceptions import InvalidFileException
from selenium.common.exceptions import WebDriverException

from config import ANALISTA_DESTINO, PASTA_RELATORIOS, PLANILHA_ORIGINAL
from distribuicao import CaixaGeral, DiarioTentativas, redistribuir_lote, simular
from navegador import abrir_chrome, inspecionar_caixa_geral
from planilha import exportar_relatorios, ler_planilha, resumir


def _menu():
    print("Red_req")
    print("1. Gerar Excel e CSV dos requerimentos")
    print("2. Simular a leitura de até cinco requerimentos no SCH")
    print("3. Inspecionar os controles da Caixa Geral")
    print("0. Sair")
    return {"1": "excel", "2": "simular", "3": "inspecionar"}.get(input("Opção > ").strip())


def _conectar():
    navegador = abrir_chrome()
    try:
        inspecionar_caixa_geral(navegador)
    except ValueError:
        input("Faça login no Chrome, abra Gerenciar → Caixa Geral e pressione Enter aqui. ")
        inspecionar_caixa_geral(navegador)
    return navegador


def gerar_relatorios(leitura, resultados, pasta=PASTA_RELATORIOS):
    for indicador, quantidade in resumir(leitura, resultados).items():
        print(f"{indicador:<45} {quantidade}")
    for caminho in exportar_relatorios(leitura, resultados, pasta):
        print(caminho)


def _autorizar_piloto(numero):
    print(f"Requerimento selecionado {numero}  •  Destino {ANALISTA_DESTINO}")
    return input("Confira a seleção no Chrome e digite o número completo para confirmar > ").strip() == numero


def _autorizar_lote(maximo, processados):
    print(f"Lote automático  •  Máximo de {maximo} requerimentos  •  Destino {ANALISTA_DESTINO}")
    print(f"Requerimentos com resultado registrado no histórico  {processados}")
    return input("Digite SIM para iniciar o lote > ").strip().casefold() == "sim"


def _executar_lote(leitura, numero=None, limite=None, automatico=False):
    caixa = CaixaGeral(_conectar())
    pasta = PASTA_RELATORIOS / "execucao"
    pasta.mkdir(parents=True, exist_ok=True)
    resultados = {}
    with closing(sqlite3.connect(pasta / "tentativas.sqlite3")) as conexao:
        diario = DiarioTentativas(conexao)
        try:
            return redistribuir_lote(
                leitura, caixa, diario, resultados, numero_alvo=numero,
                autorizar=None if automatico else _autorizar_piloto,
                autorizar_lote=_autorizar_lote if automatico else None,
                limite=limite,
            )
        finally:
            diario.completar_relatorio(leitura, resultados, numero_alvo=numero)
            gerar_relatorios(leitura, resultados, pasta / "piloto" if numero else pasta)


def main():
    parser = argparse.ArgumentParser(description="Red_req — validação e consulta de requerimentos do SCH.")
    modo = parser.add_mutually_exclusive_group()
    modo.add_argument("--excel", action="store_const", const="excel", dest="modo")
    modo.add_argument("--simular", action="store_const", const="simular", dest="modo")
    modo.add_argument("--inspecionar-sch", action="store_const", const="inspecionar", dest="modo")
    modo.add_argument(
        "--redistribuir", action="store_const", const="redistribuir", dest="modo",
        help="Executa as redistribuições reais do lote no SCH.",
    )
    parser.add_argument("--planilha", type=Path, default=PLANILHA_ORIGINAL)
    escopo = parser.add_mutually_exclusive_group()
    escopo.add_argument("--requerimento", metavar="NNNNN/AA", help="Limita --redistribuir ao piloto supervisionado.")
    escopo.add_argument("--limite", type=int, metavar="N", help="Processa até N requerimentos no lote.")
    parser.add_argument("--automatico", action="store_true", help="Pede uma única confirmação antes do lote.")
    argumentos = parser.parse_args()
    if argumentos.requerimento is not None and argumentos.modo != "redistribuir":
        parser.error("--requerimento exige --redistribuir.")
    if argumentos.automatico and (argumentos.modo != "redistribuir" or argumentos.requerimento is not None):
        parser.error("--automatico exige --redistribuir e é incompatível com --requerimento.")
    if argumentos.limite is not None:
        if argumentos.modo != "redistribuir":
            parser.error("--limite exige --redistribuir.")
        if argumentos.limite < 1:
            parser.error("--limite deve ser maior que zero.")

    try:
        escolhido = argumentos.modo or _menu()
        if escolhido is None:
            return 0
        if escolhido == "inspecionar":
            print(json.dumps(inspecionar_caixa_geral(_conectar()), ensure_ascii=False, indent=2))
            return 0
        leitura = ler_planilha(argumentos.planilha)
        if escolhido == "redistribuir":
            numero = argumentos.requerimento
            if numero is not None and not any(registro.numero == numero for registro in leitura.validos):
                raise ValueError("Informe um requerimento da planilha no formato NNNNN/AA.")
            concluido = _executar_lote(leitura, numero, argumentos.limite, argumentos.automatico)
            return 0 if concluido else 1
        resultados = simular(leitura, CaixaGeral(_conectar())) if escolhido == "simular" else {}
        gerar_relatorios(leitura, resultados)
    except (EOFError, KeyboardInterrupt):
        print("Execução encerrada.")
        return 1
    except (OSError, InvalidFileException, ValueError, WebDriverException, sqlite3.Error) as erro:
        print(f"Falha ao concluir a operação. {str(erro).splitlines()[0]}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
