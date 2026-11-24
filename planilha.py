import csv
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from config import ANALISTA_DESTINO, FUSO_HORARIO, TOTAL_ESPERADO


@dataclass(frozen=True)
class Registro:
    aba: str
    linha: int
    original: str
    numero: str | None = None
    ocorrencia: str = ""
    observacao: str = ""


@dataclass(frozen=True)
class Leitura:
    origem: Path
    registros: tuple[Registro, ...]
    abas: tuple[str, ...]

    @property
    def validos(self):
        return tuple(registro for registro in self.registros if not registro.ocorrencia)


@dataclass(frozen=True)
class Resultado:
    status: str
    data_hora: str
    observacao: str = ""
    analista_anterior: str = ""
    analista_atual: str | None = None


def horario_local():
    return datetime.now(ZoneInfo(FUSO_HORARIO)).isoformat(sep=" ", timespec="seconds")


def normalizar_requerimento(valor):
    if isinstance(valor, bool):
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).strip()
    separado = re.fullmatch(r"([0-9]{1,5})\s*/\s*([0-9]{2})", texto)
    if separado:
        sequencia, ano = separado.groups()
    elif re.fullmatch(r"[0-9]{3,7}", texto):
        sequencia, ano = texto[:-2], texto[-2:]
    else:
        return None
    return f"{sequencia.zfill(5)}/{ano}"


def _cabecalho(valor):
    sem_acentos = unicodedata.normalize("NFKD", str(valor)).casefold()
    return re.sub(r"[^a-z0-9]", "", sem_acentos) in {
        "requerimento",
        "norequerimento",
        "nodorequerimento",
        "nrequerimento",
        "ndorequerimento",
        "numerodorequerimento",
        "numerorequerimento",
    }


def _tem_valor(valor):
    return valor is not None and (not isinstance(valor, str) or bool(valor.strip()))


def ler_planilha(origem):
    origem = Path(origem)
    planilha = load_workbook(origem, read_only=True, keep_links=False)
    registros = []
    primeiros = {}
    abas = []
    try:
        for aba in planilha.worksheets:
            coluna = None
            ultima_linha = None
            # Algumas exportações têm formatação até o fim da folha. Só registramos
            # linhas vazias quando há dados abaixo delas, dentro da lista.
            aba.reset_dimensions()
            for linha, valores in enumerate(aba.iter_rows(values_only=True), start=1):
                if not any(_tem_valor(valor) for valor in valores):
                    continue
                if coluna is None:
                    indices = [indice for indice, valor in enumerate(valores) if _cabecalho(valor)]
                    if len(indices) > 1:
                        raise ValueError(f"Coluna REQUERIMENTO ambígua na aba {aba.title}, linha {linha}.")
                    if indices:
                        coluna = indices[0]
                        abas.append(aba.title)
                        ultima_linha = linha
                    continue

                for vazia in range(ultima_linha + 1, linha):
                    registros.append(Registro(aba.title, vazia, "", ocorrencia="VAZIO",
                                              observacao="Linha em branco entre registros."))
                ultima_linha = linha
                valor = valores[coluna] if coluna < len(valores) else None
                original = "" if valor is None else str(valor)
                numero = normalizar_requerimento(valor)
                if not _tem_valor(valor):
                    registro = Registro(aba.title, linha, original, ocorrencia="VAZIO",
                                        observacao="Célula de requerimento vazia.")
                elif numero is None:
                    registro = Registro(aba.title, linha, original, ocorrencia="INVALIDO",
                                        observacao="Formato de requerimento inválido.")
                elif numero in primeiros:
                    primeiro = primeiros[numero]
                    registro = Registro(aba.title, linha, original, numero, "DUPLICADO",
                                        f"Primeira ocorrência em {primeiro.aba}, linha {primeiro.linha}.")
                else:
                    registro = Registro(aba.title, linha, original, numero)
                    primeiros[numero] = registro
                registros.append(registro)
    finally:
        planilha.close()

    if not abas:
        raise ValueError("Nenhuma coluna REQUERIMENTO foi encontrada.")
    return Leitura(origem, tuple(registros), tuple(abas))


def resumir(leitura, resultados):
    ocorrencias = Counter(registro.ocorrencia for registro in leitura.registros)
    estados = Counter(resultado.status for resultado in resultados.values())
    validos = len(leitura.validos)
    return {
        "Requerimentos distintos esperados": TOTAL_ESPERADO,
        "Total de requerimentos distintos": validos,
        "Diferença em relação ao esperado": validos - TOTAL_ESPERADO,
        "Linhas lidas": len(leitura.registros),
        "Duplicados": ocorrencias["DUPLICADO"],
        "Vazios": ocorrencias["VAZIO"],
        "Inválidos": ocorrencias["INVALIDO"],
        "Com resultado registrado": len(resultados),
        "Sem analista na simulação": estados["SIMULADO_SEM_ANALISTA"],
        "Redistribuídos": estados["REDISTRIBUIDO"],
        "Já atribuídos a Leonardo": estados["JA_ATRIBUIDO"],
        "Atribuídos a terceiros": estados["ATRIBUIDO_A_TERCEIRO"],
        "Não encontrados no SCH": estados["NAO_ENCONTRADO"],
        "Erros antes do registro de tentativa": estados["ERRO"],
        "Erros de consulta": estados["ERRO_CONSULTA"],
        "Tentativas indeterminadas": estados["INDETERMINADO"],
        "Pendências para revisão humana": estados["PENDENTE_REVISAO"],
        "Operações canceladas": estados["CANCELADO"],
        "Sem resultado registrado": validos - len(resultados),
    }


COLUNAS = (
    "Número original", "Número para pesquisa", "Analista anterior", "Analista atual", "Analista de destino",
    "Status", "Data e hora (São Paulo)", "Observações", "Aba de origem", "Linha de origem",
)


def _linha_relatorio(registro, resultados, gerado_em):
    if registro.ocorrencia:
        resultado = Resultado(registro.ocorrencia, gerado_em, registro.observacao)
    else:
        resultado = resultados.get(registro.numero) or Resultado(
            "PENDENTE", gerado_em, "Requerimento validado na planilha."
        )
    analista_atual = resultado.analista_atual
    if analista_atual is None:
        # Relatórios também leem resultados gravados antes da inclusão desta coluna.
        if resultado.status in ("REDISTRIBUIDO", "JA_ATRIBUIDO"):
            analista_atual = ANALISTA_DESTINO
        elif resultado.status in ("ATRIBUIDO_A_TERCEIRO", "SIMULADO_SEM_ANALISTA"):
            analista_atual = resultado.analista_anterior or "-"
    return (
        registro.original,
        registro.numero or "",
        resultado.analista_anterior,
        analista_atual or "",
        ANALISTA_DESTINO if registro.numero else "",
        resultado.status,
        resultado.data_hora,
        resultado.observacao,
        registro.aba,
        registro.linha,
    )


def _adicionar_linha(aba, valores):
    aba.append(valores)
    for celula in aba[aba.max_row]:
        if isinstance(celula.value, str):
            celula.data_type = "s"
            celula.number_format = "@"


def _formatar_aba(aba, larguras, filtro=False):
    aba.freeze_panes = "A2"
    if filtro:
        aba.auto_filter.ref = aba.dimensions
    for celula in aba[1]:
        celula.fill = PatternFill("solid", fgColor="16324F")
        celula.font = Font(color="FFFFFF", bold=True)
    for indice, largura in enumerate(larguras, start=1):
        aba.column_dimensions[get_column_letter(indice)].width = largura


def _texto_csv(valor):
    # Protege os campos de texto quando o CSV for aberto no Excel.
    if isinstance(valor, str) and valor.strip() != "-" and valor.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + valor
    return valor


def exportar_relatorios(leitura, resultados, pasta):
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    excel, csv_path = pasta / "Requerimentos.xlsx", pasta / "Requerimentos.csv"
    gerado_em = horario_local()
    livro = Workbook()
    validos = livro.active
    validos.title = "Requerimentos"
    ocorrencias = livro.create_sheet("Ocorrências")
    resumo = livro.create_sheet("Resumo")
    for aba in (validos, ocorrencias):
        _adicionar_linha(aba, COLUNAS)
    _adicionar_linha(resumo, ("Indicador", "Valor"))
    _adicionar_linha(resumo, ("Arquivo de origem", leitura.origem.name))
    _adicionar_linha(resumo, ("Abas de origem", ", ".join(leitura.abas)))
    _adicionar_linha(resumo, ("Gerado em (São Paulo)", gerado_em))
    for indicador in resumir(leitura, resultados).items():
        _adicionar_linha(resumo, indicador)

    with TemporaryDirectory(dir=pasta) as temporaria:
        excel_temp = Path(temporaria) / excel.name
        csv_temp = Path(temporaria) / csv_path.name
        with csv_temp.open("w", encoding="utf-8-sig", newline="") as arquivo_csv:
            escritor = csv.writer(arquivo_csv, delimiter=";")
            escritor.writerow(COLUNAS)
            for registro in leitura.registros:
                linha = _linha_relatorio(registro, resultados, gerado_em)
                _adicionar_linha(ocorrencias if registro.ocorrencia else validos, linha)
                escritor.writerow([_texto_csv(valor) for valor in linha])
        for aba in (validos, ocorrencias):
            _formatar_aba(aba, (24, 24, 30, 32, 32, 27, 30, 58, 29, 18), filtro=True)
        _formatar_aba(resumo, (45, 48))
        livro.save(excel_temp)
        excel_temp.replace(excel)
        csv_temp.replace(csv_path)
    return excel, csv_path
