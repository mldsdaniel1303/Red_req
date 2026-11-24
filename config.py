from pathlib import Path

from selenium.webdriver.common.by import By

RAIZ = Path(__file__).resolve().parent
PLANILHA_ORIGINAL = RAIZ / "entrada" / "Manutenção tácita RE SCH 2026-11-23.xlsx"
PASTA_RELATORIOS = RAIZ / "relatorios"
URL_CAIXA_GERAL = "https://sistemasnet.anatel.gov.br/mosaico/sch/worklist/gerenciar/caixaGeral"
ANALISTA_DESTINO = "Leonardo Marques Campos"
TOTAL_ESPERADO = 293
LIMITE_SIMULACAO = 5
FUSO_HORARIO = "America/Sao_Paulo"
TEMPO_SCH = 25

CAMPO_REQUERIMENTO = (By.XPATH, (
    "//*[@id='form:filtros']//tr[td/label[normalize-space(.)='Nº do Requerimento:']]"
    "//input[@type='text' and contains(@class,'ui-inputmask')]"
))
BOTAO_FILTRAR = (By.XPATH, "//button[@title='Filtrar Resultado' and normalize-space(.)='Filtrar']")
CONTAINER_RESULTADOS = (By.ID, "form:caixaGeralTableContainer")
TABELA_RESULTADOS = (By.CSS_SELECTOR, "[id='form:tableCaixaGeral'] table[role='grid']")
PAGINACAO_RESULTADOS = (By.CLASS_NAME, "ui-paginator-current")
CHECKBOX_LINHA = (By.CSS_SELECTOR, "td .ui-chkbox-box")
SELECAO_TABELA = (By.NAME, "form:tableCaixaGeral_selection")
TIPO_ACAO = "form:usuariosAtribuicao1"
NOME_ANALISTA = "form:usuariosAtribuicao2"
BOTAO_CONFIRMAR = (By.XPATH, "//button[@name='form:botaoConfirmarAcao' and normalize-space(.)='Confirmar']")

# Rótulos da especificação usados para localizar candidatos durante a inspeção.
ROTULOS_CAIXA_GERAL = {
    "numero_requerimento": ["Nº do Requerimento", "N° do Requerimento", "Nº Requerimento"],
    "filtrar": ["Filtrar"],
    "tabela": ["Requerimentos Cadastrados"],
    "tipo_acao": ["Tipo de Ação"],
    "nome_analista": ["Nome do Analista"],
    "confirmar": ["Confirmar"],
}
