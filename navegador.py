from urllib.parse import urlsplit

from selenium import webdriver

from config import ROTULOS_CAIXA_GERAL, URL_CAIXA_GERAL


def abrir_chrome():
    opcoes = webdriver.ChromeOptions()
    opcoes.add_experimental_option("detach", True)
    navegador = webdriver.Chrome(options=opcoes)
    navegador.set_page_load_timeout(35)
    navegador.get(URL_CAIXA_GERAL)
    return navegador


def verificar_url_caixa_geral(navegador):
    atual = urlsplit(navegador.current_url)
    destino = urlsplit(URL_CAIXA_GERAL)
    if (atual.scheme, atual.netloc, atual.path.rstrip("/")) != (
        destino.scheme,
        destino.netloc,
        destino.path.rstrip("/"),
    ):
        raise ValueError("Abra Gerenciar → Caixa Geral após o login institucional.")


def inspecionar_caixa_geral(navegador):
    verificar_url_caixa_geral(navegador)
    resultado = navegador.execute_script(
        r"""
        const rotulos = arguments[0];
        const normalizar = (texto) => (texto || "").normalize("NFD")
            .replace(/\p{M}/gu, "").replace(/\s+/g, " ").trim()
            .replace(/\s*[:*]+$/, "").toLowerCase();
        const visivel = (elemento) => elemento.getClientRects().length > 0
            && getComputedStyle(elemento).visibility === "visible";

        const descrever = (elemento) => {
            const atributos = ["id", "name", "type", "for", "role", "class", "aria-labelledby"];
            const dados = {tag: elemento.localName};
            for (const atributo of atributos) {
                const valor = elemento.getAttribute(atributo);
                if (valor) dados[atributo] = valor;
            }
            return dados;
        };
        const nome = (elemento) => {
            const atributos = [elemento.getAttribute("aria-label"), elemento.getAttribute("title")];
            const porId = (elemento.getAttribute("aria-labelledby") || "").split(/\s+/)
                .map((id) => document.getElementById(id)?.textContent).filter(Boolean);
            const nomes = Array.from(elemento.labels || [], (label) => label.textContent);
            const texto = elemento.matches("input,select") ? "" : elemento.textContent;
            return [...atributos, ...porId, ...nomes, texto,
                elemento.matches("input[type=button],input[type=submit]") ? elemento.value : ""
            ].map(normalizar).filter(Boolean);
        };

        const elementos = Array.from(document.querySelectorAll(
            "label, legend, caption, h1, h2, h3, h4, span, a, button, input, select, " +
            "[role=button], [role=combobox], [aria-label], [aria-labelledby]"
        )).filter(visivel);
        const candidatos = {};
        for (const [campo, alternativas] of Object.entries(rotulos)) {
            const esperados = new Set(alternativas.map(normalizar));
            candidatos[campo] = elementos.filter((el) => nome(el).some((x) => esperados.has(x)))
                .map((el) => ({
                    ...descrever(el),
                    ...(el.control ? {controle: descrever(el.control)} : {})
                }));
        }
        candidatos.tabela_resultados = Array.from(document.querySelectorAll("table,[role=grid]"))
            .filter(visivel).map((tabela) => {
                const colunas = Array.from(
                    tabela.querySelectorAll("thead th,[role=columnheader]"),
                    (celula) => celula.textContent.trim()
                );
                if (!colunas.some((coluna) => normalizar(coluna) === "analista")) return null;
                const checkbox = tabela.querySelector("tbody input[type=checkbox],tbody [role=checkbox]");
                return {
                    ...descrever(tabela), colunas,
                    checkbox_da_linha: checkbox ? descrever(checkbox) : null
                };
            }).filter(Boolean);
        return candidatos;
        """,
        ROTULOS_CAIXA_GERAL,
    )
    if not resultado["tabela_resultados"] and not (
        resultado["numero_requerimento"] and resultado["filtrar"]
    ):
        raise ValueError("A página aberta ainda não exibe os controles da Caixa Geral.")
    return resultado
