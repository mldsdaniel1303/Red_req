# Red_req

Lê a planilha de requerimentos, preserva os números originais e prepara o formato `NNNNN/AA` usado na Caixa Geral do SCH. A coluna de origem é `REQUERIMENTO`. A conferência usa 293 requerimentos distintos como referência e registra a diferença no relatório.

## Primeira etapa

A planilha original está em `entrada/`. Os arquivos `relatorios/Requerimentos.xlsx` e `relatorios/Requerimentos.csv` contêm os 293 requerimentos, na ordem da planilha. O Excel abre na aba **Requerimentos**, com filtros e números armazenados como texto. As abas **Ocorrências** e **Resumo** registram a validação e as contagens.

A simulação consultou os primeiros cinco requerimentos no SCH. Os cinco estavam sem analista e os outros 288 ficaram pendentes de consulta. Nenhuma redistribuição foi executada.

## Como executar

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

- `--excel` lê o arquivo completo e regenera os relatórios sem abrir o SCH.
- `--simular` abre o Chrome para login institucional manual e consulta até cinco requerimentos. A janela permanece aberta ao término.
- `--inspecionar-sch` lista os controles presentes no DOM da Caixa Geral. O campo Nome do Analista aparece quando o menu está em **Redistribuir**.
- `--planilha "caminho.xlsx"` indica outro arquivo de entrada.

## Controles conferidos no SCH

| Controle | Identificação |
| --- | --- |
| Nº do Requerimento | `input.ui-inputmask` na linha do rótulo **Nº do Requerimento:** em `form:filtros` |
| Filtrar | Botão de texto **Filtrar** com título **Filtrar Resultado** |
| Requerimentos Cadastrados | Grade em `form:tableCaixaGeral`, com as colunas **Nº Req.** e **Analista** |
| Checkbox | `.ui-chkbox-box` na linha do requerimento |
| Tipo de Ação | `form:usuariosAtribuicao1`, com a opção **Redistribuir** |
| Nome do Analista | `form:usuariosAtribuicao2`, com a opção exata **Leonardo Marques Campos** |
| Confirmar | Botão de texto **Confirmar** com nome `form:botaoConfirmarAcao` |

## Validação antes de redistribuir

Na inspeção de `25206/22`, **Filtrar** respondeu com XML do PrimeFaces e substituiu a grade. A automação vincula a resposta AJAX ao botão acionado e exige que a grade seja substituída ou que o corpo da tabela mude. Depois, confere a fila AJAX vazia e duas leituras consecutivas e idênticas do resultado. Quando há uma linha, o número nela deve corresponder exatamente ao filtro.

O diário `relatorios/execucao/tentativas.sqlite3` registra a tentativa antes do clique e preserva a data de verificação. Na retomada, uma nova consulta revalida tanto as tentativas indeterminadas como as confirmadas. Uma divergência com o histórico fica em `PENDENTE_REVISAO` com o estado observado no SCH. Os relatórios têm contagens distintas para `ERRO` antes do registro, `INDETERMINADO`, `PENDENTE_REVISAO` e `REDISTRIBUIDO`.

## Piloto supervisionado

Use o número expressamente autorizado para o piloto, no formato `NNNNN/AA`, após validar que consta da planilha.

```powershell
.\.venv\Scripts\python.exe main.py --redistribuir --requerimento "NNNNN/AA"
```

O comando consulta apenas esse requerimento. Se estiver livre e sem tentativa registrada, o programa prepara a linha e o destino e aguarda que o operador confira a tela e digite o número completo no terminal. Depois da confirmação no SCH, o programa pesquisa o requerimento novamente para conferir se Leonardo Marques Campos aparece na coluna **Analista**. Confira esse resultado no relatório em `relatorios/execucao/piloto/` antes do lote. Piloto e lote usam o mesmo diário SQLite.

## Lotes supervisionados

```powershell
.\.venv\Scripts\python.exe main.py --redistribuir --limite 10
```

O limite conta até 10 requerimentos consultados, na ordem da planilha. Para cada redistribuição elegível, confira a seleção no Chrome e digite o número completo no terminal, como no piloto. Uma resposta diferente encerra o lote.

Cada resultado é gravado imediatamente no mesmo diário SQLite do piloto. Repita o comando para avançar. Casos com erro antes do registro ou cancelamento voltam à fila. Tentativas já registradas e pendências ficam no histórico para consulta. Os relatórios acumulados ficam em `relatorios/execucao/`. O Chrome permanece aberto.

## Lotes automáticos

```powershell
.\.venv\Scripts\python.exe main.py --redistribuir --automatico --limite 20
```

A tela informa a quantidade máxima do lote, o analista de destino e quantos requerimentos já têm resultado registrado. Digite **SIM** para autorizar o lote inteiro. Use `--limite 50` para um lote maior ou omita `--limite` para processar todos os restantes. O comando sem `--automatico` pede confirmação individual.

O lote automático para ao registrar `ERRO`, `ERRO_CONSULTA`, `INDETERMINADO` ou `PENDENTE_REVISAO`. Um `ERRO_CONSULTA` também para o lote supervisionado. Na próxima execução, uma consulta com erro volta à fila após reabrir a Caixa Geral, verificar os controles da sessão, aguardar a fila AJAX e conferir que a seleção está vazia. Tentativas indeterminadas e pendências ficam para revisão. Os relatórios acumulados registram os analistas anterior, atual e de destino. Quando o SCH não fornece um estado atual verificável, a coluna **Analista atual** fica vazia.

Um resultado `INDETERMINADO` também para o lote supervisionado. A revisão de uma tentativa por leitura fica em `revisoes_tentativas` no SQLite. Os relatórios usam a evidência mais recente e preservam as linhas originais de tentativa e resultado.

## Testes sintéticos

```powershell
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
```

## Arquivos Python

| Arquivo | Responsabilidade |
| --- | --- |
| `main.py` | Menu, modos de execução e resumo |
| `planilha.py` | Leitura, validação e relatórios Excel e CSV |
| `navegador.py` | Chrome e inspeção da Caixa Geral |
| `distribuicao.py` | Consulta, simulação, redistribuição e diário de tentativas |
| `config.py` | Entrada, destino e seletores do SCH |
