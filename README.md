# Dashboard Financeiro para Pequena Empresa/Prestador de Serviço

Projeto de portfólio: simula o fluxo real de um trabalho de análise de dados
para um pequeno negócio — receber um export "cru" de um sistema financeiro,
identificar e tratar inconsistências, e transformar isso em um dashboard
gerencial (fluxo de caixa, contas a pagar/receber, DRE simplificado).

**Publicado em [dashboard-financeiro-pme.vercel.app](https://dashboard-financeiro-pme.vercel.app)**
— atualiza sozinho a cada push na `main`.

<p align="center">
  <img src="docs/screenshot-resumo.png" width="31%" alt="Aba Resumo do dashboard, com KPIs e gráfico de fluxo de caixa mensal">
  <img src="docs/screenshot-contas.png" width="31%" alt="Aba Contas a Pagar e Receber, com lista de contas em aberto por status">
  <img src="docs/screenshot-dre.png" width="31%" alt="Aba DRE simplificado, com receita, despesas por categoria e resultado líquido">
</p>

> **Em resumo (pra quem não é da área técnica):** este é um painel que
> organiza as finanças de uma empresa — quanto entrou, quanto saiu, o
> que está pendente de pagamento — e mostra tudo em gráficos simples de
> ler, direto do celular, sem precisar abrir Excel. O diferencial é que
> ele foi feito pra lidar com dados "sujos", do jeito que realmente
> chegam de um sistema real (datas em formatos diferentes, erros de
> digitação), e ainda assim entregar o número certo no final.

## Por que esse projeto

A maioria dos portfólios mostra só o resultado bonito (o dashboard pronto).
Aqui o objetivo é mostrar o **processo inteiro**, incluindo a parte que mais
separa um profissional júnior de um sênior: lidar com dado sujo, do jeito
que ele realmente chega de um sistema real.

## Estrutura

```
.
├── README.md
├── requirements.txt
├── docs/               -> screenshots usados neste README
├── data/
│   ├── raw/           -> transacoes_brutas.csv (dado "sujo", simulando export de sistema)
│   └── processed/     -> transacoes_tratadas.csv (dado limpo, pronto para o dashboard)
├── scripts/
│   ├── gerar_dados.py          -> gera a base fake com inconsistências propositais
│   ├── tratar_dados.py         -> lê o dado bruto, aplica limpeza, gera relatório de qualidade
│   ├── gerar_dashboard.py      -> monta o dashboard final em Excel a partir do dado tratado
│   ├── gerar_dashboard_html.py -> monta a versão HTML (mobile) do dashboard
│   └── dashboard_template.html -> esqueleto HTML/CSS/JS usado pelo script acima
├── dashboard/
│   └── dashboard_financeiro.xlsx -> dashboard gerencial (Resumo, Fluxo Mensal, Contas a Pagar/Receber, DRE)
└── public/             -> o que o Vercel publica (Root Directory = public)
    ├── index.html                -> versão HTML (mobile) do dashboard
    └── dashboard_financeiro.xlsx -> cópia da planilha, pro botão "Excel" do painel
```

## O cenário simulado

Uma base de ~200 transações financeiras de uma pequena empresa fictícia,
com até 6 meses de histórico a partir da data de geração (`gerar_dados.py`
sorteia cada data em uma janela de 0 a 180 dias atrás), contendo os
seguintes problemas propositais (comuns em exports reais de sistemas):

| Problema simulado                              | % das linhas afetadas |
|-------------------------------------------------|------------------------|
| Erros de digitação/espaços na categoria          | ~10%                   |
| Datas em formato texto (dd/mm/aaaa) misturado com timestamp | ~8%       |
| Sinal de valor invertido em transações de saída  | ~5%                    |
| Status de pagamento em branco                    | ~4%                    |

## Como rodar

```bash
pip install -r requirements.txt

cd scripts
python gerar_dados.py           # gera data/raw/transacoes_brutas.csv
python tratar_dados.py          # gera data/processed/transacoes_tratadas.csv
python gerar_dashboard.py       # gera dashboard/dashboard_financeiro.xlsx
python gerar_dashboard_html.py  # gera public/index.html
```

O `tratar_dados.py` imprime um relatório de qualidade de dados ao final,
mostrando exatamente quantas linhas foram corrigidas em cada etapa — essa é
a evidência de que o pipeline está funcionando, não só "rodando sem erro".

### Exemplo de saída do relatório

```
=== Relatório de Qualidade de Dados ===
Linhas na base bruta:              200
Categorias corrigidas:              20
Datas reformatadas/padronizadas:    16
Linhas removidas (data inválida):   0
Sinais de valor corrigidos:         4
Status vazios preenchidos:          8
Linhas na base final tratada:       200
```

## Decisões de tratamento (e por quê)

- **Categorias**: normalizadas via `.str.strip()` + mapa de correção, unificando
  variações como `"Alugue"` → `"Aluguel"` e `"fornecedor "` → `"Fornecedor"`.
- **Datas**: a coluna vinha com dois formatos simultâneos (timestamp ISO e
  texto `dd/mm/aaaa`, padrão brasileiro), então cada um é parseado
  separadamente com seu próprio formato explícito — ver bug #4 abaixo, é o
  motivo de não dar pra resolver isso com uma chamada só de
  `pd.to_datetime`.
- **Valores negativos em "Saída"**: convertidos para positivo. A direção do
  fluxo (entrada/saída) já é representada pela coluna `tipo` — um valor
  negativo ali é sempre erro de lançamento, não uma "saída de verdade".
- **Status vazio**: preenchido com `"Não informado"` em vez de descartado —
  perder uma transação real por falta de status seria pior do que só marcar
  a lacuna.

## O dashboard (Excel)

Gerado por `scripts/gerar_dashboard.py` a partir de `transacoes_tratadas.csv`,
em `dashboard/dashboard_financeiro.xlsx` — e dá pra baixar direto do painel
publicado, pelo botão **Excel** no topo da página. Tem 5 abas:

- **Resumo** — capa com KPIs (recebido, pago, saldo de caixa, contas a
  receber/pagar em aberto, resultado líquido) e o gráfico de fluxo de caixa
  mensal.
- **Fluxo Mensal** — entradas, saídas, saldo do mês e saldo acumulado por mês,
  com gráfico de barras (entradas x saídas) e linha (saldo acumulado).
  **Regime de caixa**: só entram transações com status "Pago".
- **Contas a Pagar e Receber** — resumo de valores em aberto por status
  (Pendente/Atrasado) e o detalhamento linha a linha de cada conta.
- **DRE** — receita bruta, despesas por categoria e resultado líquido, com
  gráfico de despesas por categoria. **Regime de competência**: entram todas
  as transações lançadas, pagas ou não.

A diferença entre os dois regimes é proposital e é a informação mais útil
pro dono do negócio: o DRE diz se a empresa **deu lucro** no período; o
fluxo de caixa diz quanto dinheiro **de fato entrou e saiu**. Com os dados
simulados, a empresa teve R$ 61.964 de resultado, mas só R$ 21.337 de
saldo de caixa — o resto está pendente ou atrasado com clientes.
- **Dados** — a base tratada completa, com filtro no cabeçalho e o status
  de cada transação como etiqueta colorida.

**Os totais e KPIs são fórmulas do Excel (`SUMIFS`), não valores fixos** —
apontam para a aba Dados, então editar uma transação ali recalcula o resto
do arquivo automaticamente. As duas listas de "em aberto" (Contas a
Pagar/Receber) são um extrato tirado da base no momento em que o script
roda; para atualizá-las depois de mudar os dados, é só rodar
`gerar_dashboard.py` de novo — mesma lógica de reprodutibilidade dos outros
dois scripts do pipeline.

Visual igual ao do painel publicado: mesma paleta, uma fonte só em toda
a planilha, cada aba com título e uma linha dizendo o que é, e um guia das
abas no Resumo. Cores: azul para entradas, vermelho para saídas/despesas, e
uma paleta de status reservada (âmbar = pendente, vermelho = atrasado,
verde = pago/positivo) nunca reutilizada para outra coisa — o status
aparece como etiqueta colorida por formatação condicional, então a cor
acompanha se alguém editar o valor. Os dois gráficos que comparam entradas/saídas com o saldo
acumulado ficam em eixos separados (grandezas muito diferentes) em vez de um
gráfico de eixo duplo.

## O dashboard (HTML, mobile)

`scripts/gerar_dashboard_html.py` gera uma segunda versão do dashboard como
página única e autocontida (`public/index.html`), pensada para abrir direto
no celular por um link — sem precisar de Excel instalado.
Mesmas 5 visões da versão em planilha (Resumo, Fluxo Mensal, Contas a
Pagar/Receber, DRE, Dados), como abas de navegação fixas na parte de baixo
da tela, no estilo de um app:

- KPIs, gráficos (entradas x saídas por mês, saldo acumulado) e listas são
  calculados em JavaScript a partir das mesmas transações tratadas,
  embutidas como JSON na página — mesmo princípio de "uma fonte de
  verdade" das fórmulas `SUMIFS` da versão em Excel, só que recalculado no
  navegador em vez de pelo Excel.
- Gráficos são SVG desenhado à mão (sem biblioteca de gráficos), com toque
  para ver valores exatos.
- Layout responsivo mobile-first, com tema claro/escuro automático
  (segue a preferência do aparelho).
- Aba "Dados" tem busca e filtro por tipo (entrada/saída) sobre as 200
  transações.

`dashboard_template.html` guarda o HTML/CSS/JS; o script só troca o
marcador dos dados pelo JSON gerado a partir do CSV tratado — reexecutar
`gerar_dashboard_html.py` depois de mudar os dados regenera a página
inteira.

## Bugs reais encontrados no processo

Vale registrar porque são evidência de depuração real, não só "rodou sem erro":

1. **Ordem `strip` → `replace` na correção de categoria** — corrigir com
   `.replace()` antes de `.str.strip()` deixava variações como `"fornecedor "`
   escaparem da correção, porque a chave do mapa (`"fornecedor"`) não batia
   com o valor sujo com espaço. A ordem certa é limpar espaços primeiro.
2. **Pandas revertendo `dd/mm/aaaa` de volta para `Timestamp`** — atribuir uma
   string diretamente numa célula de uma coluna já `datetime64` faz o pandas
   "corrigir" sozinho e reconverter para `Timestamp`, apagando a sujeira antes
   mesmo dela existir. Precisou converter a coluna inteira para `dtype
   object` antes de sujar as datas.
3. **Relatório de qualidade contando 100% das linhas como "data reformatada"**
   — a normalização de hora (`.dt.normalize()`) alterava o valor de toda
   coluna, então uma comparação ingênua "mudou = foi reformatada" contava
   todas as 200 linhas, mascarando o número real (~8%, só as que vieram como
   texto). A correção foi marcar quais linhas eram texto **antes** do parse,
   não comparar o resultado depois.
4. **`dayfirst=True` trocando mês por dia até em timestamp ISO** — um dos mais sérios,
   e o mais difícil de perceber: `pd.to_datetime(coluna,
   format="mixed", dayfirst=True)` deveria só resolver a ambiguidade do texto
   `dd/mm/aaaa`, mas o `dayfirst` "vazava" pro formato ISO também — sempre que
   dia e mês eram os dois ≤12, `"2026-05-08"` (8 de maio) virava
   `"2026-08-05"` (5 de agosto). Isso bagunçava **39% das linhas com
   timestamp** (72 de 184), silenciosamente — sem erro, sem warning, só data
   errada. Só apareceu porque o intervalo de datas do dashboard não batia com
   os "6 meses" esperados (dava quase 12). A correção foi abandonar o parse
   único com `format="mixed"` e tratar ISO e `dd/mm/aaaa` como duas chamadas
   separadas, cada uma com o formato explícito — `dayfirst` só faz sentido
   pro texto ambíguo, nunca pro timestamp ISO (que já é ano-primeiro,
   inequívoco).
5. **Categoria e contraparte sorteadas sem relação com o tipo da transação**
   — `gerar_dados.py` sorteava `categoria` (e o cliente/fornecedor da
   descrição) de forma totalmente independente de `tipo`, então qualquer
   uma das 7 categorias podia cair tanto em Entrada quanto em Saída. O
   efeito não era só "esquisito nos dados": a aba DRE chegou a mostrar
   **"Venda de Serviço" como categoria de despesa**, e a lista de contas
   mostrava linhas como "Entrada - Fornecedor Y" (dinheiro entrando vindo
   de um fornecedor). Só ficou óbvio olhando o dashboard renderizado, não
   examinando o CSV coluna por coluna. A correção foi amarrar a lógica ao
   tipo: toda Entrada é `"Venda de Serviço"` com um cliente como
   contraparte, toda Saída sorteia entre as 6 categorias de despesa reais
   com um fornecedor como contraparte — e `CATEGORIAS_DESPESA`, em
   `gerar_dashboard.py`, parou de incluir "Venda de Serviço" na soma de
   despesas do DRE.
6. **"Fluxo de caixa" somando dinheiro que ainda não entrou** — achado numa
   revisão depois de publicado. O fluxo de caixa e o DRE somavam todas as
   transações, inclusive as pendentes e atrasadas, e o DRE ainda dizia
   "regime de caixa". Na prática, ~65% das entradas (R$ 199.713) contavam
   ao mesmo tempo como "dinheiro que entrou no caixa" e como "conta a
   receber em aberto" — o painel mostrava R$ 61.964 de saldo quando só
   R$ 21.337 tinha de fato sido pago. Um contador pegaria isso na hora.
   Corrigido separando os dois regimes: fluxo de caixa só com status
   "Pago", DRE por competência (tudo lançado), cada um com o rótulo certo.
7. **Texto do export inserido na página sem escape** — as descrições das
   transações entravam direto num `innerHTML`, e o JSON embutido na página
   não escapava `<`. Com dados simulados não há risco, mas a proposta do
   projeto é receber export de sistema real, onde uma descrição contendo
   HTML seria executada no navegador. Corrigido com a mesma proteção em
   duas camadas do [painel de criptomoedas](https://github.com/jvtdemiranda/dashboard-criptomoedas-api)
   (`escapeHtml()` no JS + escape de `<` no JSON), e testado com um
   payload real antes e depois.
8. **Planilha pouco preparada pra impressão e gráficos com risco de
   aparecer sem eixos.** Achado ao renderizar o Excel como ele sai
   impresso: as abas eram cortadas entre páginas (o resumo saía partido
   em duas folhas). Agora cada aba cabe na largura de uma folha A4 deitada
   e a tabela de dados repete o cabeçalho em cada página. Na mesma
   revisão, os eixos dos gráficos passaram a ser marcados como visíveis:
   o openpyxl 3.1 não grava essa marcação, e há relatos de versões
   recentes do Excel escondendo o eixo sem ela.
9. **Planilha aparecia vazia no celular.** Depois de colocar o botão de
   download no painel, a planilha abriu **vazia** num visualizador de
   celular. Causa: o openpyxl grava as fórmulas (`SUMIFS`, `SUM`...) mas
   não o resultado delas — o Excel do computador calcula ao abrir, mas
   visualizadores que não calculam (celular, prévias online) mostram só o
   que está gravado. As 54 células com fórmula saíam em branco, e os
   gráficos também (cada gráfico guarda uma cópia dos dados, que também
   não era gravada). Agora cada fórmula leva junto o resultado, calculado
   em Python com a mesma regra, e cada gráfico leva a cópia dos dados; a
   planilha continua com fórmulas editáveis. Pra garantir que o cálculo
   em Python e a fórmula concordam, os 54 resultados foram comparados com
   o LibreOffice recalculando a planilha do zero: nenhuma diferença. E se
   alguma fórmula ficar sem resultado, a geração falha em vez de publicar
   uma planilha vazia.
10. **Planilha com cara de rascunho.** Revendo a planilha com o mesmo
    critério visual da página: fontes misturadas (toda célula formatada
    só como "negrito" ficava sem nome de fonte, e cada programa usava a
    sua — no LibreOffice, uma serifada), a aba Dados sem título, textos
    cortados ("A Receber (Ent"), as duas listas de contas espremidas lado
    a lado, gráficos longe das tabelas, o "pendente" em amarelo claro quase
    ilegível no branco e o período escrito com mês em inglês ("Sep/2026").
    Refeita com a identidade do painel — e sem mexer nas contas: as mesmas
    54 fórmulas, conferidas de novo contra o LibreOffice recalculando do
    zero (nenhuma diferença), e os mesmos totais de antes.

## Stack

Python 3, pandas, openpyxl.
