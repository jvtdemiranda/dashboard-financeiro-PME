"""
Gera o dashboard financeiro em Excel a partir de data/processed/transacoes_tratadas.csv.

Terceira etapa do pipeline:
1. gerar_dados.py     -> gera data/raw/transacoes_brutas.csv (dado sujo)
2. tratar_dados.py    -> gera data/processed/transacoes_tratadas.csv (dado limpo)
3. gerar_dashboard.py -> gera dashboard/dashboard_financeiro.xlsx (este script)

Os KPIs e as tabelas de resumo (Fluxo Mensal, Contas a Pagar/Receber, DRE)
são escritos como fórmulas do Excel (SUMIFS) que apontam para a aba "Dados" —
se alguém editar um valor de uma transação diretamente na planilha, os totais
recalculam sozinhos. As listas de detalhamento de "Contas a Pagar/Receber"
são um extrato estático tirado do CSV tratado no momento em que este script
roda; para atualizá-las depois de uma mudança na base, é só rodar o script
de novo (mesma lógica de reprodutibilidade dos outros dois scripts).
"""

import io
import os
import re
import shutil
import zipfile
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.chart.data_source import AxDataSource, NumData, NumRef, NumVal, StrData, StrRef, StrVal
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_to_tuple
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- Paleta (dataviz skill: categórico fixo, status reservado) ---
BLUE = "2A78D6"        # Entradas / receita
RED = "E34948"         # Saídas / despesa
GOOD = "0CA30C"        # status: positivo
WARNING = "FAB219"     # status: pendente
CRITICAL = "D03B3B"    # status: atrasado
INK = "0B0B0B"
INK_SEC = "52514E"
INK_MUTED = "898781"
GRID = "E1E0D9"
SURFACE = "FCFCFB"
PAGE = "F9F9F7"
SUCCESS_TEXT = "006300"
WHITE = "FFFFFF"

CUR_FMT = '"R$" #,##0.00'
DATE_FMT = "dd/mm/yyyy"
MONTH_FMT = "mmm/yyyy"
PCT_FMT = "0.0%"

CATEGORIAS_DESPESA = [
    "Aluguel", "Fornecedor", "Imposto", "Manutenção", "Marketing", "Salário",
]

THIN = Side(style="thin", color=GRID)


def caminhos():
    pasta_raiz = os.path.join(os.path.dirname(__file__), "..")
    csv_tratado = os.path.join(pasta_raiz, "data", "processed", "transacoes_tratadas.csv")
    pasta_dashboard = os.path.join(pasta_raiz, "dashboard")
    os.makedirs(pasta_dashboard, exist_ok=True)
    xlsx_saida = os.path.join(pasta_dashboard, "dashboard_financeiro.xlsx")
    return csv_tratado, xlsx_saida


# Resultado de cada fórmula, calculado aqui em Python com a mesma regra da
# fórmula: {(aba, "B5"): 1234.5}. O openpyxl grava a fórmula sem o resultado,
# e visualizadores que não calculam (celular, prévias online) mostrariam as
# células vazias. salvar_deterministico() grava esses resultados no arquivo.
RESULTADOS = {}


def formula(ws, row, col, expr, resultado):
    c = ws.cell(row=row, column=col, value=expr)
    RESULTADOS[(ws.title, c.coordinate)] = float(resultado)
    return c


def _valores(wb, ref):
    """Valores das células de uma referência de gráfico, usando o resultado das fórmulas."""
    if ":" not in ref.split("!")[-1]:
        ref = f"{ref}:{ref.split('!')[-1]}"
    aba, (c1, l1, c2, l2) = range_to_tuple(ref)
    ws = wb[aba]
    saida = []
    for linha in range(l1, l2 + 1):
        for coluna in range(c1, c2 + 1):
            celula = ws.cell(row=linha, column=coluna)
            saida.append(RESULTADOS.get((aba, celula.coordinate), celula.value))
    return saida, ws.cell(row=l1, column=c1).number_format


def preencher_caches(wb):
    """
    Grava dentro de cada gráfico uma cópia dos dados que ele desenha (é o que
    o Excel faz ao salvar). O openpyxl grava só a referência às células, e
    visualizadores que não calculam mostrariam o gráfico em branco.
    """
    for ws in wb.worksheets:
        for grafico in ws._charts:
            for serie in grafico.series:
                valores, fmt = _valores(wb, serie.val.numRef.f)
                serie.val.numRef.numCache = NumData(
                    formatCode=fmt, ptCount=len(valores),
                    pt=[NumVal(idx=i, v=float(v)) for i, v in enumerate(valores) if v is not None])
                if serie.tx is not None and serie.tx.strRef is not None:
                    nome, _ = _valores(wb, serie.tx.strRef.f)
                    serie.tx.strRef.strCache = StrData(ptCount=1, pt=[StrVal(idx=0, v=str(nome[0]))])
                if serie.cat is not None:
                    ref = serie.cat.numRef.f if serie.cat.numRef is not None else serie.cat.strRef.f
                    rotulos, fmt = _valores(wb, ref)
                    if all(isinstance(r, datetime) for r in rotulos):
                        serie.cat = AxDataSource(numRef=NumRef(f=ref, numCache=NumData(
                            formatCode=fmt, ptCount=len(rotulos),
                            pt=[NumVal(idx=i, v=float((r - datetime(1899, 12, 30)).days)) for i, r in enumerate(rotulos)])))
                    else:
                        serie.cat = AxDataSource(strRef=StrRef(f=ref, strCache=StrData(
                            ptCount=len(rotulos), pt=[StrVal(idx=i, v=str(r).strip()) for i, r in enumerate(rotulos)])))


def eixos_visiveis(chart):
    """
    O openpyxl 3.1 não grava que os eixos são visíveis, e há relatos de
    versões recentes do Excel escondendo o eixo nesse caso (gráfico sem
    meses e sem valores). Marcar explicitamente não muda nada onde já
    funcionava.
    """
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    return chart


def draw_box(ws, r1, c1, r2, c2, color=GRID):
    """Desenha só o contorno externo de um retângulo de células (sem linhas internas)."""
    for col in range(c1, c2 + 1):
        top = ws.cell(row=r1, column=col)
        top.border = Border(top=Side(style="thin", color=color),
                             left=top.border.left, right=top.border.right, bottom=top.border.bottom)
        bot = ws.cell(row=r2, column=col)
        bot.border = Border(bottom=Side(style="thin", color=color),
                             left=bot.border.left, right=bot.border.right, top=bot.border.top)
    for row in range(r1, r2 + 1):
        left = ws.cell(row=row, column=c1)
        left.border = Border(left=Side(style="thin", color=color),
                              top=left.border.top, bottom=left.border.bottom, right=left.border.right)
        right = ws.cell(row=row, column=c2)
        right.border = Border(right=Side(style="thin", color=color),
                               top=right.border.top, bottom=right.border.bottom, left=right.border.left)


def fill_block(ws, r1, c1, r2, c2, color):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            ws.cell(row=row, column=col).fill = fill


def titulo_pagina(ws, texto, subtitulo, ultima_coluna):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ultima_coluna)
    c = ws.cell(row=1, column=1, value=texto)
    c.font = Font(name="Calibri", size=18, bold=True, color=INK)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ultima_coluna)
    s = ws.cell(row=2, column=1, value=subtitulo)
    s.font = Font(name="Calibri", size=10, italic=True, color=INK_SEC)
    ws.row_dimensions[1].height = 26
    fill_block(ws, 1, 1, 2, ultima_coluna, PAGE)


def header_row(ws, row, first_col, last_col, fill_color=BLUE):
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def tabela_excel(ws, nome, ref, estilo="TableStyleMedium9"):
    tbl = Table(displayName=nome, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=estilo, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tbl)


def montar_dados(wb, df):
    ws = wb.create_sheet("Dados")
    colunas = ["Data", "Descrição", "Categoria", "Tipo", "Valor", "Status"]
    for j, nome in enumerate(colunas, start=1):
        ws.cell(row=1, column=j, value=nome)
    header_row(ws, 1, 1, 6)

    for i, row in enumerate(df.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=row.data.to_pydatetime()).number_format = DATE_FMT
        ws.cell(row=i, column=2, value=row.descricao)
        ws.cell(row=i, column=3, value=row.categoria)
        ws.cell(row=i, column=4, value=row.tipo)
        ws.cell(row=i, column=5, value=float(row.valor)).number_format = CUR_FMT
        ws.cell(row=i, column=6, value=row.status)

    n = len(df)
    tabela_excel(ws, "tbl_dados", f"A1:F{n + 1}")

    larguras = [12, 30, 18, 10, 14, 14]
    for j, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    return ws, n


def montar_fluxo_mensal(wb, df, meses, ultima_linha_dados):
    ws = wb.create_sheet("Fluxo Mensal")
    titulo_pagina(ws, "Fluxo de Caixa Mensal",
                  f"Período: {meses[0].strftime('%b/%Y')} a {meses[-1].strftime('%b/%Y')} "
                  "· regime de caixa: só transações com status Pago (pendentes/atrasadas ficam em Contas a Pagar e Receber)",
                  5)

    cabecalhos = ["Mês", "Entradas", "Saídas", "Saldo do Mês", "Saldo Acumulado"]
    for j, nome in enumerate(cabecalhos, start=1):
        ws.cell(row=4, column=j, value=nome)
    header_row(ws, 4, 1, 5)

    primeira_linha = 5
    u = ultima_linha_dados
    pagos = df[df["status"] == "Pago"]
    acumulado = 0.0
    totais = {"B": 0.0, "C": 0.0, "D": 0.0}
    for i, mes in enumerate(meses):
        r = primeira_linha + i
        ws.cell(row=r, column=1, value=mes).number_format = MONTH_FMT
        no_mes = pagos[(pagos["data"] >= mes) & (pagos["data"] < mes + pd.DateOffset(months=1))]
        valores = {}
        for col, tipo in [(2, "Entrada"), (3, "Saída")]:
            valores[col] = no_mes.loc[no_mes["tipo"] == tipo, "valor"].sum()
            formula(ws, r, col,
                    f'=SUMIFS(Dados!$E$2:$E${u},Dados!$D$2:$D${u},"{tipo}",Dados!$F$2:$F${u},"Pago",'
                    f'Dados!$A$2:$A${u},">="&A{r},Dados!$A$2:$A${u},"<"&EDATE(A{r},1))',
                    valores[col]).number_format = CUR_FMT
        saldo = valores[2] - valores[3]
        acumulado += saldo
        formula(ws, r, 4, f"=B{r}-C{r}", saldo).number_format = CUR_FMT
        formula(ws, r, 5, f"=D{r}" if i == 0 else f"=E{r - 1}+D{r}", acumulado).number_format = CUR_FMT
        totais["B"] += valores[2]
        totais["C"] += valores[3]
        totais["D"] += saldo

    ultima_linha = primeira_linha + len(meses) - 1
    linha_total = ultima_linha + 1
    ws.cell(row=linha_total, column=1, value="Total").font = Font(bold=True)
    for col, letra in [(2, "B"), (3, "C"), (4, "D")]:
        cell = formula(ws, linha_total, col, f"=SUM({letra}{primeira_linha}:{letra}{ultima_linha})", totais[letra])
        cell.font = Font(bold=True)
        cell.number_format = CUR_FMT
    for col in range(1, 6):
        ws.cell(row=linha_total, column=col).border = Border(top=Side(style="thin", color=INK_MUTED))

    # cor condicional: saldo do mês negativo em vermelho, positivo em verde
    faixa_saldo = f"D{primeira_linha}:D{ultima_linha}"
    ws.conditional_formatting.add(
        faixa_saldo, CellIsRule(operator="lessThan", formula=["0"], font=Font(color=CRITICAL)))
    ws.conditional_formatting.add(
        faixa_saldo, CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(color=SUCCESS_TEXT)))

    larguras = [12, 16, 16, 16, 18]
    for j, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # Gráfico 1: Entradas x Saídas por mês (barras, mesma escala -> mesmo eixo)
    chart1 = eixos_visiveis(BarChart())
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Entradas x Saídas por Mês"
    chart1.style = 10
    chart1.y_axis.title = "R$"
    chart1.x_axis.title = None
    chart1.gapWidth = 60
    dados = Reference(ws, min_col=2, max_col=3, min_row=4, max_row=ultima_linha)
    categorias = Reference(ws, min_col=1, min_row=primeira_linha, max_row=ultima_linha)
    chart1.add_data(dados, titles_from_data=True)
    chart1.set_categories(categorias)
    chart1.series[0].graphicalProperties.solidFill = BLUE
    chart1.series[1].graphicalProperties.solidFill = RED
    chart1.height = 8.5
    chart1.width = 20
    ws.add_chart(chart1, "G4")

    # Gráfico 2: Saldo acumulado (série única -> sem legenda, linha)
    chart2 = eixos_visiveis(LineChart())
    chart2.title = "Saldo Acumulado"
    chart2.style = 10
    chart2.y_axis.title = "R$"
    dados2 = Reference(ws, min_col=5, min_row=4, max_row=ultima_linha)
    chart2.add_data(dados2, titles_from_data=True)
    chart2.set_categories(categorias)
    s = chart2.series[0]
    s.graphicalProperties.line.solidFill = BLUE
    s.graphicalProperties.line.width = 22000
    s.marker = Marker(symbol="circle", size=6)
    s.marker.graphicalProperties.solidFill = BLUE
    s.marker.graphicalProperties.line.solidFill = BLUE
    s.smooth = False
    chart2.legend = None
    chart2.height = 8.5
    chart2.width = 20
    ws.add_chart(chart2, "G21")

    return ws, primeira_linha, ultima_linha, linha_total


def montar_contas(wb, df, ultima_linha_dados):
    ws = wb.create_sheet("Contas a Pagar e Receber")
    titulo_pagina(ws, "Contas a Pagar e Receber",
                  "Transações com status Pendente ou Atrasado (em aberto)", 10)

    ws.cell(row=4, column=2, value="Pendente")
    ws.cell(row=4, column=3, value="Atrasado")
    ws.cell(row=4, column=4, value="Total em Aberto")
    header_row(ws, 4, 2, 4)

    ws.cell(row=5, column=1, value="A Receber (Entradas)")
    ws.cell(row=6, column=1, value="A Pagar (Saídas)")
    em_aberto = {}
    for r, tipo in [(5, "Entrada"), (6, "Saída")]:
        soma = 0.0
        for c, status in [(2, "Pendente"), (3, "Atrasado")]:
            valor = df.loc[(df["tipo"] == tipo) & (df["status"] == status), "valor"].sum()
            soma += valor
            formula(ws, r, c,
                    f'=SUMIFS(Dados!$E$2:$E${ultima_linha_dados},Dados!$D$2:$D${ultima_linha_dados},"{tipo}",'
                    f'Dados!$F$2:$F${ultima_linha_dados},"{status}")',
                    valor).number_format = CUR_FMT
        em_aberto[r] = soma
        formula(ws, r, 4, f"=B{r}+C{r}", soma).number_format = CUR_FMT
        ws.cell(row=r, column=4).font = Font(bold=True)

    ws.cell(row=7, column=1, value="Saldo em Aberto (Receber − Pagar)").font = Font(bold=True)
    saldo_cell = formula(ws, 7, 4, "=D5-D6", em_aberto[5] - em_aberto[6])
    saldo_cell.font = Font(bold=True)
    saldo_cell.number_format = CUR_FMT
    ws.conditional_formatting.add(
        "D7", CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, color=CRITICAL)))
    ws.conditional_formatting.add(
        "D7", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(bold=True, color=SUCCESS_TEXT)))

    draw_box(ws, 4, 1, 7, 4)

    receber = df[(df["tipo"] == "Entrada") & (df["status"].isin(["Pendente", "Atrasado"]))].sort_values("data")
    pagar = df[(df["tipo"] == "Saída") & (df["status"].isin(["Pendente", "Atrasado"]))].sort_values("data")

    cor_por_status = {"Pendente": WARNING, "Atrasado": CRITICAL}

    def escrever_detalhe(bloco_df, titulo, col_inicial):
        linha_titulo = 10
        linha_cabecalho = 11
        ws.merge_cells(start_row=linha_titulo, start_column=col_inicial,
                        end_row=linha_titulo, end_column=col_inicial + 4)
        tcell = ws.cell(row=linha_titulo, column=col_inicial,
                         value=f"{titulo} ({len(bloco_df)} transações)")
        tcell.font = Font(bold=True, color=INK, size=11)

        cabecalhos = ["Data", "Descrição", "Categoria", "Valor", "Status"]
        for j, nome in enumerate(cabecalhos):
            ws.cell(row=linha_cabecalho, column=col_inicial + j, value=nome)
        header_row(ws, linha_cabecalho, col_inicial, col_inicial + 4, fill_color=INK_SEC)

        for i, row in enumerate(bloco_df.itertuples(index=False), start=1):
            r = linha_cabecalho + i
            ws.cell(row=r, column=col_inicial, value=row.data.to_pydatetime()).number_format = DATE_FMT
            ws.cell(row=r, column=col_inicial + 1, value=row.descricao)
            ws.cell(row=r, column=col_inicial + 2, value=row.categoria)
            ws.cell(row=r, column=col_inicial + 3, value=float(row.valor)).number_format = CUR_FMT
            status_cell = ws.cell(row=r, column=col_inicial + 4, value=row.status)
            status_cell.font = Font(color=cor_por_status.get(row.status, INK), bold=True)

        linha_final = linha_cabecalho + len(bloco_df)
        col_letra_ini = get_column_letter(col_inicial)
        col_letra_fim = get_column_letter(col_inicial + 4)
        nome_tabela = "tbl_receber" if col_inicial == 1 else "tbl_pagar"
        tabela_excel(ws, nome_tabela, f"{col_letra_ini}{linha_cabecalho}:{col_letra_fim}{linha_final}")
        return linha_final

    escrever_detalhe(receber, "Contas a Receber em Aberto", 1)
    escrever_detalhe(pagar, "Contas a Pagar em Aberto", 7)

    larguras = {1: 12, 2: 26, 3: 16, 4: 13, 5: 13, 7: 12, 8: 26, 9: 16, 10: 13, 11: 13}
    for col, w in larguras.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.column_dimensions["F"].width = 3

    return ws


def montar_dre(wb, df, ultima_linha_dados):
    ws = wb.create_sheet("DRE")
    titulo_pagina(ws, "DRE Simplificado",
                  "Demonstrativo de Resultado — regime de competência (todas as transações lançadas, pagas ou não)", 5)

    ws.cell(row=4, column=1, value="Receita Bruta (Entradas)").font = Font(bold=True)
    receita = df.loc[df["tipo"] == "Entrada", "valor"].sum()
    receita_cell = formula(ws, 4, 2, f'=SUMIFS(Dados!$E$2:$E${ultima_linha_dados},Dados!$D$2:$D${ultima_linha_dados},"Entrada")', receita)
    receita_cell.font = Font(bold=True, color=BLUE)
    receita_cell.number_format = CUR_FMT

    ws.cell(row=5, column=1, value="Despesas Operacionais").font = Font(bold=True, color=INK_SEC)

    primeira = 6
    total_despesas = 0.0
    for i, categoria in enumerate(CATEGORIAS_DESPESA):
        r = primeira + i
        ws.cell(row=r, column=1, value=f"   {categoria}")
        valor = df.loc[(df["tipo"] == "Saída") & (df["categoria"] == categoria), "valor"].sum()
        total_despesas += valor
        formula(ws, r, 2,
                f'=SUMIFS(Dados!$E$2:$E${ultima_linha_dados},Dados!$D$2:$D${ultima_linha_dados},"Saída",'
                f'Dados!$C$2:$C${ultima_linha_dados},"{categoria}")',
                valor).number_format = CUR_FMT
    ultima = primeira + len(CATEGORIAS_DESPESA) - 1

    linha_total_desp = ultima + 1
    ws.cell(row=linha_total_desp, column=1, value="Total de Despesas").font = Font(bold=True)
    total_desp_cell = formula(ws, linha_total_desp, 2, f"=SUM(B{primeira}:B{ultima})", total_despesas)
    total_desp_cell.font = Font(bold=True, color=RED)
    total_desp_cell.number_format = CUR_FMT
    for col in (1, 2):
        ws.cell(row=linha_total_desp, column=col).border = Border(top=Side(style="thin", color=INK_MUTED))

    linha_resultado = linha_total_desp + 2
    ws.cell(row=linha_resultado, column=1, value="Resultado Líquido").font = Font(bold=True, size=12)
    resultado = receita - total_despesas
    resultado_cell = formula(ws, linha_resultado, 2, f"=B4-B{linha_total_desp}", resultado)
    resultado_cell.font = Font(bold=True, size=12)
    resultado_cell.number_format = CUR_FMT
    ref = f"B{linha_resultado}"
    ws.conditional_formatting.add(
        ref, CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, size=12, color=CRITICAL)))
    ws.conditional_formatting.add(
        ref, CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(bold=True, size=12, color=SUCCESS_TEXT)))

    linha_margem = linha_resultado + 1
    ws.cell(row=linha_margem, column=1, value="Margem Líquida").font = Font(italic=True, color=INK_SEC)
    margem_cell = formula(ws, linha_margem, 2, f"=B{linha_resultado}/B4", resultado / receita if receita else 0.0)
    margem_cell.number_format = PCT_FMT
    margem_cell.font = Font(italic=True, color=INK_SEC)

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18

    chart = eixos_visiveis(BarChart())
    chart.type = "bar"
    chart.title = "Despesas por Categoria"
    chart.style = 10
    chart.legend = None
    chart.x_axis.title = None
    chart.y_axis.title = "R$"
    dados = Reference(ws, min_col=2, min_row=primeira, max_row=ultima)
    categorias = Reference(ws, min_col=1, min_row=primeira, max_row=ultima)
    chart.add_data(dados, titles_from_data=False)
    chart.set_categories(categorias)
    chart.series[0].graphicalProperties.solidFill = RED
    chart.height = 9
    chart.width = 18
    ws.add_chart(chart, "D4")

    return ws, linha_resultado


def montar_resumo(wb, meses, linha_total_fluxo, linha_resultado_dre):
    ws = wb.create_sheet("Resumo", 0)
    titulo_pagina(ws, "Dashboard Financeiro — PME",
                  f"Período: {meses[0].strftime('%b/%Y')} a {meses[-1].strftime('%b/%Y')} "
                  "· dados tratados a partir de um export bruto simulado (ver README)", 9)

    def tile(row_label, row_value, col1, col2, titulo, aba, celula, cor_valor, fmt=CUR_FMT):
        ws.merge_cells(start_row=row_label, start_column=col1, end_row=row_label, end_column=col2)
        lbl = ws.cell(row=row_label, column=col1, value=titulo)
        lbl.font = Font(size=9, color=INK_SEC)
        ws.merge_cells(start_row=row_value, start_column=col1, end_row=row_value, end_column=col2)
        ref = f"'{aba}'!{celula}" if " " in aba else f"{aba}!{celula}"
        val = formula(ws, row_value, col1, "=" + ref, RESULTADOS[(aba, celula)])
        val.font = Font(size=18, bold=True, color=cor_valor)
        val.number_format = fmt
        fill_block(ws, row_label, col1, row_value, col2, SURFACE)
        draw_box(ws, row_label, col1, row_value, col2)
        ws.row_dimensions[row_label].height = 16
        ws.row_dimensions[row_value].height = 28

    tile(4, 5, 1, 3, "RECEBIDO (CAIXA)", "Fluxo Mensal", f"B{linha_total_fluxo}", BLUE)
    tile(4, 5, 4, 6, "PAGO (CAIXA)", "Fluxo Mensal", f"C{linha_total_fluxo}", RED)
    tile(4, 5, 7, 9, "SALDO DE CAIXA", "Fluxo Mensal", f"D{linha_total_fluxo}", INK)

    tile(7, 8, 1, 3, "CONTAS A RECEBER (EM ABERTO)", "Contas a Pagar e Receber", "D5", BLUE)
    tile(7, 8, 4, 6, "CONTAS A PAGAR (EM ABERTO)", "Contas a Pagar e Receber", "D6", RED)
    tile(7, 8, 7, 9, "RESULTADO LÍQUIDO (DRE)", "DRE", f"B{linha_resultado_dre}", INK)

    ws.cell(row=10, column=1, value="Fluxo de caixa mensal").font = Font(bold=True, size=11, color=INK)

    chart = eixos_visiveis(BarChart())
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = None
    chart.style = 10
    chart.y_axis.title = "R$"
    fluxo_ws = wb["Fluxo Mensal"]
    primeira_linha_fluxo = linha_total_fluxo - len(meses)
    ultima_linha_fluxo = linha_total_fluxo - 1
    dados = Reference(fluxo_ws, min_col=2, max_col=3, min_row=4, max_row=ultima_linha_fluxo)
    categorias = Reference(fluxo_ws, min_col=1, min_row=primeira_linha_fluxo, max_row=ultima_linha_fluxo)
    chart.add_data(dados, titles_from_data=True)
    chart.set_categories(categorias)
    chart.series[0].graphicalProperties.solidFill = BLUE
    chart.series[1].graphicalProperties.solidFill = RED
    chart.height = 9
    chart.width = 24
    ws.add_chart(chart, "A11")

    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 13

    nota = ws.cell(row=28, column=1,
                    value="Dados tratados a partir de um export financeiro simulado com sujeira "
                          "proposital (categorias, datas, sinais e status) — ver README do projeto "
                          "para o pipeline completo de tratamento e o relatório de qualidade de dados.")
    nota.font = Font(size=8, italic=True, color=INK_MUTED)
    ws.merge_cells(start_row=28, start_column=1, end_row=28, end_column=9)


def salvar_deterministico(wb, caminho):
    """
    Mesmos dados -> arquivo idêntico, byte a byte. O openpyxl carimba a
    hora atual em cada arquivo interno do .xlsx (um zip) e na data de
    "modificado"; fixar essas datas permite que o CI confira se a cópia
    publicada em public/ está em dia com os dados.
    """
    wb.properties.created = wb.properties.modified = datetime(2026, 1, 1)
    # O Excel recalcula tudo ao abrir; os resultados gravados abaixo servem
    # pra quem abre num visualizador que não calcula fórmulas.
    wb.calculation.fullCalcOnLoad = True
    arquivo_da_aba = {f"xl/worksheets/sheet{i}.xml": ws.title for i, ws in enumerate(wb.worksheets, start=1)}
    gravados = 0
    buffer = io.BytesIO()
    wb.save(buffer)
    with zipfile.ZipFile(buffer) as origem, zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as destino:
        for item in origem.infolist():
            conteudo = origem.read(item.filename)
            aba = arquivo_da_aba.get(item.filename)
            if aba:
                for (nome, celula), valor in RESULTADOS.items():
                    if nome != aba:
                        continue
                    conteudo, n = re.subn(
                        rb'(<c r="' + celula.encode() + rb'"[^>]*>)(<f>.*?</f>)(?:<v\s*/>|<v></v>)',
                        rb"\g<1>\g<2><v>" + repr(round(valor, 10)).encode() + b"</v>", conteudo)
                    gravados += n
            if item.filename == "docProps/core.xml":
                conteudo = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2026-01-01T00:00:00Z", conteudo)
            info = zipfile.ZipInfo(item.filename, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            destino.writestr(info, conteudo)
    if gravados != len(RESULTADOS):
        raise RuntimeError(f"{len(RESULTADOS) - gravados} fórmula(s) ficaram sem resultado gravado — "
                           "a planilha apareceria vazia em visualizadores que não calculam")


def main():
    csv_tratado, xlsx_saida = caminhos()
    df = pd.read_csv(csv_tratado, parse_dates=["data"])

    meses = pd.date_range(df["data"].min().to_period("M").to_timestamp(),
                           df["data"].max().to_period("M").to_timestamp(), freq="MS")
    meses = list(meses)

    wb = Workbook()
    wb.remove(wb.active)  # remove a aba "Sheet" padrão

    _, n_linhas = montar_dados(wb, df)
    ultima_linha_dados = n_linhas + 1
    _, primeira_linha, ultima_linha, linha_total_fluxo = montar_fluxo_mensal(wb, df, meses, ultima_linha_dados)
    montar_contas(wb, df, ultima_linha_dados)
    _, linha_resultado_dre = montar_dre(wb, df, ultima_linha_dados)
    montar_resumo(wb, meses, linha_total_fluxo, linha_resultado_dre)

    wb["Resumo"].sheet_view.showGridLines = False
    for nome in ["Fluxo Mensal", "Contas a Pagar e Receber", "DRE"]:
        wb[nome].sheet_view.showGridLines = False

    # Impressão: cada aba cabe na largura de uma folha A4 deitada (sem isso,
    # tabelas e gráficos saíam cortados entre páginas).
    for ws in wb.worksheets:
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb["Dados"].print_title_rows = "1:1"

    wb.active = 0
    preencher_caches(wb)
    salvar_deterministico(wb, xlsx_saida)
    # Cópia em public/ pro botão "Baixar planilha" do painel publicado.
    publico = os.path.join(os.path.dirname(xlsx_saida), "..", "public", "dashboard_financeiro.xlsx")
    shutil.copyfile(xlsx_saida, publico)
    print(f"Dashboard gerado com sucesso: {xlsx_saida}")
    print(f"Linhas de dados: {n_linhas} | Meses no fluxo de caixa: {len(meses)}")


if __name__ == "__main__":
    main()
