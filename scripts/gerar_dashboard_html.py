"""
Gera o dashboard financeiro em HTML (mobile-first) a partir de
data/processed/transacoes_tratadas.csv.

Versão HTML do dashboard, para acesso rápido pelo celular via link — a
versão em Excel (gerar_dashboard.py) continua sendo a entrega "de
trabalho" (fórmulas editáveis, planilha de verdade). Esta gera uma página
única e autocontida: os dados tratados são embutidos como JSON e todos os
KPIs, gráficos (SVG desenhado em JS, sem biblioteca) e listas são
calculados no navegador a partir deles — mesmo princípio de "uma fonte de
verdade" das fórmulas SUMIFS da planilha.

dashboard_template.html é o esqueleto (HTML/CSS/JS); este script só
substitui o marcador /*__DATA__*/ pelas transações em JSON e escreve o
resultado em public/index.html — pasta que o Vercel publica diretamente
(Root Directory = public, sem build step).
"""

import json
import os

import pandas as pd


def caminhos():
    pasta_raiz = os.path.join(os.path.dirname(__file__), "..")
    csv_tratado = os.path.join(pasta_raiz, "data", "processed", "transacoes_tratadas.csv")
    template = os.path.join(os.path.dirname(__file__), "dashboard_template.html")
    pasta_public = os.path.join(pasta_raiz, "public")
    os.makedirs(pasta_public, exist_ok=True)
    html_saida = os.path.join(pasta_public, "index.html")
    return csv_tratado, template, html_saida


def montar_dados_json(df: pd.DataFrame) -> str:
    df = df.sort_values("data")
    linhas = [
        [row.data.strftime("%Y-%m-%d"), row.descricao, row.categoria, row.tipo,
         round(float(row.valor), 2), row.status]
        for row in df.itertuples(index=False)
    ]
    return json.dumps(linhas, ensure_ascii=False, separators=(",", ":"))


def main():
    csv_tratado, template_path, html_saida = caminhos()
    df = pd.read_csv(csv_tratado, parse_dates=["data"])

    dados_json = montar_dados_json(df)
    template = open(template_path, encoding="utf-8").read()
    html_final = template.replace("/*__DATA__*/", dados_json)

    with open(html_saida, "w", encoding="utf-8") as f:
        f.write(html_final)

    tamanho_kb = os.path.getsize(html_saida) / 1024
    print(f"Dashboard HTML gerado com sucesso: {html_saida}")
    print(f"Linhas de dados: {len(df)} | Tamanho do arquivo: {tamanho_kb:.1f} KB")


if __name__ == "__main__":
    main()
