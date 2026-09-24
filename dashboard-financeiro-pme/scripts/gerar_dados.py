"""
Gera uma base fake de transações financeiras de uma pequena empresa,
simulando um "export bruto de sistema": com inconsistências propositais
(erros de digitação, formatos de data diferentes, sinais errados, campos vazios).

Esse é o ponto de partida do projeto "dashboard-financeiro-pme":
1. gerar_dados.py   -> gera data/raw/transacoes_brutas.csv (dado sujo)
2. tratar_dados.py  -> lê o csv sujo, trata, e gera data/processed/transacoes_tratadas.csv
"""

import os
import random
from datetime import datetime, timedelta

import pandas as pd

# Semente fixa para o resultado ser reproduzível (sempre gera a mesma base)
random.seed(42)

# --- "Matéria-prima" para sortear os dados ---
categorias = [
    "Aluguel", "Fornecedor", "Salário", "Venda de Serviço",
    "Imposto", "Marketing", "Manutenção",
]
tipos = ["Entrada", "Saída"]
status_pagamento = ["Pago", "Pendente", "Atrasado"]
clientes_fornecedores = [
    "Cliente A", "Cliente B", "Cliente C",
    "Fornecedor X", "Fornecedor Y", "Fornecedor Z",
]


def gerar_data_aleatoria() -> datetime:
    """Sorteia uma data dentro dos últimos 180 dias (~6 meses)."""
    hoje = datetime.now()
    dias_atras = random.randint(0, 180)
    return hoje - timedelta(days=dias_atras)


def gerar_transacao() -> dict:
    """Monta uma linha (dicionário) representando uma transação financeira."""
    tipo = random.choice(tipos)
    valor = round(random.uniform(50, 5000), 2)
    return {
        "data": gerar_data_aleatoria(),
        "descricao": f"{tipo} - {random.choice(clientes_fornecedores)}",
        "categoria": random.choice(categorias),
        "tipo": tipo,
        "valor": valor,
        "status": random.choice(status_pagamento),
    }


def gerar_base(n_linhas: int = 200) -> pd.DataFrame:
    """Gera n_linhas de transações 'limpas' e retorna como DataFrame."""
    transacoes = [gerar_transacao() for _ in range(n_linhas)]
    return pd.DataFrame(transacoes)


def sujar_categorias(df: pd.DataFrame, frac: float = 0.10) -> None:
    """Introduz erros de digitação em uma fração das categorias (in-place)."""
    mapa_erros = {
        "Aluguel": "Alugue",
        "Fornecedor": "fornecedor ",
        "Salário": "Salario",
        "Venda de Serviço": "venda de servico",
        "Imposto": "Imposto ",
        "Marketing": "marketing",
        "Manutenção": "Manutencao",
    }
    idx = df.sample(frac=frac, random_state=1).index
    for i in idx:
        original = df.loc[i, "categoria"]
        df.loc[i, "categoria"] = mapa_erros.get(original, original)


def sujar_datas(df: pd.DataFrame, frac: float = 0.08) -> None:
    """Converte parte das datas para texto em formato dd/mm/aaaa (in-place)."""
    # Importante: a coluna 'data' nasce com dtype datetime64. Se a gente
    # simplesmente atribuir uma string numa célula dessas, o pandas
    # "corrige" sozinho e converte de volta para Timestamp — nossa sujeira
    # desapareceria silenciosamente. Por isso convertemos a coluna inteira
    # para dtype 'object' antes, o que permite misturar Timestamp e str.
    df["data"] = df["data"].astype(object)
    idx = df.sample(frac=frac, random_state=2).index
    for i in idx:
        df.loc[i, "data"] = df.loc[i, "data"].strftime("%d/%m/%Y")


def sujar_sinais_valor(df: pd.DataFrame, frac: float = 0.05) -> None:
    """Inverte o sinal de parte das saídas, simulando erro de lançamento (in-place)."""
    saidas = df[df["tipo"] == "Saída"]
    idx = saidas.sample(frac=frac, random_state=3).index
    for i in idx:
        df.loc[i, "valor"] = -abs(df.loc[i, "valor"])


def sujar_status_vazio(df: pd.DataFrame, frac: float = 0.04) -> None:
    """Deixa o status em branco em parte das linhas (in-place)."""
    idx = df.sample(frac=frac, random_state=4).index
    df.loc[idx, "status"] = None


def main():
    df = gerar_base(n_linhas=200)

    sujar_categorias(df)
    sujar_datas(df)
    sujar_sinais_valor(df)
    sujar_status_vazio(df)

    pasta_saida = os.path.join(os.path.dirname(__file__), "..", "data", "raw")
    os.makedirs(pasta_saida, exist_ok=True)
    caminho_csv = os.path.join(pasta_saida, "transacoes_brutas.csv")
    df.to_csv(caminho_csv, index=False)

    print(f"Base gerada com sucesso: {df.shape[0]} linhas, {df.shape[1]} colunas.")
    print(f"Salvo em: {caminho_csv}\n")
    print("Amostra (10 primeiras linhas):")
    print(df.head(10).to_string())
    print("\nValores únicos de 'categoria' (repare nas variações sujas):")
    print(sorted(df["categoria"].unique()))
    print(f"\nLinhas com status vazio: {df['status'].isna().sum()}")
    print(f"Linhas com valor negativo em 'Saída': {(df[df['tipo']=='Saída']['valor'] < 0).sum()}")


if __name__ == "__main__":
    main()
