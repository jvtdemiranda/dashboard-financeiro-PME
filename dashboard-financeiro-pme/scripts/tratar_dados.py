"""
Lê data/raw/transacoes_brutas.csv (dado "sujo", como um export real de sistema)
e aplica um pipeline de limpeza, gerando data/processed/transacoes_tratadas.csv,
pronto para alimentar o dashboard.

Cada etapa de tratamento registra quantas linhas foram corrigidas, para
compor um pequeno "relatório de qualidade de dados" — bom para mostrar
no README do projeto o que foi encontrado e corrigido.
"""

import os

import pandas as pd

# Mapa de "categoria suja" -> "categoria correta".
# Importante: as chaves aqui já estão SEM espaços nas pontas, porque o
# .str.strip() roda antes desse mapa ser aplicado (ver padronizar_categorias).
MAPA_CATEGORIAS_CORRETAS = {
    "Alugue": "Aluguel",
    "fornecedor": "Fornecedor",
    "Salario": "Salário",
    "venda de servico": "Venda de Serviço",
    "Imposto": "Imposto",
    "marketing": "Marketing",
    "Manutencao": "Manutenção",
}


def carregar_dados_brutos(caminho_csv: str) -> pd.DataFrame:
    df = pd.read_csv(caminho_csv)
    return df


def padronizar_categorias(df: pd.DataFrame) -> int:
    """Corrige erros de digitação/espaços nas categorias. Retorna nº de linhas alteradas."""
    categoria_original = df["categoria"].copy()
    df["categoria"] = df["categoria"].str.strip()
    df["categoria"] = df["categoria"].replace(MAPA_CATEGORIAS_CORRETAS)
    # Ainda pode sobrar variação de maiúscula/minúscula; padroniza por segurança
    df["categoria"] = df["categoria"].str.strip()
    return int((categoria_original != df["categoria"]).sum())


def padronizar_datas(df: pd.DataFrame) -> int:
    """
    Unifica a coluna 'data' em um único formato datetime, mesmo vindo
    misturada (parte como timestamp ISO, parte como texto dd/mm/aaaa).
    """
    # Antes de converter, identifica quais linhas vieram como texto
    # dd/mm/aaaa (a "sujeira" real) em vez de timestamp ISO.
    eram_texto_dd_mm = df["data"].astype(str).str.match(r"^\d{2}/\d{2}/\d{4}$")

    # format="mixed" porque a coluna vem com dois formatos diferentes ao
    # mesmo tempo (timestamp ISO e texto dd/mm/aaaa); dayfirst=True resolve
    # a ambiguidade do formato texto (padrão brasileiro de data).
    df["data"] = pd.to_datetime(df["data"], format="mixed", dayfirst=True, errors="coerce")
    # Remove a parte de hora/minuto/segundo: para este dashboard só a data importa
    df["data"] = df["data"].dt.normalize()

    return int(eram_texto_dd_mm.sum())


def corrigir_sinais_valor(df: pd.DataFrame) -> int:
    """
    Garante que todas as transações fiquem com valor positivo na coluna
    'valor' — a direção (entrada/saída) já é representada pela coluna 'tipo',
    então valor negativo é sempre erro de lançamento, não uma saída "de verdade".
    """
    mascara_negativos = df["valor"] < 0
    n_corrigidos = int(mascara_negativos.sum())
    df.loc[mascara_negativos, "valor"] = df.loc[mascara_negativos, "valor"].abs()
    return n_corrigidos


def tratar_status_vazio(df: pd.DataFrame) -> int:
    """Preenche status vazio com 'Não informado' em vez de deixar em branco."""
    mascara_vazio = df["status"].isna()
    n_preenchidos = int(mascara_vazio.sum())
    df.loc[mascara_vazio, "status"] = "Não informado"
    return n_preenchidos


def remover_linhas_invalidas(df: pd.DataFrame) -> int:
    """Remove linhas onde a data não pôde ser interpretada (ficou nula após o parse)."""
    antes = len(df)
    df.dropna(subset=["data"], inplace=True)
    return antes - len(df)


def main():
    pasta_raiz = os.path.join(os.path.dirname(__file__), "..")
    caminho_bruto = os.path.join(pasta_raiz, "data", "raw", "transacoes_brutas.csv")
    pasta_processed = os.path.join(pasta_raiz, "data", "processed")
    os.makedirs(pasta_processed, exist_ok=True)
    caminho_tratado = os.path.join(pasta_processed, "transacoes_tratadas.csv")

    df = carregar_dados_brutos(caminho_bruto)
    total_linhas_original = len(df)

    n_categorias = padronizar_categorias(df)
    n_datas = padronizar_datas(df)
    n_linhas_removidas = remover_linhas_invalidas(df)
    n_sinais = corrigir_sinais_valor(df)
    n_status = tratar_status_vazio(df)

    # Ordena por data para o dashboard já ficar cronológico
    df.sort_values("data", inplace=True)
    df.to_csv(caminho_tratado, index=False)

    print("=== Relatório de Qualidade de Dados ===")
    print(f"Linhas na base bruta:              {total_linhas_original}")
    print(f"Categorias corrigidas:              {n_categorias}")
    print(f"Datas reformatadas/padronizadas:    {n_datas}")
    print(f"Linhas removidas (data inválida):   {n_linhas_removidas}")
    print(f"Sinais de valor corrigidos:         {n_sinais}")
    print(f"Status vazios preenchidos:          {n_status}")
    print(f"Linhas na base final tratada:       {len(df)}")
    print(f"\nSalvo em: {caminho_tratado}")
    print("\nCategorias finais (deve haver só 7, sem duplicata/erro):")
    print(sorted(df["categoria"].unique()))


if __name__ == "__main__":
    main()
