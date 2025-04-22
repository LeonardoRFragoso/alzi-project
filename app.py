import pandas as pd
from datetime import timedelta

# --- CONFIGURAÇÕES ---
input_path  = "TC.xlsx"
output_path = "TC_output.xlsx"
sheet_in    = "Recuperada_Planilha1"

# --- 0. Definição dos limites de tempo ---
limite1 = timedelta(minutes=45)
limite2 = timedelta(hours=1)

# --- 1. Carregar dados e normalizar cabeçalhos ---
df = pd.read_excel(
    input_path,
    sheet_name=sheet_in,
    header=14,             # linha com nomes reais das colunas
    dtype=str,
    engine="openpyxl"
)
df.columns = (
    df.columns
      .str.replace("\n", " ", regex=False)
      .str.strip()
)
print("Colunas disponíveis:", df.columns.tolist())

# --- 2. Identificar dinamicamente as colunas de data e tempo ---
def find_col(key):
    for c in df.columns:
        if key.lower() in c.lower().replace(" ", ""):
            return c
    return None

col_dt    = find_col("Dt.Entrada")
col_tempo = find_col("Tempo")
if col_dt is None or col_tempo is None:
    raise KeyError(
        f"Não encontrei as colunas esperadas.\n"
        f"Procurei por algo como 'Dt.Entrada' e 'Tempo' em: {df.columns.tolist()}"
    )

df[col_dt]    = df[col_dt].str.strip()
df[col_tempo] = df[col_tempo].str.strip()

# --- 3. Conversão de tipos ---
# 3.1 Dt.Entrada → datetime
df["Dt.Entrada_dt"] = pd.to_datetime(
    df[col_dt],
    format="%Y-%m-%d %H:%M:%S",
    errors="coerce"
)

# 3.2 Tempo → timedelta
def parse_tempo(s: str) -> timedelta:
    try:
        h, m = s.split(":")
        return timedelta(hours=int(h), minutes=int(m))
    except:
        return timedelta(0)

df["Tempo_td"] = df[col_tempo].apply(parse_tempo)

# --- 4. Categorização e agrupamento por dia ---
# extrair dia
df["dia"] = df["Dt.Entrada_dt"].dt.day

# função de categorização
def categorizar(td: timedelta) -> str:
    if td <= limite1:
        return "ATÉ 45 MIN"
    elif td <= limite2:
        return "46 MIN até 1H"
    else:
        return "> 1h"

df["categoria"] = df["Tempo_td"].apply(categorizar)

# montar resumo: linhas = dia, colunas = categoria
resumo = (
    df.groupby(["dia", "categoria"])
      .size()
      .unstack(fill_value=0)
)

# garantir ordem das colunas
resumo = resumo.reindex(
    columns=["ATÉ 45 MIN", "46 MIN até 1H", "> 1h"],
    fill_value=0
)

resumo = (
    resumo
      .reset_index()
      .rename(columns={"dia": "DAY"})
      .sort_values("DAY")
)

# --- 5. Gravar resultados em Excel ---
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    # 5.1 Aba original (sem colunas auxiliares)
    (
        df
          .drop(columns=["Dt.Entrada_dt", "Tempo_td", "dia", "categoria"])
          .to_excel(writer, sheet_name=sheet_in, index=False)
    )

    # 5.2 Planilha2 formatada com DAY e as três faixas
    resumo.to_excel(writer, sheet_name="Planilha2", index=False)

    # (Opcional) detalhamento por categoria, se ainda precisar:
    for cat in ["ATÉ 45 MIN", "46 MIN até 1H", "> 1h"]:
        (
            df[df["categoria"] == cat]
              .drop(columns=["Dt.Entrada_dt", "Tempo_td", "dia", "categoria"])
              .to_excel(
                  writer,
                  sheet_name=f"Detalhe_{cat.replace(' ', '_')}",
                  index=False
              )
        )

print(f"✔️ Processamento concluído. Arquivo salvo em: {output_path}")
