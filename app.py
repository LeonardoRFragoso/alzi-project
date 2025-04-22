import pandas as pd
from datetime import timedelta

# --- CONFIGURAÇÕES ---
input_path  = "TC.xlsx"
output_path = "TC_output.xlsx"
sheet_in    = "Recuperada_Planilha1"
threshold   = timedelta(minutes=45)

# --- 1. Carregar dados e normalizar cabeçalhos ---
df = pd.read_excel(
    input_path,
    sheet_name=sheet_in,
    header=14,             # <-- linha com nomes das colunas reais
    dtype=str,
    engine="openpyxl"
)

# 1.1 Limpar quebras de linha e espaços em branco nos nomes de coluna
df.columns = (
    df.columns
      .str.replace("\n", " ", regex=False)  # quebras de linha
      .str.strip()                          # espaços laterais
)

print("Colunas disponíveis:", df.columns.tolist())

# --- 2. Identificar dinamicamente as colunas de data e tempo ---
def find_col(key):
    for c in df.columns:
        if key.lower() in c.lower().replace(" ", ""):
            return c
    return None

col_dt     = find_col("Dt.Entrada")
col_tempo  = find_col("Tempo")

if col_dt is None or col_tempo is None:
    raise KeyError(
        f"Não encontrei as colunas esperadas.\n"
        f"Procurei por algo como 'Dt.Entrada' e 'Tempo' em: {df.columns.tolist()}"
    )

# Limpeza de espaços extras
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

# --- 4. Filtragem e agrupamento ---
df["data"] = df["Dt.Entrada_dt"].dt.date
df["dia"]  = df["Dt.Entrada_dt"].dt.day

df_acima = df[df["Tempo_td"] > threshold].copy()
df_ate   = df[df["Tempo_td"] <= threshold].copy()

# Garantir que as colunas estejam presentes
for subset in [df_acima, df_ate]:
    subset["dia"] = subset["Dt.Entrada_dt"].dt.day

# Agrupamentos por dia
ontime_por_dia = df_ate.groupby("dia", as_index=True).size().rename("ON TIME")
delay_por_dia  = df_acima.groupby("dia", as_index=True).size().rename("DELAY")

# Juntar os dois resultados
resumo = pd.concat([ontime_por_dia, delay_por_dia], axis=1)
resumo = resumo.fillna(0).astype(int)
resumo.index.name = None
resumo = resumo.reset_index(names=["DAY"])

# Ordenar por dia
resumo = resumo.sort_values("DAY")

# --- 5. Gravar resultados em Excel ---
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    # 5.1 Aba original (sem colunas auxiliares)
    (
        df.drop(columns=["Dt.Entrada_dt", "Tempo_td", "data", "dia"])
          .to_excel(writer, sheet_name=sheet_in, index=False)
    )

    # 5.2 Nova Planilha2 formatada com DAY, ON TIME, DELAY
    resumo.to_excel(writer, sheet_name="Planilha2", index=False)

    # 5.3 Detalhes acima de 45min
    (
        df_acima.drop(columns=["Dt.Entrada_dt", "Tempo_td", "data", "dia"])
                .to_excel(writer, sheet_name="Detalhe_Acima_45min", index=False)
    )

    # 5.4 Detalhes até 45min
    (
        df_ate.drop(columns=["Dt.Entrada_dt", "Tempo_td", "data", "dia"])
              .to_excel(writer, sheet_name="Detalhe_Ate_45min", index=False)
    )

print(f"✔️ Processamento concluído. Arquivo salvo em:\n   {output_path}")