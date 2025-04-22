import pandas as pd
import os
import sys

try:
    import xlrd
except ImportError:
    print("❌ A biblioteca 'xlrd' não está instalada. Execute: pip install xlrd==2.0.1")
    sys.exit(1)

def converter_xls_para_xlsx(caminho_entrada, caminho_saida):
    xls_file = pd.read_excel(caminho_entrada, sheet_name=None, engine="xlrd")
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        for sheet_name, sheet_df in xls_file.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ Arquivo convertido com sucesso: {caminho_saida}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("❌ Uso correto: python convert.py entrada.xls saida.xlsx")
        sys.exit(1)

    entrada = sys.argv[1]
    saida = sys.argv[2]

    if not os.path.exists(entrada):
        print(f"❌ Arquivo não encontrado: {entrada}")
        sys.exit(1)

    converter_xls_para_xlsx(entrada, saida)
