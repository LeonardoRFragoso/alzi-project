import streamlit as st
import os
import tempfile
import subprocess
import shutil

# Título
st.set_page_config(page_title="Conversão e Análise de Planilha", layout="centered")
st.title("🧮 Processador de Planilha de Contêineres (.xls → .xlsx)")

# Upload
uploaded_file = st.file_uploader("📎 Faça o upload da planilha (.xls)", type=["xls"])

if uploaded_file:
    st.success("Arquivo recebido com sucesso!")

    with tempfile.TemporaryDirectory() as tmpdir:
        input_xls_path = os.path.join(tmpdir, "entrada.xls")
        converted_xlsx_path = os.path.join(tmpdir, "TC.xlsx")
        output_path = os.path.join(tmpdir, "TC_output.xlsx")

        # Salvar temporariamente o arquivo enviado
        with open(input_xls_path, "wb") as f:
            f.write(uploaded_file.read())

        # Converter .xls para .xlsx
        st.info("🔄 Convertendo arquivo para .xlsx...")
        try:
            subprocess.run(
                [
                    r"C:/Users/leonardo.fragoso/Desktop/Projetos/alzi-project/venv/Scripts/python.exe",
                    "C:/Users/leonardo.fragoso/Desktop/Projetos/alzi-project/convert.py",
                    input_xls_path,
                    converted_xlsx_path
                ],
                check=True
            )
            st.success("✅ Conversão concluída.")
        except subprocess.CalledProcessError:
            st.error("❌ Erro ao converter o arquivo. Verifique o script convert.py.")
            st.stop()

        # Copiar o arquivo convertido para local esperado por app.py
        shutil.copy(converted_xlsx_path, "C:/Users/leonardo.fragoso/Desktop/Projetos/alzi-project/TC.xlsx")

        # Rodar o script app.py para gerar a planilha output
        st.info("⚙️ Processando dados com app.py...")
        try:
            subprocess.run(
                [
                    r"C:/Users/leonardo.fragoso/Desktop/Projetos/alzi-project/venv/Scripts/python.exe",
                    "C:/Users/leonardo.fragoso/Desktop/Projetos/alzi-project/app.py"
                ],
                check=True
            )
            st.success("📊 Planilha processada com sucesso!")
        except subprocess.CalledProcessError:
            st.error("❌ Erro durante o processamento com app.py.")
            st.stop()

        # Copiar resultado para a pasta temporária e exibir para download
        final_output_path = "C:/Users/leonardo.fragoso/Desktop/Projetos/alzi-project/TC_output.xlsx"
        shutil.copy(final_output_path, output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Baixar Planilha Final (.xlsx)",
                data=f,
                file_name="TC_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
