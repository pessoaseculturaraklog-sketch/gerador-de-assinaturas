import streamlit as st
import os
import zipfile
import io
import re

def format_phone_number(telefone_str):
    numeros = re.sub(r'\D', '', telefone_str)
    
    if len(numeros) < 10:
        return telefone_str

    ddd = numeros[:2]
    numero = numeros[2:]
    
    if len(numero) == 8:
        return f"({ddd}) {numero[:4]}-{numero[4:]}"
    elif len(numero) == 9:
        return f"({ddd}) {numero[:5]}-{numero[5:]}"
    else:
        return f"({ddd}) {numero}"

def create_zip_package(nome, cargo, telefone):
    telefone_formatado = format_phone_number(telefone)
    nome_arquivo_sem_extensao = f"Instalar_Assinatura_{nome.replace(' ', '_')}"

    try:
        with open("import.ps1", "r", encoding="utf-8") as f:
            script_content = f.read()
    except FileNotFoundError:
        st.error("ERRO: O arquivo 'import.ps1' não foi encontrado no mesmo diretório do app.")
        return None

    # --- 2. Personalizar o script com os dados do usuário ---
    script_content = script_content.replace("NOME_AQUI", nome)
    script_content = script_content.replace("CARGO_AQUI", cargo)
    script_content = script_content.replace("TELEFONE_AQUI", telefone_formatado)
    script_content = re.sub(
        r'\$networkPath\s*=\s*".*?"', 
        r'$networkPath = Join-Path -Path $PSScriptRoot -ChildPath "arquivos_de_assinatura"', 
        script_content
    )

    batch_content = f"""@echo off
chcp 65001 > nul
echo.
echo    ================================================
echo      INSTALADOR DE ASSINATURA - {nome}
echo    ================================================
echo.
echo    Este script ira configurar sua assinatura de email no Outlook.
echo    Pressione qualquer tecla para comecar...
pause > nul
REM Executa o script PowerShell
powershell.exe -ExecutionPolicy Bypass -File "%~dp0{nome_arquivo_sem_extensao}.ps1"
echo.
echo    Processo concluido.
pause
"""
    
    # --- 4. Criar o arquivo ZIP em memória ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"{nome_arquivo_sem_extensao}.ps1", script_content.encode("utf-8"))
        
        zip_file.writestr(f"{nome_arquivo_sem_extensao}.bat", batch_content.encode("cp1252", errors="replace"))

        source_dir = "arquivos_de_assinatura"
        if not os.path.isdir(source_dir):
            st.error(f"ERRO: A pasta '{source_dir}' não foi encontrada.")
            st.info("Certifique-se que a pasta com os modelos de assinatura (.htm, .rtf etc) está no mesmo local do app.")
            return None
            
        for root, _, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                archive_name = os.path.relpath(file_path)
                zip_file.write(file_path, archive_name)

    st.success("Pacote de instalação gerado com sucesso!")
    return zip_buffer.getvalue()


# --- Interface do Streamlit ---

st.set_page_config(page_title="Gerador de Assinatura", layout="centered")

st.title("Gerador de Assinatura de E-mail")
st.write("Preencha seus dados abaixo para gerar o instalador da sua assinatura personalizada para o Outlook.")

with st.form(key="signature_form"):
    nome_completo = st.text_input("Nome Completo")
    cargo_usuario = st.text_input("Seu Cargo (Ex: Analista de Suporte)")
    telefone_usuario = st.text_input("Telefone com DDD (Ex: 19912345678)")

    submit_button = st.form_submit_button(label="✔️ Gerar Instalador")

if submit_button:
    if not all([nome_completo, cargo_usuario, telefone_usuario]):
        st.warning("Por favor, preencha todos os campos.")
    else:
        zip_data = create_zip_package(nome_completo, cargo_usuario, telefone_usuario)
        
        if zip_data:
            st.markdown("""
                ---
                ### Pronto! Agora siga os passos abaixo:
                
                1.  Clique no botão para baixar o arquivo `.zip`.
                2.  **Extraia todo o conteúdo** do arquivo para uma nova pasta no seu computador.
                3.  Dentro da pasta que você extraiu, clique duas vezes no arquivo **.bat** (ele terá um ícone de engrenagens).
                4.  Siga as instruções na tela que aparecer.
            """)

            st.download_button(
                label="📥 Baixar Pacote de Instalação",
                data=zip_data,
                file_name=f"Instalador_Assinatura_{nome_completo.replace(' ', '_')}.zip",
                mime="application/zip"
            )
