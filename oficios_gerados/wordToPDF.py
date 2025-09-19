import os
from docx2pdf import convert

# Caminho da pasta com os arquivos .docx
pasta = os.getcwd()  # usa a pasta atual, mas pode ser substituído por outro caminho

# Lista todos os arquivos na pasta
arquivos = os.listdir(pasta)

# Filtra os arquivos .docx (excluindo os temporários iniciados com ~)
docx_files = [f for f in arquivos if f.endswith(".docx") and not f.startswith("~")]

# Converte cada arquivo .docx para .pdf
for docx in docx_files:
    caminho_completo = os.path.join(pasta, docx)
    try:
        convert(caminho_completo)
        print(f"Convertido: {docx}")
    except Exception as e:
        print(f"Erro ao converter {docx}: {e}")
