import os
import re
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
from PyPDF2 import PdfReader, PdfWriter

# Configuração do caminho do executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Configuração do caminho do Poppler
poppler_path = r'C:\Program Files\Poppler\Library\bin'  # Atualize para o caminho correto

# Função para extrair texto de uma lista de imagens
def extract_text_from_images(images):
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang='por')  # Define o idioma como português
    return text

def extract_name_and_obra(text):
    nome_patterns = [
        re.compile(r'Nome Funcionario:\s*([^\n=]+)', re.IGNORECASE),
        re.compile(r'Nome Funcionario:\s\n*([^\n]+)', re.IGNORECASE),
        re.compile(r'Nome Funcionario:\s*=\s*([^\n]+)', re.IGNORECASE)
    ]
    obra_patterns = [
        re.compile(r'Obra:\s*([0-9]+\.[0-9]+)', re.IGNORECASE),
        re.compile(r'Obra:\s*\+\s*([0-9]+\.[0-9]+)', re.IGNORECASE | re.DOTALL),
        re.compile(r'Obra: É ([0-9]+\.[0-9]+)', re.IGNORECASE | re.DOTALL)
    ]
    nome_completo = None
    obra = None
    # Tente encontrar uma correspondência para o nome
    for pattern in nome_patterns:
        match = pattern.search(text)
        if match:
            nome_completo = match.group(1).strip()
            break
    # Tente encontrar uma correspondência para a obra
    for pattern in obra_patterns:
        match = pattern.search(text)
        if match:
            obra = match.group(1).strip()
            break
    # Verifique se encontramos tanto o nome quanto a obra
    if nome_completo and obra:
        return nome_completo, obra
    return None, None

# Função para salvar a segunda página do PDF
def save_second_page(pdf_path, output_path):
    with open(pdf_path, 'rb') as infile:
        reader = PdfReader(infile)
        tamanho= len(reader.pages) 
        if len(reader.pages) > 1:
            writer = PdfWriter()
            writer.add_page(reader.pages[tamanho-1])  # Segunda página é índice 1
            with open(output_path, 'wb') as outfile:
                writer.write(outfile)

# Caminho para os arquivos PDF
pdf_folder_path = r'\\pasta_para_arquivos'

# Processar cada PDF na pasta
for pdf_file in os.listdir(pdf_folder_path):
    if pdf_file.endswith('.pdf'):
        pdf_path = os.path.join(pdf_folder_path, pdf_file)
        # Converter a primeira página do PDF em imagem
        images = convert_from_path(pdf_path, first_page=1, last_page=1, poppler_path=poppler_path)
        # Extrair texto da imagem
        text = extract_text_from_images(images)
        # Identificar o nome completo e a obra
        nome_completo, obra = extract_name_and_obra(text)
        if nome_completo and obra:
            # Criar caminho para salvar a segunda página do PDF na pasta do centro de custo (obra)
            output_folder = os.path.join(pdf_folder_path, obra)
            os.makedirs(output_folder, exist_ok=True)  # Cria a pasta se não existir
            nome_completo_sanitized = re.sub(r'[^\w\s]', '', nome_completo).replace(' ', '_')
            output_path = os.path.join(output_folder, f'ENCONTRADO_{nome_completo_sanitized}.pdf')
            # Salvar a segunda página do PDF
            save_second_page(pdf_path, output_path)
            print(f'Arquivo: {pdf_file}')
            print(f'Nome Completo: {nome_completo}')
            print(f'Obra: {obra}')
            print(f'Segunda página salva em: {output_path}')
            print('-' * 40)
        else:
            print(f'Nome ou obra não encontrados no PDF {pdf_file}.')
            print('-' * 40)
