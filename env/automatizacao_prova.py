import re
import PyPDF2
import fitz
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


def extrair_e_adicionar_imagens(pdf_file, document):
    pdf = fitz.open(pdf_file)
    for page_number in range(len(pdf)):
        for img_index, img in enumerate(pdf.get_page_images(page_number)):
            xref = img[0]
            base_image = pdf.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image_filename = f"image{page_number + 1}_{img_index}.{image_ext}"
            with open(image_filename, "wb") as image_file:
                image_file.write(image_bytes)
            # Adiciona a imagem ao documento do Word
            document.add_picture(image_filename)
    pdf.close()

def extrair_questoes(pdf_file):
    questoes = []
    questao_regex = re.compile(r'QUESTÃO \d+')
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        texto_atual = ''
        for page in pdf_reader.pages:
            texto_pagina = page.extract_text() or ''
            for linha in texto_pagina.split('\n'):
                if questao_regex.search(linha):
                 if texto_atual:
                    questoes.append(texto_atual)
                    texto_atual = ''
                    texto_atual += linha + '\n' # Se já temos texto acumulado, significa que é o fim de uma questão.
                else:
                    texto_atual += linha + '\n'     # Adiciona a última questão da página ao finalizar o loop.
            
            if texto_atual:
                questoes.append(texto_atual)
                texto_atual = ''
            
    return questoes





def criar_documento_word(questoes, output_file):
    # Inicializa o documento Word
    document = Document()

    # Cria e configura os estilos aqui
    estilo_questao = document.styles.add_style('EstiloQuestao', WD_STYLE_TYPE.PARAGRAPH)
    estilo_questao.font.name = 'Calibri'
    estilo_questao.font.size = Pt(11)
    estilo_questao.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # ... (Configure outros estilos conforme necessário)

    # Adiciona as questões ao documento Word aplicando os estilos
    for questao in questoes:
        p = document.add_paragraph(questao)
        p.style = document.styles['EstiloQuestao']

    # Salva o documento Word
    document.save(output_file)

if __name__ == "__main__":
    pdf_file = 'Histologia.pdf'  # Substitua 'exemplo.pdf' pelo nome do seu arquivo PDF
    output_file = 'prova_formatada.docx'

    questoes = extrair_questoes(pdf_file)
    criar_documento_word(questoes, output_file)
    print('Prova Formatada com Sucesso!')
