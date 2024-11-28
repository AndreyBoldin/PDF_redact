import streamlit as st
import pymupdf
import fitz
import pandas as pd
from decimal import Decimal
from io import BytesIO
import PyPDF2
from pdf2image import convert_from_path, convert_from_bytes
import img2pdf
from PIL import Image
import os

hide_label = """
<style>
[data-testid="stFileUploader"] {
    margin-top: -20px;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    color: white;
    padding: 0px;
}
}
div[data-testid="stFileUploaderDropzoneInstructions"] div{
    display: flex;
}
div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]>button[data-testid="baseButton-secondary"] {
    color: #131720;
    width: 20%;
    height: 100%;
    padding: 0px;
}
div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]>button[data-testid="baseButton-secondary"]:hover {
    color: #131720;
}
div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]>button[data-testid="baseButton-secondary"]:focus {
    color: #131720;
}
div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]>button[data-testid="baseButton-secondary"]:active {
    color: #5854f4;
}
div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]> button[data-testid="baseButton-secondary"]::after {
        content: "Загрузить файл";
        color: white;
        hover: blue;
        display: block;
        position: absolute;
    }
    
[data-testid="stFileUploaderDropzoneInstructions"] div::before {color:white; font-size: 0.9em; content:"Загрузите или перетяните файлы сюда"}
[data-testid="stFileUploaderDropzoneInstructions"] div span{display:none;}
[data-testid="stFileUploaderDropzoneInstructions"] div::after {color:white; font-size: .8em; content:"Лимит 200MB на файл"}
[data-testid="stFileUploaderDropzoneInstructions"] div small{display:none;}
[data-testid="stFileUploaderDropzoneInstructions"] button{display:flex;width: 30%; padding: 0px;}
"""


# Функция для редактирования текста в PDF
def redact_text_on_page(page, df, page_number, new_width, new_width_2, new_height, new_height_2):
    df_page = df[df['Page'] == page_number]

    for i, raw_text in enumerate(df_page['Old Value'].values):
        try:
            new_text = str(df_page['New Value'].values[i])
            raw_text = str(df_page['Old Value'].values[i])

            #Расчет знаков после запятой
            if '.' in raw_text:
                decimals = len(raw_text.split('.')[1])
            else:
                decimals = 0

            if decimals == 0:
                round_option = '1.'
            else:
                round_option = '.' + '0' * (decimals - 1) + '1'

            new_text = str(Decimal(new_text).quantize(Decimal(round_option)))
            if new_text == 'NaN':
                new_text = ''

            hits = page.search_for(raw_text)
        except:
            new_text = str(df_page['New Value'].values[i])
            raw_text = str(df_page['Old Value'].values[i])
            hits = page.search_for(raw_text)

        # Координаты нового прямоугольника
        for rect in hits:
            x1, y1, x2, y2 = rect
            new_x1 = x1 + new_width_2 - 15
            new_x2 = x2 + new_width
            new_y2 = y2 + new_height_2
            new_y1 = y1 + new_height
            new_rect = fitz.Rect(new_x1, new_y1, new_x2, new_y2)

            # Редактирование поверх старого
            page.add_redact_annot(new_rect, '')
            page.apply_redactions(text=1)

            page.add_freetext_annot(new_rect, new_text,
                                    align=fitz.TEXT_ALIGN_RIGHT,fontname=page.get_fonts()[1][4], border_color=(1, 1, 1), fontsize=7.85, fill_color=(1, 1, 1))
            
        page.apply_redactions()

# Основная функция приложения
def main():
    
    st.set_page_config(layout="wide")
    st.title("Редактирование текста в PDF")
    st.markdown(hide_label, unsafe_allow_html=True)
    # Загрузка файлов
    pdf_file = st.file_uploader("Загрузите PDF файл", type="pdf")
    
    excel_file = st.file_uploader("Загрузите Excel файл", type="xlsx")
    col1, col2 = st.columns([1,2])
    if pdf_file and excel_file:
        # Чтение файлов
        pdf_buffer = BytesIO(pdf_file.getvalue())
        doc = pymupdf.open(stream=pdf_buffer)
        data_to_change = pd.read_excel(excel_file, dtype=object)
        unique_page_numbers = data_to_change['Page'].unique()
        new_width = col1.number_input("Подвинуть вправо(+)/влево(-)", value=0.00)
        new_width_2 = 0.00
        new_height = col1.number_input("Подвинуть вниз(+)/вверх(-)", value=0.85)
        new_height_2 = 0.00
        # Редактирование текста
        for page_index in range(len(doc)):
            page = doc[page_index]

            redact_text_on_page(page, data_to_change, page_index, new_width, new_width_2, new_height, new_height_2)

            # Вывод редактированной страницы
            if page_index in unique_page_numbers:
                pix = page.get_pixmap(dpi=300)
                image_data = pix.tobytes()
                col2.image(image_data, caption=f'Страница {page_index + 1}')
                
        # BytesIO 
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        #Create dir buffer relative to current directory
        if not os.path.exists("buffer"):
            os.makedirs("buffer")
            
        #save pdf as png file
        for page_index in range(len(doc)):
            page = doc[page_index]
            pix = page.get_pixmap(dpi=300)
            image_data = pix.tobytes()
            with open(f"buffer/page_{page_index + 1}.png", "wb") as f:
                f.write(image_data)
        
        #Merge images to pdf using Image on individual pages
        images = [Image.open(f"buffer/page_{i + 1}.png") for i in range(len(doc))]
        images[0].save("buffer/intermediate_document.pdf", save_all=True, append_images=images[1:])
        file = open("buffer/intermediate_document.pdf", "rb")

        # Отображение результата
        col1.download_button("Нередактируемый PDF", file, file_name="not_editable_document.pdf")

        col1.download_button("Редактируемый, но с комментариями!", buffer, file_name="ediatble_document.pdf")
        col1.success("Файл успешно отредактирован!")

if __name__ == "__main__":
    main()