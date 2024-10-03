import streamlit as st
import pymupdf
import fitz
import pandas as pd
from decimal import Decimal
from io import BytesIO


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
def redact_text_on_page(page, df, page_number, round_option,  new_width, new_width_2, new_height, new_height_2):
    """
    Заменяет текст на указанной странице документа.


    :param page: Объект страницы документа
    :param df: DataFrame с данными для замены
    :param page_number: Номер страницы для обработки
    :param new_fontsize: Новый размер шрифта
    :param new_width: Новая ширина прямоугольника
    :param new_width_2: Новая ширина прямоугольника
    :param new_height: Новая высота прямоугольника
    :param new_height_2: Новая высота прямоугольника
    """
    # Фильтруем данные для текущей страницы
    
    df_page = df[df['Page'] == page_number]
    
    for i, raw_text in enumerate(df_page['Old Value'].values):
        try:
            new_text = str(Decimal(df_page['New Value'].values[i]).quantize(Decimal(round_option)))
            if new_text == 'NaN':
                new_text = ''
            raw_text = str(Decimal(df_page['Old Value'].values[i]).quantize(Decimal(round_option)))
            hits = page.search_for(raw_text)
        except:
            new_text = str(df_page['New Value'].values[i])
            raw_text = str(df_page['Old Value'].values[i])
            hits = page.search_for(raw_text)
        # Добавляем аннотацию для редактирования
        for rect in hits:
            x1, y1, x2, y2 = rect
            new_x1 = x1 + new_width_2
            new_x2 = x2 + new_width
            new_y2 = y2 + new_height_2
            new_y1 = y1 + new_height
            new_rect = fitz.Rect(new_x1,new_y1, new_x2, new_y2)
            
            #Добавить прямоугольник
            page.add_rect_annot(new_rect)

            # Добавить текст без фона, но с рамочкой
            page.add_freetext_annot(new_rect, new_text,
                                align=fitz.TEXT_ALIGN_RIGHT, border_color=(1,1,1),fontsize=7.85, fill_color=(1,1,1))


        # Применяем редактирование
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
        data_to_change = pd.read_excel(excel_file)
        unique_page_numbers = data_to_change['Page'].unique()
        round_option = col1.selectbox("Округлить значения", ['1.','0.1','0.01','0.001','0.0001','0.00001'], index=3)
        col1.info("Необходимо выбрать ту же точность, как и в оригинальном PDF")
        new_width = col1.number_input("Подвинуть вправо(+)/влево(-)", value=0.00)
        new_width_2 = 0.00
        new_height = col1.number_input("Подвинуть вниз(+)/вверх(-)", value=0.85)
        new_height_2 = 0.00
        # Редактирование текста
        for page_index in range(len(doc)):
            page = doc[page_index]

            redact_text_on_page(page, data_to_change, page_index, round_option, new_width, new_width_2, new_height, new_height_2)

            # Вывод редактированной страницы
            if page_index in unique_page_numbers:
                pix = page.get_pixmap(dpi=300)
                image_data = pix.tobytes()
                col2.image(image_data, caption=f'Страница {page_index + 1}')
                
        # BytesIO
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)


        # Отображение результата
        col1.download_button("Скачать отредактированный файл", buffer, file_name="updated_document.pdf")
        col1.success("Файл успешно отредактирован!")

if __name__ == "__main__":
    main()