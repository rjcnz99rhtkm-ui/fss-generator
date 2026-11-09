import os
import re
import io
import zipfile
import subprocess
import pdfplumber
import streamlit as st
from docxtpl import DocxTemplate

# --- 1. ЛОГИКА МГНОВЕННОГО РАСПОЗНАВАНИЯ ДАННЫХ ---
def extract_data_from_pdf(pdf_stream):
    extracted_data = {
        'blank_series': 'не найдено',
        'certificate_num': 'не найдено',
        'issue_day': 'не найдено',
        'issue_month': 'не найдено',
        'issue_year': '2026',
        'product_name': 'не найдено',
        'quantity': 'не найдено',
        'destination': 'China, Manzhouli',
        'transport_num': 'не найдено'
    }
    
    PRODUCT_DICTIONARY = {
        'MILLET': 'SETARIA ITALICA (L.) BEAUV / Пшено (КОД ТН ВЭД 1103199009)',
        'RAPE SEEDS': 'BRASSICA NAPUS L. / Семена рапса (КОД ТН ВЭД 1205109000)',
        'OATS': 'AVENA SATIVA L. / Овес (КОД ТН ВЭД 1004900000)',
        'OAT': 'AVENA SATIVA L. / Овес (КОД ТН ВЭД 1004900000)',
        'WHEAT': 'TRITICUM AESTIVUM L. / Пшеница мягкая (КОД ТН ВЭД 1001990000)',
        'FLAX SEEDS': 'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
        'LINSEED': 'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
        'BARLEY': 'HORDEUM VULGARE L. / Ячмень (КОД ТН ВЭД 1003900000)'
    }
    
    months_dict = {
        'JAN': 'Января', 'FEB': 'Февраля', 'MAR': 'Марта', 'APR': 'Апреля',
        'MAY': 'Мая', 'JUN': 'Июня', 'JUL': 'Июля', 'AUG': 'Августа',
        'SEP': 'Сентября', 'OCT': 'Октября', 'NOV': 'Ноября', 'DEC': 'Декабря'
    }

    with pdfplumber.open(pdf_stream) as pdf:
        # Исправлено: берем строго первую страницу [0]
        page1 = pdf.pages[0]
        text1 = page1.extract_text() or ""
        
        series_match = re.search(r'\b([A-Z]\d{7})\b', text1)
        if series_match:
            extracted_data['blank_series'] = series_match.group(1)

        num_match = re.search(r'№\s*(\d+)', text1)
        if num_match:
            extracted_data['certificate_num'] = num_match.group(1)

        date_match = re.search(r'(\d{2})\s+([A-Z]{3})\s+(\d{4})', text1)
        if date_match:
            extracted_data['issue_day'] = date_match.group(1)
            eng_month = date_match.group(2)
            extracted_data['issue_month'] = months_dict.get(eng_month, eng_month)
            extracted_data['issue_year'] = date_match.group(3)

        detected_culture = None
        for culture_keyword in PRODUCT_DICTIONARY.keys():
            if culture_keyword in text1.upper():
                detected_culture = culture_keyword
                extracted_data['product_name'] = PRODUCT_DICTIONARY[culture_keyword]
                break
        
        weight_match = re.search(r'(\d+\.?\d*)\s*(?:T\.|KG|PC)', text1, re.IGNORECASE)
        if weight_match:
            raw_weight = weight_match.group(1).replace('.', ',')
            extracted_data['quantity'] = f"{raw_weight} тонн"
        
        if not detected_culture:
            extracted_data['product_name'] = "Не удалось распознать название культуры"

        transport_match = re.search(r'(?:CARRIAGE|TRANSPORT|CONTAINER)\s+([A-Z0-9]+)', text1, re.IGNORECASE)
        if transport_match:
            extracted_data['transport_num'] = transport_match.group(1)

        if len(pdf.pages) > 1:
            # Исправлено: берем вторую страницу [1] для приложений
            page2 = pdf.pages[1]
            text2 = page2.extract_text() or ""
            all_wagons = re.findall(r'\b(\d{8})\b', text2)
            if all_wagons:
                extracted_data['transport_num'] = ", ".join(sorted(list(set(all_wagons))))

    return extracted_data

def convert_docx_to_pdf_server(docx_bytes, filename_base):
    """Скрытая конвертация DOCX в PDF через LibreOffice CLI."""
    temp_docx = f"temp_{filename_base}.docx"
    temp_pdf = f"temp_{filename_base}.pdf"
    
    with open(temp_docx, "wb") as f:
        f.write(docx_bytes)
        
    try:
        subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf', temp_docx
        ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15)
        
        if os.path.exists(temp_pdf):
            with open(temp_pdf, "rb") as f:
                pdf_bytes = f.read()
            os.remove(temp_pdf)
            os.remove(temp_docx)
            return pdf_bytes
    except Exception:
        pass
    
    if os.path.exists(temp_docx):
        os.remove(temp_docx)
    return None

# --- 2. ИНТЕРФЕЙС STREAMLIT ---
st.set_page_config(page_title="Переоформление ФСС", page_icon="🌾", layout="centered")

st.title("🌾 Автогенерация заявлений ФСС (DOCX + PDF)")
st.write("Сайт мгновенно обработает файлы и подготовит архивы.")

TEMPLATE_FILE = "template.docx"
if not os.path.exists(TEMPLATE_FILE):
    st.error(f"❌ Файл шаблона '{TEMPLATE_FILE}' не найден.")
else:
    uploaded_files = st.file_uploader(
        "Перетащите сюда файлы PDF (допускается массовая загрузка)", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        status_text = st.empty()
        status_text.info(f"🚀 Создание документов...")
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            success_count = 0
            
            for uploaded_file in uploaded_files:
                file_name_clean = os.path.splitext(uploaded_file.name)[0]
                
                try:
                    pdf_io = io.BytesIO(uploaded_file.read())
                    data = extract_data_from_pdf(pdf_io)
                    
                    doc = DocxTemplate(TEMPLATE_FILE)
                    doc.render(data)
                    
                    docx_io = io.BytesIO()
                    doc.save(docx_io)
                    docx_bytes = docx_io.getvalue()
                    
                    # Добавляем Word-файл в архив
                    zip_file.writestr(f"Заявление_{file_name_clean}.docx", docx_bytes)
                    
                    # Пробуем сгенерировать PDF скрытым конвертером
                    pdf_bytes = convert_docx_to_pdf_server(docx_bytes, file_name_clean)
                    if pdf_bytes:
                        zip_file.writestr(f"Заявление_{file_name_clean}.pdf", pdf_bytes)
                        
                    success_count += 1
                except Exception as e:
                    st.error(f"⚠️ Ошибка в файле {uploaded_file.name}: {e}")
            
        status_text.empty()
        
        if success_count > 0:
            st.success(f"🔥 Готово! Обработано файлов: {success_count}.")
            st.download_button(
                label="📥 Скачать архив со всеми заявлениями (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Сгенерированные_Заявления.zip",
                mime="application/zip",
                use_container_width=True
            )
