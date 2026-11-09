# -*- coding: utf-8 -*-
"""
app.py — автогенерация заявлений ФСС (DOCX + PDF).
Файл-компаньон: template.docx (Jinja2-шаблон).
"""
import io
import os
import re
import shutil
import subprocess
import tempfile
import uuid
import zipfile

import pdfplumber
import streamlit as st
from docxtpl import DocxTemplate


# ============ СПРАВОЧНИКИ ============

# Ботаническое название -> описание для заявления (самый надёжный ключ)
BOTANICAL_TO_PRODUCT = {
    'LINUM USITATISSIMUM': 'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
    'BRASSICA NAPUS':      'BRASSICA NAPUS L. / Семена рапса (КОД ТН ВЭД 1205109000)',
    'SETARIA ITALICA':     'SETARIA ITALICA (L.) BEAUV / Пшено (КОД ТН ВЭД 1103199009)',
    'AVENA SATIVA':        'AVENA SATIVA L. / Овес (КОД ТН ВЭД 1004900000)',
    'TRITICUM AESTIVUM':   'TRITICUM AESTIVUM L. / Пшеница мягкая (КОД ТН ВЭД 1001990000)',
    'HORDEUM VULGARE':     'HORDEUM VULGARE L. / Ячмень (КОД ТН ВЭД 1003900000)',
}

# Common name -> описание (фолбэк). Учтена опечатка LINSED в реальных сертификатах.
COMMON_TO_PRODUCT = {
    'FLAX SEEDS': 'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
    'LINSEED':    'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
    'LINSED':     'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
    'RAPE SEEDS': 'BRASSICA NAPUS L. / Семена рапса (КОД ТН ВЭД 1205109000)',
    'MILLET':     'SETARIA ITALICA (L.) BEAUV / Пшено (КОД ТН ВЭД 1103199009)',
    'OATS':       'AVENA SATIVA L. / Овес (КОД ТН ВЭД 1004900000)',
    'OAT':        'AVENA SATIVA L. / Овес (КОД ТН ВЭД 1004900000)',
    'WHEAT':      'TRITICUM AESTIVUM L. / Пшеница мягкая (КОД ТН ВЭД 1001990000)',
    'BARLEY':     'HORDEUM VULGARE L. / Ячмень (КОД ТН ВЭД 1003900000)',
}

MONTHS_DICT = {
    'JAN': 'Января', 'FEB': 'Февраля', 'MAR': 'Марта', 'APR': 'Апреля',
    'MAY': 'Мая', 'JUN': 'Июня', 'JUL': 'Июля', 'AUG': 'Августа',
    'SEP': 'Сентября', 'OCT': 'Октября', 'NOV': 'Ноября', 'DEC': 'Декабря',
    'JANUARY': 'Января', 'FEBRUARY': 'Февраля', 'MARCH': 'Марта', 'APRIL': 'Апреля',
    'JUNE': 'Июня', 'JULY': 'Июля', 'AUGUST': 'Августа', 'SEPTEMBER': 'Сентября',
    'OCTOBER': 'Октября', 'NOVEMBER': 'Ноября', 'DECEMBER': 'Декабря',
    'ЯНВ': 'Января', 'ФЕВ': 'Февраля', 'МАР': 'Марта', 'АПР': 'Апреля',
    'МАЙ': 'Мая', 'ИЮН': 'Июня', 'ИЮЛ': 'Июля', 'АВГ': 'Августа',
    'СЕН': 'Сентября', 'ОКТ': 'Октября', 'НОЯ': 'Ноября', 'ДЕК': 'Декабря',
}

DEFAULT_DESTINATION = 'China, Manzhouli'


# ============ ХЕЛПЕРЫ ============

def sanitize_filename(name: str, fallback: str = "document") -> str:
    base = os.path.basename(name or "")
    base = os.path.splitext(base)[0]
    base = re.sub(r'[^\w\-]+', '_', base, flags=re.UNICODE).strip('._-')
    return base or fallback


def _extract_weight(text: str) -> str:
    """
    Ищет вес в формате: '66.3 T.', '25.956 T.', '1000 kg', '63,0 t', '1000 т'.
    Ключевой момент: после 'T.' идёт пробел, поэтому стандартный '\\b'
    не работает — используем (?!\\w).
    """
    pattern = re.compile(
        r'(\d[\d\s.,]*)\s*'
        r'(TONNES?|TONS?|MT|KGS?|KG|PCS?|PC|ТОНН[А-ЯЁ]*|КГ|ШТ|T|Т)'
        r'\.?'
        r'(?!\w)',
        re.IGNORECASE | re.UNICODE,
    )
    match = pattern.search(text)
    if not match:
        return "не найдено"

    raw = re.sub(r'\s+', '', match.group(1)).strip('.,')
    unit = match.group(2).upper().rstrip('.')

    try:
        if ',' in raw and '.' in raw:
            if raw.rfind(',') > raw.rfind('.'):
                value = float(raw.replace('.', '').replace(',', '.'))
            else:
                value = float(raw.replace(',', ''))
        elif ',' in raw:
            value = float(raw.replace(',', '.'))
        else:
            value = float(raw)
    except ValueError:
        return "не найдено"

    if unit in ('PC', 'PCS', 'ШТ'):
        return f"{int(value)} шт"
    if unit in ('KG', 'KGS', 'КГ'):
        value = value / 1000.0

    formatted = f"{value:.3f}".rstrip('0').rstrip('.').replace('.', ',')
    return f"{formatted} тонн"


def _extract_issue_date(text: str):
    """
    Возвращает (day, month_ru, year) или ('не найдено',)*3.
    Пропускает диапазоны вида '26 MAY 2026 - 31 MAY 2026' (поле 18),
    чтобы не путать их с датой выдачи (поле 20).
    """
    pattern = re.compile(
        r'(?<![-\u2013\u2014]\s)'
        r'(\d{1,2})\s+([A-ZА-ЯЁ]{3,9})\s+(\d{4})'
        r'(?!\s*[-\u2013\u2014])',
        re.IGNORECASE,
    )
    for m in pattern.finditer(text):
        day, month, year = m.group(1), m.group(2), m.group(3)
        month_ru = MONTHS_DICT.get(month.upper(), month)
        return day.zfill(2), month_ru, year
    return 'не найдено', 'не найдено', 'не найдено'


def _extract_transport(pages_text: list) -> str:
    """Номера вагонов/контейнеров. Сначала по контексту, потом фолбэк."""
    ctx_pattern = re.compile(
        r'(?:CARRIAGE|TRANSPORT|CONTAINER|WAGON|ВАГОН|КОНТЕЙНЕР)'
        r'\s*[№#:/]?\s*([A-Z]{0,4}\d{4,})',
        re.IGNORECASE,
    )
    found = []
    for page_text in pages_text:
        for m in ctx_pattern.finditer(page_text):
            found.append(m.group(1).upper())
    if found:
        return ", ".join(sorted(set(found)))

    if len(pages_text) > 1:
        wagons = []
        for page_text in pages_text[1:]:
            wagons.extend(re.findall(r'\b(\d{8})\b', page_text))
        if wagons:
            return ", ".join(sorted(set(wagons)))

    return "не найдено"


def _extract_destination(text: str) -> str:
    """Значение сразу после 'point of entry' / 'пункт ввоза'."""
    m = re.search(
        r'(?:пункт\s+ввоза|point\s+of\s+entry)\b[^\n]*\n',
        text, re.IGNORECASE,
    )
    if m:
        for line in text[m.end():].split('\n'):
            line = line.strip()
            if line:
                if len(line) < 100:
                    return line
                break
    return DEFAULT_DESTINATION


def _extract_product(full_upper: str) -> str:
    """Сначала ботаническое название, затем common name."""
    for key, value in BOTANICAL_TO_PRODUCT.items():
        if re.search(rf'\b{re.escape(key)}\b', full_upper):
            return value
    for key, value in COMMON_TO_PRODUCT.items():
        if re.search(rf'\b{re.escape(key)}\b', full_upper):
            return value
    return "Не удалось распознать название культуры"


# ============ ПАРСИНГ PDF ============

def extract_data_from_pdf(pdf_stream) -> dict:
    data = {
        'blank_series': 'не найдено',
        'certificate_num': 'не найдено',
        'issue_day': 'не найдено',
        'issue_month': 'не найдено',
        'issue_year': 'не найдено',
        'product_name': 'не найдено',
        'quantity': 'не найдено',
        'destination': DEFAULT_DESTINATION,
        'transport_num': 'не найдено',
    }

    with pdfplumber.open(pdf_stream) as pdf:
        if not pdf.pages:
            raise ValueError("PDF не содержит страниц")

        pages_text = [(page.extract_text() or "") for page in pdf.pages]
        full_text = "\n".join(pages_text)
        full_upper = full_text.upper()
        text1 = pages_text[0]

        # Серия (буква + 7 цифр) — ищем на всех страницах
        m = re.search(r'\b([A-Z]\d{7})\b', full_upper)
        if m:
            data['blank_series'] = m.group(1)

        # Номер сертификата: 18-значный. Не путать с №42/000... из поля 19
        m = re.search(r'[№#]\s*(\d{15,20})', full_text)
        if m:
            data['certificate_num'] = m.group(1)

        # Дата выдачи (поле 20)
        day, month_ru, year = _extract_issue_date(text1)
        data['issue_day'] = day
        data['issue_month'] = month_ru
        data['issue_year'] = year

        # Продукция (ботаническое название приоритетно)
        data['product_name'] = _extract_product(full_upper)

        # Вес
        data['quantity'] = _extract_weight(full_text)

        # Пункт ввоза
        data['destination'] = _extract_destination(text1)

        # Транспорт
        data['transport_num'] = _extract_transport(pages_text)

    return data


# ============ КОНВЕРТАЦИЯ В PDF ============

def _find_soffice():
    for name in ('soffice', 'libreoffice'):
        path = shutil.which(name)
        if path:
            return path
    return None


def convert_docx_to_pdf_server(docx_bytes: bytes):
    soffice = _find_soffice()
    if not soffice:
        return None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "document.docx")
            pdf_path = os.path.join(tmpdir, "document.pdf")
            profile_dir = os.path.join(tmpdir, f"lo_profile_{uuid.uuid4().hex}")

            with open(docx_path, "wb") as f:
                f.write(docx_bytes)

            result = subprocess.run(
                [
                    soffice, '--headless',
                    f'-env:UserInstallation=file://{profile_dir}',
                    '--convert-to', 'pdf',
                    '--outdir', tmpdir,
                    docx_path,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=60,
            )
            if result.returncode == 0 and os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    return f.read()
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        pass
    return None


def _missing_fields(data: dict) -> list:
    missing = []
    if data.get('blank_series') == 'не найдено':
        missing.append('серия бланка')
    if data.get('certificate_num') == 'не найдено':
        missing.append('номер сертификата')
    if data.get('product_name') in ('не найдено', 'Не удалось распознать название культуры'):
        missing.append('название культуры')
    if data.get('quantity') == 'не найдено':
        missing.append('количество')
    if data.get('transport_num') == 'не найдено':
        missing.append('номер вагона/транспорта')
    if data.get('issue_day') == 'не найдено':
        missing.append('дата выдачи')
    return missing


# ============ ИНТЕРФЕЙС STREAMLIT ============

st.set_page_config(page_title="Переоформление ФСС", page_icon="🌾", layout="centered")

st.title("🌾 Автогенерация заявлений ФСС (DOCX + PDF)")
st.write("Загрузите PDF фитосанитарных сертификатов — сервис подготовит заявления.")

TEMPLATE_FILE = "template.docx"

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"❌ Файл шаблона '{TEMPLATE_FILE}' не найден.")
else:
    uploaded_files = st.file_uploader(
        "Перетащите сюда PDF-файлы (массовая загрузка поддерживается)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        soffice_available = _find_soffice() is not None
        if not soffice_available:
            st.warning(
                "⚠️ LibreOffice не найден — будут созданы только DOCX, "
                "без PDF. Проверьте packages.txt."
            )

        status_text = st.empty()
        status_text.info("🚀 Обработка файлов...")

        zip_buffer = io.BytesIO()
        success_count = 0
        pdf_failed = 0
        errors, warnings = [], []

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for uploaded_file in uploaded_files:
                safe_name = sanitize_filename(uploaded_file.name)
                try:
                    pdf_io = io.BytesIO(uploaded_file.read())
                    data = extract_data_from_pdf(pdf_io)

                    miss = _missing_fields(data)
                    if miss:
                        warnings.append(
                            f"{uploaded_file.name}: не распознаны — {', '.join(miss)}"
                        )

                    doc = DocxTemplate(TEMPLATE_FILE)
                    doc.render(data)

                    docx_io = io.BytesIO()
                    doc.save(docx_io)
                    docx_bytes = docx_io.getvalue()

                    zip_file.writestr(f"Заявление_{safe_name}.docx", docx_bytes)

                    if soffice_available:
                        pdf_bytes = convert_docx_to_pdf_server(docx_bytes)
                        if pdf_bytes:
                            zip_file.writestr(f"Заявление_{safe_name}.pdf", pdf_bytes)
                        else:
                            pdf_failed += 1

                    success_count += 1
                except Exception as e:
                    errors.append(f"{uploaded_file.name}: {e}")

        status_text.empty()

        if warnings:
            with st.expander(f"⚠️ Предупреждения по распознаванию ({len(warnings)})"):
                for w in warnings:
                    st.warning(w)

        if errors:
            with st.expander(f"❌ Ошибки ({len(errors)})", expanded=True):
                for e in errors:
                    st.error(e)

        if pdf_failed:
            st.warning(f"⚠️ PDF не создан для {pdf_failed} файлов.")

        if success_count > 0:
            st.success(f"🔥 Готово! Обработано файлов: {success_count}.")
            st.download_button(
                label="📥 Скачать архив с заявлениями (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Сгенерированные_Заявления.zip",
                mime="application/zip",
                use_container_width=True,
            )
        else:
            st.error("❌ Не удалось обработать ни один файл.")