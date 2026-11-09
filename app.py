# -*- coding: utf-8 -*-
"""
app.py — автогенерация заявлений ФСС (DOCX + PDF).
Дата определяется в три уровня:
  1. Текст PDF, собранный всеми методами pdfplumber.
  2. Номер сертификата (DDMMYY в позициях 9..15).
  3. Имя файла (например, 'ФСС за 22 января-1.pdf').
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
BOTANICAL_TO_PRODUCT = {
    'LINUM USITATISSIMUM': 'LINUM USITATISSIMUM L. / Семена льна (КОД ТН ВЭД 1204009000)',
    'BRASSICA NAPUS':      'BRASSICA NAPUS L. / Семена рапса (КОД ТН ВЭД 1205109000)',
    'SETARIA ITALICA':     'SETARIA ITALICA (L.) BEAUV / Пшено (КОД ТН ВЭД 1103199009)',
    'AVENA SATIVA':        'AVENA SATIVA L. / Овес (КОД ТН ВЭД 1004900000)',
    'TRITICUM AESTIVUM':   'TRITICUM AESTIVUM L. / Пшеница мягкая (КОД ТН ВЭД 1001990000)',
    'HORDEUM VULGARE':     'HORDEUM VULGARE L. / Ячмень (КОД ТН ВЭД 1003900000)',
}

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
    'ЯНВАРЯ': 'Января', 'ФЕВРАЛЯ': 'Февраля', 'МАРТА': 'Марта', 'АПРЕЛЯ': 'Апреля',
    'МАЯ': 'Мая', 'ИЮНЯ': 'Июня', 'ИЮЛЯ': 'Июля', 'АВГУСТА': 'Августа',
    'СЕНТЯБРЯ': 'Сентября', 'ОКТЯБРЯ': 'Октября', 'НОЯБРЯ': 'Ноября', 'ДЕКАБРЯ': 'Декабря',
}

MONTH_NAMES_RU = ['', 'Января', 'Февраля', 'Марта', 'Апреля', 'Мая', 'Июня',
                  'Июля', 'Августа', 'Сентября', 'Октября', 'Ноября', 'Декабря']

DEFAULT_DESTINATION = 'China, Manzhouli'


# ============ СБОР ТЕКСТА ============
def _page_lines_from_words(page):
    """Пересобирает строки из слов по вертикальным позициям.
    Ловит содержимое, которое extract_text() иногда теряет."""
    try:
        words = page.extract_words(use_text_flow=False)
    except Exception:
        return []
    if not words:
        return []
    rows = {}
    for w in words:
        try:
            key = round(float(w["top"]) / 3)
            rows.setdefault(key, []).append((float(w["x0"]), w["text"]))
        except (KeyError, TypeError, ValueError):
            continue
    lines = []
    for key in sorted(rows):
        parts = [t for _, t in sorted(rows[key])]
        lines.append(" ".join(parts))
    return lines


def _page_tables_text(page):
    try:
        tables = page.extract_tables()
    except Exception:
        return ""
    chunks = []
    for table in tables:
        for row in table:
            for cell in row:
                if cell:
                    chunks.append(str(cell))
    return "\n".join(chunks)


def _collect_page_text(page):
    """Максимум текста со страницы: extract_text, layout, слова, таблицы."""
    parts = []
    try:
        t = page.extract_text()
        if t:
            parts.append(t)
    except Exception:
        pass
    try:
        t = page.extract_text(layout=True)
        if t:
            parts.append(t)
    except Exception:
        pass
    lines = _page_lines_from_words(page)
    if lines:
        parts.append("\n".join(lines))
    tables = _page_tables_text(page)
    if tables:
        parts.append(tables)
    return "\n".join(parts)


# ============ ДАТА ============
_DATE_RE = re.compile(r'\b(\d{1,2})\s+([A-ZА-ЯЁ]{3,9})\.?\s+(\d{4})\b', re.IGNORECASE)


def _valid_date_candidates(text):
    out = []
    for m in _DATE_RE.finditer(text):
        try:
            day = int(m.group(1))
        except ValueError:
            continue
        month_key = m.group(2).upper()
        year = m.group(3)
        if not (1 <= day <= 31):
            continue
        month_ru = MONTHS_DICT.get(month_key)
        if not month_ru:
            continue
        # отбрасываем диапазоны «26 MAY 2026 - 31 MAY 2026»
        after = text[m.end():m.end() + 5]
        before = text[max(0, m.start() - 5):m.start()]
        if re.match(r'\s*[-\u2013\u2014]', after):
            continue
        if re.search(r'[-\u2013\u2014]\s*$', before):
            continue
        out.append((str(day).zfill(2), month_ru, year))
    return out


def _extract_issue_date(full_text):
    # 1) рядом с полем «Place of issue / Место выдачи»
    for marker in (r'Place\s+of\s+issue', r'Место\s+выдачи'):
        m_marker = re.search(marker, full_text, re.IGNORECASE)
        if not m_marker:
            continue
        tail = full_text[m_marker.end(): m_marker.end() + 500]
        cands = _valid_date_candidates(tail)
        if cands:
            return cands[0]
    # 2) последняя валидная дата во всём тексте
    cands = _valid_date_candidates(full_text)
    if cands:
        return cands[-1]
    return None


def _date_from_cert_number(cert_num):
    """Сертификаты Россельхознадзора: 18 цифр, [9:15] = DDMMYY."""
    if not cert_num or len(cert_num) < 15:
        return None
    try:
        dd = int(cert_num[9:11])
        mm = int(cert_num[11:13])
        yy = int(cert_num[13:15])
    except (ValueError, IndexError):
        return None
    if not (1 <= dd <= 31 and 1 <= mm <= 12):
        return None
    return f"{dd:02d}", MONTH_NAMES_RU[mm], str(2000 + yy)


def _date_from_filename(filename):
    name = os.path.basename(filename or '')
    pattern = re.compile(r'\b(\d{1,2})\s+([A-Za-zА-Яа-яЁё]{3,9})\.?(?:\s+(\d{4}))?')
    for m in pattern.finditer(name):
        try:
            day = int(m.group(1))
        except ValueError:
            continue
        if not (1 <= day <= 31):
            continue
        month_key = m.group(2).upper()
        month_ru = MONTHS_DICT.get(month_key) or MONTHS_DICT.get(month_key[:3])
        if not month_ru:
            continue
        return str(day).zfill(2), month_ru, m.group(3) or 'не найдено'
    return None


def resolve_issue_date(full_text, cert_num, filename):
    d = _extract_issue_date(full_text)
    if d:
        return (*d, 'текст PDF')
    d = _date_from_cert_number(cert_num)
    if d:
        return (*d, 'номер сертификата')
    d = _date_from_filename(filename)
    if d:
        return (*d, 'имя файла')
    return None, None, None, None


# ============ ОСТАЛЬНЫЕ ПОЛЯ ============
def sanitize_filename(name, fallback="document"):
    base = os.path.basename(name or "")
    base = os.path.splitext(base)[0]
    base = re.sub(r'[^\w\-]+', '_', base, flags=re.UNICODE).strip('._-')
    return base or fallback


def _extract_weight(text):
    pattern = re.compile(
        r'(\d[\d\s.,]*)\s*'
        r'(TONNES?|TONS?|MT|KGS?|KG|PCS?|PC|ТОНН[А-ЯЁ]*|КГ|ШТ|T|Т)'
        r'\.?'
        r'(?!\w)',
        re.IGNORECASE | re.UNICODE,
    )
    m = pattern.search(text)
    if not m:
        return "не найдено"
    raw = re.sub(r'\s+', '', m.group(1)).strip('.,')
    unit = m.group(2).upper().rstrip('.')
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


def _extract_transport(pages_text):
    ctx = re.compile(
        r'(?:CARRIAGE|TRANSPORT|CONTAINER|WAGON|ВАГОН|КОНТЕЙНЕР)'
        r'\s*[№#:/]?\s*([A-Z]{0,4}\d{4,})',
        re.IGNORECASE,
    )
    found = []
    for t in pages_text:
        for m in ctx.finditer(t):
            found.append(m.group(1).upper())
    if found:
        return ", ".join(sorted(set(found)))
    if len(pages_text) > 1:
        wagons = []
        for t in pages_text[1:]:
            wagons.extend(re.findall(r'\b(\d{8})\b', t))
        if wagons:
            return ", ".join(sorted(set(wagons)))
    return "не найдено"


def _extract_destination(text):
    """
    Логика:
    1) Прямой поиск известных значений — надёжнее любого контекста.
    2) Поиск после "point of entry / пункт ввоза" — но с фильтром
       строк-заголовков (layout=True подмешивает шапку таблицы).
    3) Фолбэк — DEFAULT_DESTINATION.
    """
    # 1) Прямой поиск
    if re.search(r'\bChina\s*,\s*Manzhouli\b', text, re.IGNORECASE):
        return 'China, Manzhouli'

    # 2) Контекстный поиск с отсечением заголовков
    label_words = re.compile(
        r'(Place\s+of|Declared|Заявленный|Кому|To\s*:|'
        r'Экспортер|Exporter|получатель|consignee|'
        r'Наименование|Name\s+of|Ботаническое|Botanical|'
        r'Количество|Number\s+of|Описание|Description|'
        r'Отличительные|Distinguishing|means\s+of\s+conveyance|'
        r'point\s+of\s+entry|origin)',
        re.IGNORECASE,
    )
    m = re.search(
        r'(?:пункт\s+ввоза|point\s+of\s+entry)\b[^\n]*\n',
        text, re.IGNORECASE,
    )
    if m:
        for line in text[m.end():].split('\n'):
            line = line.strip()
            if not line or len(line) > 200:
                continue
            if label_words.search(line):
                continue
            # значение должно быть короткой строкой вида "China, Manzhouli"
            return line

    return DEFAULT_DESTINATION


def _extract_product(full_upper):
    for key, value in BOTANICAL_TO_PRODUCT.items():
        if re.search(rf'\b{re.escape(key)}\b', full_upper):
            return value
    for key, value in COMMON_TO_PRODUCT.items():
        if re.search(rf'\b{re.escape(key)}\b', full_upper):
            return value
    return "Не удалось распознать название культуры"


# ============ ГЛАВНАЯ ФУНКЦИЯ ПАРСИНГА ============
def extract_data_from_pdf(pdf_stream, filename=""):
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

        pages_text = [_collect_page_text(p) for p in pdf.pages]
        full_text = "\n".join(pages_text)
        full_upper = full_text.upper()

        m = re.search(r'\b([A-Z]\d{7})\b', full_upper)
        if m:
            data['blank_series'] = m.group(1)

        m = re.search(r'[№#]\s*(\d{15,20})', full_text)
        if m:
            data['certificate_num'] = m.group(1)

        day, month_ru, year, source = resolve_issue_date(
            full_text, data['certificate_num'], filename
        )
        if day:
            data['issue_day'] = day
            data['issue_month'] = month_ru
            data['issue_year'] = year
            data['_date_source'] = source
        else:
            data['_date_source'] = None

        data['product_name'] = _extract_product(full_upper)
        data['quantity'] = _extract_weight(full_text)
        data['destination'] = _extract_destination(full_text)
        data['transport_num'] = _extract_transport(pages_text)

    return data


# ============ КОНВЕРТАЦИЯ В PDF ============
def _find_soffice():
    for name in ('soffice', 'libreoffice'):
        path = shutil.which(name)
        if path:
            return path
    return None


def convert_docx_to_pdf_server(docx_bytes):
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
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60,
            )
            if result.returncode == 0 and os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    return f.read()
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        pass
    return None


def _missing_fields(data):
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
        missing.append('номер транспорта')
    if data.get('issue_day') == 'не найдено':
        missing.append('дата выдачи')
    return missing


# ============ STREAMLIT UI ============
st.set_page_config(page_title="Переоформление ФСС", page_icon="🌾", layout="centered")
st.title("🌾 Автогенерация заявлений ФСС (DOCX + PDF)")
st.write("Загрузите PDF фитосанитарных сертификатов — сервис подготовит заявления.")

TEMPLATE_FILE = "template.docx"

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"❌ Файл шаблона '{TEMPLATE_FILE}' не найден.")
else:
    uploaded_files = st.file_uploader(
        "Перетащите PDF-файлы (массовая загрузка поддерживается)",
        type=["pdf"], accept_multiple_files=True,
    )

    if uploaded_files:
        soffice_available = _find_soffice() is not None
        if not soffice_available:
            st.warning("⚠️ LibreOffice не найден — будут только DOCX без PDF.")

        status_text = st.empty()
        status_text.info("🚀 Обработка файлов...")

        zip_buffer = io.BytesIO()
        success_count = 0
        pdf_failed = 0
        errors, warnings, details = [], [], []

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for uploaded_file in uploaded_files:
                safe_name = sanitize_filename(uploaded_file.name)
                try:
                    pdf_io = io.BytesIO(uploaded_file.read())
                    data = extract_data_from_pdf(pdf_io, uploaded_file.name)

                    miss = _missing_fields(data)
                    if miss:
                        warnings.append(
                            f"{uploaded_file.name}: не распознаны — {', '.join(miss)}"
                        )

                    date_source = data.pop('_date_source', None)
                    details.append(
                        f"{uploaded_file.name}: "
                        f"дата = {data['issue_day']} {data['issue_month']} {data['issue_year']} "
                        f"(источник: {date_source or '—'}), "
                        f"вес = {data['quantity']}"
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

        if details:
            with st.expander(f"ℹ️ Что распознано ({len(details)})"):
                for d in details:
                    st.text(d)

        if warnings:
            with st.expander(f"⚠️ Предупреждения ({len(warnings)})"):
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