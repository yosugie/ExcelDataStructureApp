# -*- coding: utf-8 -*-
"""
Извлечение номеров заказа/детали из эскизов для таблицы учёта.

Объединяет два источника в одной программе с автоопределением типа файла:

  1) .bln — библиотека Базис Мебельщик (индивидуальные заказы).
     Формат — старый Microsoft-контейнер Compound File Binary / OLE2.
     Внутри хранится служебный XML-индекс "$$Lib_structure$$", где
     перечислены папки/файлы библиотеки. Папка с эскизами ищется по
     двум признакам сразу (конструкторы называют её по-разному):
       - имя папки содержит "эск";
       - ИЛИ имя файла начинается с "Эск<номер>".
     Код детали "номер_заказа номер_секции номер_детали" ищется прямо
     в бинарных данных самого файла чертежа (.ldw) — это надёжнее,
     чем разбор имени файла, и работает даже когда конструктор не
     вынес код в название. Разбор имени файла используется как запасной
     способ. Свой минимальный читатель CFBF — без внешних библиотек
     (в рабочей среде нет интернета, чтобы поставить olefile).

  2) .pdf — эскизы из inSight INQ (розничные стандартные заказы).
     Номер заказа и номер(а) детали зашиты как скрытый текст в PDF
     (штрихкод-текст вида *ORD######* и ICN######, либо несколько
     номеров через "num;num;num" на одном листе) — читается через
     pdfplumber, без OCR.
     Тип для столбца F определяется по имени файла:
       - имя содержит "корпус"  -> "inSight"
       - имя содержит "прочее"  -> "inSight (4 уч.)"
     Если по имени не удалось определить — по умолчанию ставится
     "inSight" с предупреждением в журнале; тип всегда можно поменять
     вручную в выпадающем списке наверху, изменение применится сразу
     ко всем строкам.

Как запустить:
    pip install pdfplumber
    (опционально, для drag-and-drop) pip install tkinterdnd2
    (опционально, для кнопки-календаря у поля даты) pip install tkcalendar
    python ExcelDataStructureApp.py
"""

import os
import re
import sys
import struct
import ctypes
import datetime
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

try:
    from tkcalendar import Calendar
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

VK_C = 67
VK_V = 86
VK_X = 88
VK_A = 65

SOURCE_TYPES = ("Bazis", "inSight", "inSight (4 уч.)")

# Иконка приложения (32x32 и 64x64 PNG, встроены как base64, чтобы не таскать
# отдельный файл рядом со скриптом/exe).
ICON_PNG_32 = (
    "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAq0lEQVR4nO2XwQ2AIAxFS+PRARjDeRzQeRyjA3jXk0kDrSjRNib9J/IJ/EegBBIw5Zx3MBARpbON1uFlFlqHlxDJI5xrkMx13j4Jm5ax8rA0vgrX5q4AWpJWceW39BjgbQVAAIj3AMD1qe6pBK28VQBtwLSMYp/mt+S+BQHgDhBV4L4FARAA/wPQSq33MVsB9D4u70ia2/1jgvyjaC0iSng2PMIB2BmwhOBZB4IZRQw9Z0acAAAAAElFTkSuQmCC"
)
ICON_PNG_64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAABYUlEQVR4nO2bUQ6DIBBEV9LPHsBj9Dw9YM/TY3CA/rdfJAYVd0l1wu68T2GbmVGgEJ2kwTzP31b7KOScp722zQYvxmu2gkj1Ba/mRba9paMO3qg9pr0Gzyy9pvpCFIrn1RwQjSni3V9ys3R+Pz9n6fgrj9dd3Vc9BEYxL2LTqgpgJPMFrebDAEY0X9Bov2QVeLzupnHZW9ND+GWQAaAFoGEAaAFoGABaABoGgBaAhgGgBaAJH4DpQETEdtjwj1prjXX3yifAWtBzPlDuoqW2p6aH8E8AA0ALQMMA0ALQMAC0ADQMAC0ATfgAuBky9XYIN0On/voAMAC0ADQMAC0ADQNAC0DDANAC0DAAtAA05r1ADz3/5696QfPwCbjiZcWz0GhXDYERQ9BqVs8BI4Vg0Rr+g4nU+qjQOznnKfwymETan5Z6pXhO9YUILL2mvQav1B5Xc4DnELa8Nc16WSJbN/UHRltrEBL1qTAAAAAASUVORK5CYII="
)


def set_windows_dark_titlebar(root, dark):
    """На Windows 10/11 красит системную шапку окна в тёмный цвет, чтобы она не
    контрастировала белым на фоне тёмной темы приложения. На других ОС — no-op."""
    if sys.platform != "win32":
        return
    try:
        root.update_idletasks()  # окно должно быть полностью создано, иначе хэндл ещё не готов
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        for attribute in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE: 20 (новее), 19 (старые сборки Win10)
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value)
            )
            if result == 0:
                break
    except Exception:
        pass  # косметика, не критично, если недоступно


def fix_clipboard_shortcuts(widget):
    """Чинит Ctrl+C/V/X и Ctrl+A в Entry/Text независимо от раскладки клавиатуры."""

    def handler(event):
        code = event.keycode
        if code == VK_V:
            try:
                widget.event_generate("<<Paste>>")
            except tk.TclError:
                pass
            return "break"
        elif code == VK_C:
            try:
                widget.event_generate("<<Copy>>")
            except tk.TclError:
                pass
            return "break"
        elif code == VK_X:
            try:
                widget.event_generate("<<Cut>>")
            except tk.TclError:
                pass
            return "break"
        elif code == VK_A:
            if isinstance(widget, tk.Text):
                widget.tag_add("sel", "1.0", "end")
            elif isinstance(widget, tk.Entry):
                widget.select_range(0, "end")
            return "break"
        return None

    widget.bind("<Control-KeyPress>", handler)


# ---------------------------------------------------------------------------
# Минимальный читатель CFBF/OLE2 (формат .bln), без внешних библиотек.
# ---------------------------------------------------------------------------

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE


class CfbReadError(Exception):
    pass


class CfbContainer:
    """Открывает CFBF-файл один раз и позволяет доставать любые потоки по имени."""

    def __init__(self, path):
        with open(path, "rb") as f:
            self.data = f.read()

        if self.data[0:8] != b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
            raise CfbReadError("Файл не похож на .bln (нет сигнатуры OLE2/CFBF)")

        header = self.data[0:512]
        sector_shift = struct.unpack_from("<H", header, 30)[0]
        first_dir_sector = struct.unpack_from("<I", header, 48)[0]
        first_difat_sector = struct.unpack_from("<I", header, 68)[0]
        self.sector_size = 1 << sector_shift

        difat = list(struct.unpack_from("<109I", header, 76))
        sec = first_difat_sector
        while sec not in (ENDOFCHAIN, FREESECT) and sec < 0xFFFFFFF0:
            s = self._read_sector(sec)
            difat.extend(struct.unpack_from("<%dI" % (self.sector_size // 4 - 1), s, 0))
            sec = struct.unpack_from("<I", s, self.sector_size - 4)[0]

        self.fat = []
        for sec_id in difat:
            if sec_id == FREESECT or sec_id >= 0xFFFFFFF0:
                continue
            s = self._read_sector(sec_id)
            self.fat.extend(struct.unpack_from("<%dI" % (self.sector_size // 4), s, 0))

        dir_data = self._read_chain(first_dir_sector)
        n_entries = len(dir_data) // 128
        self.index = {}
        for i in range(n_entries):
            e = dir_data[i * 128:(i + 1) * 128]
            name_len = struct.unpack_from("<H", e, 64)[0]
            raw_name = e[0:max(name_len - 2, 0)] if name_len >= 2 else b""
            name = raw_name.decode("utf-16-le", errors="replace")
            start_sec = struct.unpack_from("<I", e, 116)[0]
            stream_size = struct.unpack_from("<Q", e, 120)[0]
            self.index[name] = (start_sec, stream_size)

    def _read_sector(self, sec_id):
        off = 512 + sec_id * self.sector_size
        return self.data[off: off + self.sector_size]

    def _read_chain(self, start_sec, size=None):
        chunks = []
        sec = start_sec
        seen = set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < 0xFFFFFFF0:
            if sec in seen:
                break
            seen.add(sec)
            chunks.append(self._read_sector(sec))
            sec = self.fat[sec] if sec < len(self.fat) else ENDOFCHAIN
        result = b"".join(chunks)
        if size is not None:
            result = result[:size]
        return result

    def get_stream(self, name):
        if name not in self.index:
            return None
        start_sec, size = self.index[name]
        return self._read_chain(start_sec, size)


# ---------------------------------------------------------------------------
# Базис: разбор структуры библиотеки и поиск эскизов
# ---------------------------------------------------------------------------

# Код детали, зашитый прямо в бинарные данные чертежа: "214704 01 003"
CODE_IN_STREAM_RE = re.compile(rb"(\d{5,7})\s+(\d{2}\s+\d{3})")

# На случай, если код есть прямо в имени файла: "Эск1 03 010 (Бок правый).ldw"
CODE_IN_NAME_RE = re.compile(r"(\d{2}\s+\d{3})\s*(?:\(([^)]*)\))?")

# Описание в скобках, если оно есть в имени файла: "... (Бок левый).ldw"
DESC_RE = re.compile(r"\(([^)]*)\)")

# Сборочный чертёж (не отдельная деталь), например:
# "214663 02 СБ на Фасад GARDA 69 L в946ш564 (без ручки) Лист1"
ASSEMBLY_WORD_RE = re.compile(r"\bСБ\b", re.IGNORECASE)
ASSEMBLY_SECTION_RE = re.compile(r"(\d{2})\s+СБ", re.IGNORECASE)

# Материал детали, зашитый в бинарных данных чертежа как строка UTF-8 с
# 4-байтовым числом (длина строки) прямо перед ней, например:
# "ЛДСП 16 Белый премиум W1000 ST9 (Артикул ЭР00442)"
MATERIAL_MARKER = "Артикул".encode("utf-8")


def extract_material(stream):
    """Ищет строку материала в бинарных данных чертежа (см. MATERIAL_MARKER)."""
    idx = stream.find(MATERIAL_MARKER)
    if idx == -1:
        return None
    lowest_start = max(0, idx - 4 - 300)
    for start in range(idx - 4, lowest_start - 1, -1):
        length = struct.unpack_from("<I", stream, start)[0]
        if not (5 <= length <= 250):
            continue
        end = start + 4 + length
        if end > len(stream) or end <= idx:
            continue
        candidate = stream[start + 4:end]
        try:
            decoded = candidate.decode("utf-8")
        except UnicodeDecodeError:
            continue
        if "Артикул" in decoded and decoded.isprintable():
            return decoded
    return None


# Толщина — первое число в строке материала: "ЛДСП 16 Белый ..." -> 16
MATERIAL_THICKNESS_RE = re.compile(r"^\S+\s+(\d+)")

# Материалы такой толщины (мм) станок не пилит — такие строки по умолчанию
# идут со снятой галочкой копирования, как и сборочные чертежи.
NOT_MACHINABLE_THICKNESS_MM = {3}


def material_thickness_mm(material):
    if not material:
        return None
    m = MATERIAL_THICKNESS_RE.match(material)
    if not m:
        return None
    try:
        return int(m.group(1))
    except ValueError:
        return None


# Названия деталей, которые станок не пилит (например рейки для сборки на
# месте, не сама деталь из плиты) — так же, как сборочные чертежи и
# материалы 3мм, идут со снятой галочкой копирования по умолчанию.
NOT_MACHINABLE_PART_NAMES = {"рейка на щит"}


def is_excluded_part_name(description):
    return bool(description) and description.strip().lower() in NOT_MACHINABLE_PART_NAMES


def looks_like_sketch_folder(name):
    return "эск" in (name or "").lower()


def looks_like_sketch_filename(name):
    return bool(re.match(r"^Эск(?:из)?\s*\d+", name or "", re.IGNORECASE))


def is_assembly_filename(name):
    """Сборочный чертёж (СБ) — показываем в таблице для проверки, но не
    копируем как отдельную деталь (это не то же самое, что "деталь")."""
    return bool(ASSEMBLY_WORD_RE.search(name or ""))


def parse_bln_sketches(bln_path):
    """Возвращает (order_number, [ {part_code, description, source_dir} ], warnings)."""
    warnings = []
    container = CfbContainer(bln_path)

    xml_raw = container.get_stream("$$Lib_structure$$")
    if xml_raw is None:
        raise CfbReadError(
            'Внутри файла не найден служебный индекс "$$Lib_structure$$". '
            "Возможно, это не библиотека Базис Мебельщик, либо другая версия формата."
        )

    text = xml_raw.decode("utf-8-sig", errors="replace")
    try:
        root = ET.fromstring(text)
    except ET.ParseError as e:
        raise CfbReadError(f"Не удалось разобрать внутренний XML-индекс: {e}")

    candidates = []
    for directory in root.findall("Directory"):
        name_el = directory.find("Name")
        dir_name = name_el.text if name_el is not None else ""
        is_sketch_dir = looks_like_sketch_folder(dir_name)
        for file_el in directory.findall("File"):
            fname_el = file_el.find("Name")
            storage_el = file_el.find("Storage")
            if fname_el is None or storage_el is None or not fname_el.text:
                continue
            fname = fname_el.text
            if is_sketch_dir or looks_like_sketch_filename(fname):
                candidates.append((dir_name, fname, storage_el.text))

    results = []
    order_votes = {}

    n_assembly = 0
    n_not_machinable = 0

    for dir_name, fname, storage in candidates:
        order_number = None
        part_codes = []
        description = ""
        material = None
        is_assembly = is_assembly_filename(fname)

        stream = container.get_stream(storage)
        if stream:
            if not is_assembly:
                # Один лист иногда содержит сразу две детали (например одна
                # заготовка распиливается на две) — берём все найденные коды,
                # а не только первый.
                seen = set()
                for code_m in CODE_IN_STREAM_RE.finditer(stream):
                    if order_number is None:
                        order_number = code_m.group(1).decode()
                    candidate_part = code_m.group(2).decode()
                    if candidate_part not in seen:
                        seen.add(candidate_part)
                        part_codes.append(candidate_part)
            else:
                code_m = CODE_IN_STREAM_RE.search(stream)
                if code_m:
                    order_number = code_m.group(1).decode()
            material = extract_material(stream)

        thickness_mm = material_thickness_mm(material)
        is_too_thin = thickness_mm in NOT_MACHINABLE_THICKNESS_MM

        if not part_codes and not is_assembly:
            m2 = CODE_IN_NAME_RE.search(fname)
            if m2:
                part_codes = [m2.group(1)]

        desc_m = DESC_RE.search(fname)
        if desc_m:
            description = desc_m.group(1).strip()

        is_excluded_name = is_excluded_part_name(description)

        if is_assembly:
            n_assembly += 1
            sec_m = ASSEMBLY_SECTION_RE.search(fname)
            part_codes = [f"{sec_m.group(1)} СБ" if sec_m else "СБ"]
            description = f"[Сборочный чертёж, галочка снята] {fname}"
        elif is_too_thin or is_excluded_name:
            n_not_machinable += 1

        if not part_codes:
            warnings.append(f'Не удалось определить код детали для "{fname}" — пропущено.')
            continue

        if len(part_codes) > 1:
            sheet_label = os.path.splitext(fname)[0]
            warnings.append(
                f'Лист {sheet_label}: на одном чертеже найдено несколько деталей '
                f'({" / ".join(part_codes)}) — добавлена отдельная строка для каждой.'
            )

        if order_number:
            order_votes[order_number] = order_votes.get(order_number, 0) + 1

        for part_code in part_codes:
            results.append({
                "order_from_content": order_number,
                "part_code": part_code,
                "description": description,
                "material": material,
                "source_dir": dir_name,
                "raw_name": fname,
                "is_assembly": is_assembly,
                "auto_exclude": is_assembly or is_too_thin or is_excluded_name,
            })

    if n_not_machinable:
        thickness_list = ", ".join(str(t) for t in sorted(NOT_MACHINABLE_THICKNESS_MM))
        warnings.append(
            f"Найдено деталей, которые станок не пилит (толщина {thickness_list} мм или "
            f"особое название вроде «Рейка на щит»): {n_not_machinable} — показаны в таблице "
            f"для проверки, галочка копирования снята по умолчанию."
        )

    if n_assembly:
        warnings.append(
            f"Найдено сборочных чертежей (СБ): {n_assembly} — показаны в таблице для проверки, "
            f"но не войдут в копирование строк."
        )

    # Итоговый номер заказа: приоритет — то, что реально найдено внутри чертежей
    # (самое частое совпадение), а не имя файла .bln, которое могут переименовать.
    if order_votes:
        order_number = max(order_votes.items(), key=lambda kv: kv[1])[0]
    else:
        base = os.path.splitext(os.path.basename(bln_path))[0]
        digits = re.sub(r"\D", "", base)
        order_number = digits if 5 <= len(digits) <= 7 else None
        if order_number:
            warnings.append(
                "Номер заказа не найден внутри чертежей — использован номер из имени файла .bln."
            )
        else:
            warnings.append("Не удалось определить номер заказа автоматически.")

    if not results:
        warnings.append(
            "Не нашёл ни одного эскиза в этой библиотеке. Возможно, здесь другой формат "
            "именования папки/файлов — пришлите пример для проверки."
        )

    return order_number, results, warnings


# ---------------------------------------------------------------------------
# inSight: разбор PDF-эскизов
# ---------------------------------------------------------------------------

ORDER_RE = re.compile(r"\*ORD(\d+)\*")
ORDER_SPACED_RE = re.compile(r"\b(\d{3})\s(\d{3})\b")
ICN_RE = re.compile(r"ICN(\d+)")
MULTI_PART_RE = re.compile(r"\b(\d{6,9}(?:\s*;\s*\d{6,9})+)\b")

# Материал с толщиной в тексте PDF, например:
# "0 ЭР35521 ЛДСП 16 Кашемир U702 СТ9 2800х2070 ..." -> "ЛДСП 16 Кашемир U702 СТ9"
MATERIAL_PDF_RE = re.compile(
    r"(ЛДСП|МДФ|ХДФ|ЛХДФ|ДВП|ДСП|Массив|Стекло|Зеркало)\s+\d+.*?(?=\s+\d+х\d+)"
)


def extract_material_pdf(text: str):
    m = MATERIAL_PDF_RE.search(text)
    return m.group(0).strip() if m else None


# Название детали в тексте PDF — идёт перед размерами вида "267.0* 700.0",
# например: "Панель под рейки 2 267.0* 700.0 2шт ori:92111063" -> "Панель под рейки"
PART_NAME_RE = re.compile(r"^(.+?)\s+(?:\d+\s+)?\d+\.\d+\*", re.MULTILINE)


def extract_part_name_pdf(text: str):
    m = PART_NAME_RE.search(text)
    return m.group(1).strip() if m else None


def extract_order_number(text: str):
    """Извлечь номер заказа, убрать пробелы."""
    m = ORDER_RE.search(text)
    if m:
        return m.group(1)
    # запасной вариант: номер вида "214 090" где-то в тексте
    m = ORDER_SPACED_RE.search(text)
    if m:
        return m.group(1) + m.group(2)
    return None


def extract_part_numbers(text: str):
    """Извлечь номер(а) детали согласно правилам ТЗ."""
    m = MULTI_PART_RE.search(text)
    if m:
        return m.group(1)  # оставляем как в PDF, включая пробелы после ';'

    m = ICN_RE.search(text)
    if m:
        return m.group(0)  # с префиксом ICN

    return None


def detect_pdf_source_type(filename):
    """По имени файла определяет "inSight" или "inSight (4 уч.)".

    "План запуска в производство корпус_ДД.ММ.ГГГГ.pdf" -> inSight
    "Запуск Прочее_ДД.ММ.ГГГГ.pdf"                       -> inSight (4 уч.)
    Возвращает None, если по имени не удалось определить.
    """
    name = (filename or "").lower()
    if "прочее" in name:
        return "inSight (4 уч.)"
    if "корпус" in name:
        return "inSight"
    return None


def parse_pdf_sketches(pdf_path):
    """Возвращает (order_number, [ {part_code, description, source_dir} ], warnings).

    Формат результата такой же, как у parse_bln_sketches, чтобы GUI мог
    обрабатывать оба источника одним и тем же кодом.
    """
    if not HAS_PDFPLUMBER:
        raise CfbReadError(
            "Не установлена библиотека pdfplumber.\n\n"
            "Выполните в командной строке:\npip install pdfplumber"
        )

    warnings = []
    results = []
    order_votes = {}

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            order = extract_order_number(text) if text.strip() else None
            part = extract_part_numbers(text) if text.strip() else None
            material = extract_material_pdf(text) if text.strip() else None
            part_name = extract_part_name_pdf(text) if text.strip() else None
            thickness_mm = material_thickness_mm(material)
            is_too_thin = thickness_mm in NOT_MACHINABLE_THICKNESS_MM
            is_excluded_name = is_excluded_part_name(part_name)

            if order:
                order_votes[order] = order_votes.get(order, 0) + 1

            results.append({
                "order_from_content": order,
                "part_code": part,
                "description": part_name or "",
                "material": material,
                "source_dir": "",
                "raw_name": os.path.basename(pdf_path),
                "auto_exclude": is_too_thin or is_excluded_name,
                "page": i,
            })

    if order_votes:
        order_number = max(order_votes.items(), key=lambda kv: kv[1])[0]
    else:
        order_number = None
        warnings.append("Не удалось определить номер заказа ни на одной странице PDF.")

    n_missing_part = sum(1 for r in results if not r["part_code"])
    if n_missing_part:
        warnings.append(
            f'Номер детали не найден на {n_missing_part} стр. из {len(results)} — отмечено "Не найдено".'
        )

    n_not_machinable = sum(1 for r in results if r.get("auto_exclude"))
    if n_not_machinable:
        thickness_list = ", ".join(str(t) for t in sorted(NOT_MACHINABLE_THICKNESS_MM))
        warnings.append(
            f"Найдено деталей, которые станок не пилит (толщина {thickness_list} мм или "
            f"особое название вроде «Рейка на щит»): {n_not_machinable} — показаны в таблице "
            f"для проверки, галочка копирования снята по умолчанию."
        )

    if not results:
        warnings.append("В этом PDF не найдено ни одной страницы.")

    return order_number, results, warnings


# ---------------------------------------------------------------------------
# Интерфейс
# ---------------------------------------------------------------------------

THEMES = {
    "light": {
        "bg": "#f4f6f8",
        "fg": "#1a1a1a",
        "muted_fg": "#666666",
        "entry_bg": "#ffffff",
        "entry_fg": "#1a1a1a",
        "widget_bg": "#e4e7eb",
        "widget_fg": "#1a1a1a",
        "bg_alt": "#e4e4e4",
        "tree_bg": "#ffffff",
        "tree_fg": "#1a1a1a",
        "tree_heading_bg": "#e4e7eb",
        "tree_heading_fg": "#1a1a1a",
        "tree_selected_bg": "#2196F3",
        "tree_selected_fg": "#ffffff",
        "assembly_bg": "#eeeeee",
        "assembly_fg": "#888888",
        "log_bg": "#ffffff",
        "log_fg": "#1a1a1a",
        "border": "#c9d0d6",
        "accent": "#2196F3",
        "accent_fg": "#ffffff",
    },
    "dark": {
        # Тёмная тема в стиле производственных программ (чёрный фон + оранжевые акценты)
        "bg": "#181818",
        "fg": "#f0f0f0",
        "muted_fg": "#b5b5b5",
        "entry_bg": "#242424",
        "entry_fg": "#f0f0f0",
        "widget_bg": "#e8720c",
        "widget_fg": "#181818",
        "bg_alt": "#212121",
        "tree_bg": "#1f1f1f",
        "tree_fg": "#e8e8e8",
        "tree_heading_bg": "#101010",
        "tree_heading_fg": "#ffffff",
        "tree_selected_bg": "#e8720c",
        "tree_selected_fg": "#181818",
        "assembly_bg": "#2a2a2a",
        "assembly_fg": "#8a8a8a",
        "log_bg": "#1f1f1f",
        "log_fg": "#e8e8e8",
        "border": "#3a3a3a",
        "accent": "#e8720c",
        "accent_fg": "#181818",
    },
}


class SketchExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("Парсер учёта эскизов")
        root.geometry("1300x820")
        try:
            self._icon_imgs = [
                tk.PhotoImage(data=ICON_PNG_32),
                tk.PhotoImage(data=ICON_PNG_64),
            ]
            root.iconphoto(True, *self._icon_imgs)
        except tk.TclError:
            pass

        pad = {"padx": 8, "pady": 4}

        self.theme = "dark"

        top = tk.Frame(root)
        top.pack(fill="x", **pad)
        self.theme_btn = tk.Button(
            top, text="☀", width=3, command=self.toggle_theme, relief="flat", cursor="hand2",
        )
        self.theme_btn.pack(side="right")
        tk.Label(top, text="Файл (.bln или .pdf):", width=18, anchor="w").pack(side="left")
        self.path_entry = tk.Entry(top)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=4)
        fix_clipboard_shortcuts(self.path_entry)
        tk.Button(
            top, text="Обзор...", command=self.browse_file, relief="flat", cursor="hand2",
        ).pack(side="left")

        self.path_entry.bind("<Return>", lambda event: self.run_parse())

        if HAS_DND:
            self.path_entry.drop_target_register(DND_FILES)
            self.path_entry.dnd_bind("<<Drop>>", self.on_drop)

        opts_frame = tk.Frame(root)
        opts_frame.pack(fill="x", **pad)

        tk.Label(opts_frame, text="Тип эскиза:", width=18, anchor="w").pack(side="left")
        self.type_var = tk.StringVar(value="")
        self.type_combo = ttk.Combobox(
            opts_frame, textvariable=self.type_var, values=SOURCE_TYPES,
            state="disabled", width=16,
        )
        self.type_combo.pack(side="left", padx=4)
        self.type_combo.bind("<<ComboboxSelected>>", self.on_type_change)

        date_frame = tk.Frame(root)
        date_frame.pack(fill="x", **pad)
        tk.Label(date_frame, text="Дата запуска:", width=22, anchor="w").pack(side="left")
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%d.%m.%Y"))
        self.date_entry = tk.Entry(date_frame, width=14, textvariable=self.date_var)
        self.date_entry.pack(side="left", padx=4)
        fix_clipboard_shortcuts(self.date_entry)
        self.date_var.trace_add("write", lambda *args: self.refresh_date_column())
        self._install_date_validation(self.date_entry)
        self.calendar_btn = None
        if HAS_TKCALENDAR:
            self.calendar_btn = tk.Button(
                date_frame, text="📅", command=self.open_calendar_picker, relief="flat", cursor="hand2",
            )
            self.calendar_btn.pack(side="left", padx=(4, 0))
        self.today_btn = tk.Button(
            date_frame, text="Сегодня", command=self.set_today, relief="flat", cursor="hand2",
        )
        self.today_btn.pack(side="left", padx=(4, 0))

        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(root, textvariable=self.status_var, anchor="w")
        self.status_label.pack(fill="x", padx=8, pady=(0, 6))

        # Таблица результатов
        self.tree_frame = tree_frame = tk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=4)

        columns = ("include", "date", "order", "part", "type", "edge", "material", "description", "page", "folder")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=16)
        headers = {
            "include": "Копир.",
            "date": "Дата запуска",
            "order": "№ заказа",
            "part": "№ детали",
            "type": "Тип эскиза",
            "edge": "Кромка",
            "material": "Материал",
            "description": "Описание",
            "page": "Страница",
            "folder": "Папка / файл",
        }
        widths = {
            "include": 55, "date": 90, "order": 80, "part": 100, "type": 100,
            "edge": 60, "material": 260, "description": 220, "page": 70, "folder": 200,
        }
        for c in columns:
            self.tree.heading(c, text=headers[c], anchor="center")
            self.tree.column(c, width=widths[c], anchor="center")
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<space>", self.on_tree_space)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=8, pady=6)
        self.COPY_BTN_TEXT = "Скопировать строки"
        self.copy_btn = tk.Button(
            btn_frame,
            text=self.COPY_BTN_TEXT,
            command=self.copy_for_table,
            relief="flat", cursor="hand2",
        )
        self.copy_btn.pack(side="left", ipadx=14, ipady=6)
        self._copy_btn_reset_job = None

        self.clear_btn = tk.Button(
            btn_frame, text="Очистить", command=self.clear_all, relief="flat", cursor="hand2",
        )
        self.clear_btn.pack(side="left", padx=(8, 0), ipadx=14, ipady=6)

        self.log_expanded = False
        self.log_count = 0

        log_frame = tk.Frame(root)
        log_frame.pack(fill="x", padx=8, pady=(0, 8))

        self.log_header_var = tk.StringVar(value="▸ Журнал")
        self.log_header = tk.Label(
            log_frame, textvariable=self.log_header_var, anchor="w",
            cursor="hand2", font=("TkDefaultFont", 9, "bold"),
        )
        self.log_header.pack(fill="x")
        self.log_header.bind("<Button-1>", lambda event: self.toggle_log())

        self.log_body = tk.Frame(log_frame)
        self.log_text = scrolledtext.ScrolledText(self.log_body, height=12, state="disabled")
        self.log_text.pack(fill="x")
        fix_clipboard_shortcuts(self.log_text)
        # log_body остаётся не упакованным — журнал стартует свёрнутым.

        self.current_rows = []  # список dict-ов с результатами разбора
        self.current_kind = None  # "bazis" или "pdf"

        self.apply_theme()

    def toggle_log(self, expanded=None):
        self.log_expanded = (not self.log_expanded) if expanded is None else expanded
        if self.log_expanded:
            self.log_body.pack(fill="x", pady=(4, 0))
            arrow = "▾"
        else:
            self.log_body.pack_forget()
            arrow = "▸"
        suffix = f" ({self.log_count})" if self.log_count else ""
        self.log_header_var.set(f"{arrow} Журнал{suffix}")

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self.theme_btn.configure(text="☀" if self.theme == "dark" else "🌙")
        self.apply_theme()

    def _copy_btn_colors(self):
        t = THEMES[self.theme]
        return t["accent"], t["accent_fg"]

    def apply_theme(self):
        t = THEMES[self.theme]
        set_windows_dark_titlebar(self.root, dark=(self.theme == "dark"))

        def walk(widget):
            if isinstance(widget, ttk.Widget):
                # ttk-виджеты (Combobox, Treeview, Scrollbar) стилизуются через ttk.Style ниже.
                pass
            elif isinstance(widget, tk.Frame):
                widget.configure(bg=t["bg"], highlightbackground=t["border"], highlightcolor=t["border"])
            elif widget is self.status_label:
                widget.configure(bg=t["bg"], fg=t["muted_fg"])
            elif isinstance(widget, tk.Label):
                widget.configure(bg=t["bg"], fg=t["fg"])
            elif isinstance(widget, tk.Entry):
                widget.configure(
                    bg=t["entry_bg"], fg=t["entry_fg"],
                    insertbackground=t["entry_fg"], disabledbackground=t["entry_bg"],
                )
            elif widget in (self.clear_btn, self.calendar_btn, self.today_btn):
                widget.configure(
                    bg=t["bg_alt"], fg=t["fg"],
                    activebackground=t["bg_alt"], activeforeground=t["fg"],
                )
            elif widget is self.copy_btn:
                bg, fg = self._copy_btn_colors()
                widget.configure(bg=bg, fg=fg, activebackground=bg, activeforeground=fg)
            elif isinstance(widget, tk.Button):
                widget.configure(
                    bg=t["widget_bg"], fg=t["widget_fg"],
                    activebackground=t["widget_bg"], activeforeground=t["widget_fg"],
                )
            elif isinstance(widget, tk.Text):
                widget.configure(bg=t["log_bg"], fg=t["log_fg"], insertbackground=t["log_fg"])
            for child in widget.winfo_children():
                walk(child)

        self.root.configure(bg=t["bg"])
        walk(self.root)

        # Рамка вокруг таблицы результатов — тонкая панель, как на референсе.
        self.tree_frame.configure(highlightthickness=1, highlightbackground=t["border"], highlightcolor=t["border"])

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=t["tree_bg"], fieldbackground=t["tree_bg"], foreground=t["tree_fg"], borderwidth=0)
        style.map("Treeview", background=[("selected", t["tree_selected_bg"])], foreground=[("selected", t["tree_selected_fg"])])
        style.configure("Treeview.Heading", background=t["tree_heading_bg"], foreground=t["tree_heading_fg"], relief="flat")
        style.map("Treeview.Heading", background=[("active", t["tree_heading_bg"])])
        style.configure("TCombobox", fieldbackground=t["entry_bg"], background=t["widget_bg"], foreground=t["entry_fg"])
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", t["entry_bg"]), ("disabled", t["entry_bg"])],
            foreground=[("readonly", t["entry_fg"]), ("disabled", t["muted_fg"])],
            background=[("readonly", t["widget_bg"]), ("disabled", t["widget_bg"])],
            selectbackground=[("readonly", t["entry_bg"])],
            selectforeground=[("readonly", t["entry_fg"])],
        )
        style.configure(
            "TScrollbar", background=t["accent"], troughcolor=t["bg"],
            bordercolor=t["bg"], arrowcolor=t["fg"],
        )
        self.tree.tag_configure("assembly", background=t["assembly_bg"], foreground=t["assembly_fg"])

    def log(self, message):
        self.log_text.configure(state="normal")
        prefix = "\n" if self.log_text.get("1.0", "end-1c") else ""
        self.log_text.insert(tk.END, prefix + message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")
        self.log_count += 1
        if message.startswith("⚠") and not self.log_expanded:
            self.toggle_log(expanded=True)
        else:
            self.toggle_log(expanded=self.log_expanded)  # обновить счётчик в заголовке

    def set_today(self):
        self.date_var.set(datetime.date.today().strftime("%d.%m.%Y"))

    def refresh_date_column(self):
        new_date = self.date_var.get()
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            vals[1] = new_date  # индекс столбца "date"
            self.tree.item(iid, values=vals)

    def toggle_row_include(self, iid):
        idx = self.tree.index(iid)
        if idx >= len(self.current_rows):
            return
        row = self.current_rows[idx]
        row["include"] = not row["include"]
        vals = list(self.tree.item(iid, "values"))
        vals[0] = "☑" if row["include"] else "☐"
        self.tree.item(iid, values=vals)

    def on_tree_click(self, event):
        # Клик по столбцу "Копир." переключает галочку — включена ли строка
        # в копирование в буфер обмена.
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":  # "include" — первый столбец
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        self.toggle_row_include(iid)

    def on_tree_space(self, event):
        # Пробел переключает галочку у выделенной строки (строк) — удобно
        # при навигации стрелочками, без необходимости кликать мышью.
        selected = self.tree.selection()
        if not selected:
            focused = self.tree.focus()
            selected = (focused,) if focused else ()
        for iid in selected:
            self.toggle_row_include(iid)
        return "break"

    def _install_date_validation(self, entry):
        # Разрешаем печатать только цифры — точки в формате ДД.ММ.ГГГГ
        # расставляются автоматически при вводе, вручную их набрать нельзя.
        vcmd = (entry.register(self._validate_date_keystroke), "%d", "%S")
        entry.configure(validate="key", validatecommand=vcmd)
        entry.bind("<KeyRelease>", self._reformat_date_entry)

    @staticmethod
    def _validate_date_keystroke(action, inserted):
        if action != "1":  # не блокируем удаление
            return True
        return inserted.isdigit()

    def _reformat_date_entry(self, event=None):
        digits = re.sub(r"\D", "", self.date_var.get())[:8]
        if len(digits) > 4:
            formatted = f"{digits[:2]}.{digits[2:4]}.{digits[4:]}"
        elif len(digits) > 2:
            formatted = f"{digits[:2]}.{digits[2:]}"
        else:
            formatted = digits
        if formatted != self.date_var.get():
            self.date_var.set(formatted)
            self.date_entry.icursor(tk.END)

    def open_calendar_picker(self):
        try:
            initial_date = datetime.datetime.strptime(self.date_var.get().strip(), "%d.%m.%Y").date()
        except ValueError:
            initial_date = datetime.date.today()

        t = THEMES[self.theme]

        dialog = tk.Toplevel(self.root)
        dialog.title("Выберите дату")
        dialog.resizable(False, False)
        dialog.configure(bg=t["bg"])
        dialog.transient(self.root)

        cal = Calendar(
            dialog, date_pattern="dd.mm.yyyy", selectmode="day",
            year=initial_date.year, month=initial_date.month, day=initial_date.day,
            normalbackground=t["entry_bg"], normalforeground=t["fg"],
            weekendbackground=t["bg_alt"], weekendforeground=t["fg"],
            othermonthbackground=t["bg"], othermonthforeground=t["border"],
            othermonthwebackground=t["bg"], othermonthweforeground=t["border"],
            headersbackground=t["bg_alt"], headersforeground=t["fg"],
            selectbackground=t["accent"], selectforeground=t["accent_fg"],
            bordercolor=t["border"], background=t["bg_alt"], foreground=t["fg"],
        )
        cal.pack(padx=10, pady=10)

        def on_select(event=None):
            self.date_var.set(cal.get_date())
            dialog.destroy()

        cal.bind("<<CalendarSelected>>", on_select)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

        dialog.update_idletasks()
        x = self.date_entry.winfo_rootx()
        y = self.date_entry.winfo_rooty() + self.date_entry.winfo_height()
        dialog.geometry(f"+{x}+{y}")
        dialog.grab_set()

    def clear_all(self):
        self.path_entry.delete(0, tk.END)
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.current_rows = []
        self.current_kind = None
        self.type_var.set("")
        self.type_combo.configure(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")
        self.log_count = 0
        self.toggle_log(expanded=False)
        self.status_var.set("")

    def on_drop(self, event):
        path = event.data.strip().strip("{}")
        self.path_entry.delete(0, tk.END)
        self.path_entry.insert(0, path)
        self.run_parse()

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Выберите файл заказа",
            filetypes=[
                ("Библиотека Базис / PDF-эскиз", "*.bln *.pdf"),
                ("Библиотека Базис", "*.bln"),
                ("PDF-эскиз inSight", "*.pdf"),
                ("Все файлы", "*.*"),
            ],
        )
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path)
            self.run_parse()

    def on_type_change(self, event=None):
        new_type = self.type_var.get()
        for iid, row in zip(self.tree.get_children(), self.current_rows):
            row["type"] = new_type
            vals = list(self.tree.item(iid, "values"))
            vals[4] = new_type  # индекс столбца "type"
            self.tree.item(iid, values=vals)

    def run_parse(self):
        path = self.path_entry.get().strip()
        if not path:
            messagebox.showwarning("Нет файла", "Выберите файл .bln или .pdf.")
            return
        if not os.path.isfile(path):
            messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
            return

        ext = os.path.splitext(path)[1].lower()
        if ext == ".bln":
            kind = "bazis"
        elif ext == ".pdf":
            kind = "pdf"
        else:
            messagebox.showerror(
                "Неподдерживаемый файл",
                f'Расширение "{ext}" не поддерживается. Нужен файл .bln (Базис) или .pdf (inSight).',
            )
            return

        for row in self.tree.get_children():
            self.tree.delete(row)
        self.current_rows = []
        self.current_kind = kind

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")
        self.log_count = 0
        self.toggle_log(expanded=False)

        try:
            if kind == "bazis":
                order_number, results, warnings = parse_bln_sketches(path)
                row_type = "Bazis"
                self.type_var.set(row_type)
                self.type_combo.configure(state="readonly")
            else:
                order_number, results, warnings = parse_pdf_sketches(path)
                row_type = detect_pdf_source_type(os.path.basename(path))
                if row_type is None:
                    row_type = "inSight"
                    warnings.append(
                        f'Не удалось определить тип по имени файла "{os.path.basename(path)}" '
                        f'— установлен "{row_type}" по умолчанию. Проверьте и при необходимости '
                        f'смените в выпадающем списке "Тип" выше.'
                    )
                self.type_var.set(row_type)
                self.type_combo.configure(state="readonly")
        except CfbReadError as e:
            messagebox.showerror("Не удалось разобрать файл", str(e))
            self.status_var.set("Ошибка при разборе файла.")
            return
        except Exception as e:
            messagebox.showerror("Непредвиденная ошибка", str(e))
            self.status_var.set("Ошибка при разборе файла.")
            return

        for w in warnings:
            self.log(f"⚠ {w}")

        missing_marker = "Не найдено" if kind == "pdf" else ""
        current_date = self.date_var.get()

        for item in results:
            order_for_row = item["order_from_content"] or order_number or missing_marker
            part_for_row = item["part_code"] or missing_marker
            is_assembly = item.get("is_assembly", False)
            auto_exclude = item.get("auto_exclude", is_assembly)
            row = {
                "date": current_date,
                "order": order_for_row,
                "part": part_for_row,
                "type": row_type,
                "edge": "/",
                "material": item.get("material") or "",
                "description": item["description"],
                "page": str(item["page"]) if "page" in item else "-",
                "folder": item["source_dir"] or item.get("raw_name", ""),
                "is_assembly": is_assembly,
                "include": not auto_exclude,  # сборочные чертежи и материалы 3мм по умолчанию не копируем
            }
            self.tree.insert("", tk.END, values=(
                "☑" if row["include"] else "☐",
                row["date"], row["order"], row["part"], row["type"],
                row["edge"], row["material"], row["description"], row["page"], row["folder"],
            ), tags=("assembly",) if auto_exclude else ())
            self.current_rows.append(row)

        self.status_var.set(f"Найдено эскизов: {len(results)}")

    def copy_for_table(self):
        if not self.current_rows:
            messagebox.showinfo("Нечего копировать", "Сначала разберите файл.")
            return
        date_start = self.date_var.get().strip()
        lines = []
        skipped = 0
        for row in self.current_rows:
            if not row.get("include", True):
                skipped += 1
                continue
            # B=дата запуска, C=заказ, D=деталь, E=(пусто), F=тип, G=кромка
            lines.append("\t".join([
                date_start, row["order"], row["part"], "", row["type"], row["edge"],
            ]))

        if not lines:
            messagebox.showinfo(
                "Нечего копировать",
                "Все строки сняты с копирования (галочка «Копир.» снята).",
            )
            return

        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.log(f"Выделенных строк скопировано: {len(lines)}")
        self.status_var.set(f"Скопировано в буфер: {len(lines)} строк")
        self.flash_copy_button()

    def flash_copy_button(self):
        if self._copy_btn_reset_job is not None:
            self.root.after_cancel(self._copy_btn_reset_job)
        self.copy_btn.configure(bg="#4CAF50", text="Скопировано ✓")
        self._copy_btn_reset_job = self.root.after(900, self.reset_copy_button)

    def reset_copy_button(self):
        bg, fg = self._copy_btn_colors()
        self.copy_btn.configure(bg=bg, fg=fg, text=self.COPY_BTN_TEXT)
        self._copy_btn_reset_job = None


if __name__ == "__main__":
    RootClass = TkinterDnD.Tk if HAS_DND else tk.Tk
    root = RootClass()
    app = SketchExtractorApp(root)
    root.mainloop()
