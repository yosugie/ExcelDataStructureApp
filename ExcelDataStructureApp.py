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
    pip install pdfplumber customtkinter
    (опционально, для drag-and-drop) pip install tkinterdnd2
    (опционально, для кнопки-календаря у поля даты) pip install tkcalendar
    python ExcelDataStructureApp.py
"""

import os
import re
import sys
import struct
import ctypes
import base64
import tempfile
import datetime
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog

import customtkinter as ctk

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
# отдельный файл рядом со скриптом/exe) — плоская пиктограмма-таблица,
# перекрашенная в фирменный синий акцент (см. THEMES["dark"]["accent"]).
ICON_PNG_32 = (
    "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAqklEQVR4nO2XMQ6AIAxFS+Pqzux1PK/Xce7uAXQyaaAVJdrGpH8in8B/BEogAVPOeQcDEVE622gdXmahdXgJkTzCuQbJ3Ob1k7BxmSoPS+OrcG3uCqAlaRVXfkuPAd5WAASAeA8AXJ/qnkrQylsF0AaMyyT2aX5L7lsQAO4AUQXuWxAAAfA/AK3Ueh+zFUDv4/KOpLndPybIP4rWIqKEZ8MjHICdAUsInnUAoO9FEN13oUgAAAAASUVORK5CYII="
)
ICON_PNG_64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAABX0lEQVR4nO2bwQ6DIBBEV9Jr7577O/3e/k7P3PsB7YnEoOIuqU7YnXcUtpkZBQrRSRrM8/xttY9Cznnaa9ts8GK8ZiuIVF/wal5k21s66uCN2mPaa/DM0muqL0SheF7NAdGYIt79JTdL58/zfZaOv3J/PdR91UNgFPMiNq2qAEYyX9BqPgxgRPMFjfZLVoH762Eal701PYRfBhkAWgAaBoAWgIYBoAWgYQBoAWgYAFoAmvABmA5ERGyHDf+otdZYd698AqwFPecD5S5aantqegj/BDAAtAA0DAAtAA0DQAtAwwDQAtCED4CbIVNvh3AzdOqvDwADQAtAwwDQAtAwALQANAwALQANA0ALQGPeC/TQ83/+qhc0D5+AK15WPAuNdtUQGDEErWb1HDBSCBat4T+YSK2PCr2Tc57CL4NJpP1pqVeK51RfiMDSa9pr8ErtcTUHeA5hy1vTrJclsnVTfytJaxg5UXGGAAAAAElFTkSuQmCC"
)

# Multi-size .ico (16/32/48/256) для панели задач Windows — CustomTkinter там
# подменяет иконку окна на свою дефолтную через 200мс после создания, если к
# этому моменту не вызван iconbitmap()/wm_iconbitmap() — см. _apply_window_icon.
ICON_ICO = base64.b64decode(
    "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"
)

_icon_ico_path = None


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


def _apply_window_icon(window, icon_images):
    """Ставит и PNG (кроссплатформенно через iconphoto), и .ico (для панели
    задач Windows). У каждого top-level окна (диалоги в том числе) — своё
    состояние иконки, поэтому вызывается для главного окна И для каждого
    диалога отдельно, сразу после создания — синхронно, до того как
    CustomTkinter на Windows подменит иконку своей дефолтной через 200мс."""
    try:
        window.iconphoto(True, *icon_images)
    except tk.TclError:
        pass
    if sys.platform == "win32":
        global _icon_ico_path
        if _icon_ico_path is None:
            _icon_ico_path = os.path.join(tempfile.gettempdir(), "exceldatastructureapp_icon.ico")
            with open(_icon_ico_path, "wb") as f:
                f.write(ICON_ICO)
        try:
            window.iconbitmap(_icon_ico_path)
        except tk.TclError:
            pass


def fix_clipboard_shortcuts(widget):
    """Чинит Ctrl+C/V/X и Ctrl+A в полях ввода/тексте независимо от раскладки клавиатуры."""

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
            if isinstance(widget, ctk.CTkTextbox):
                widget.tag_add("sel", "1.0", "end")
            elif isinstance(widget, ctk.CTkEntry):
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
        mini_sector_shift = struct.unpack_from("<H", header, 32)[0]
        first_dir_sector = struct.unpack_from("<I", header, 48)[0]
        self.mini_stream_cutoff = struct.unpack_from("<I", header, 56)[0]
        first_minifat_sector = struct.unpack_from("<I", header, 60)[0]
        first_difat_sector = struct.unpack_from("<I", header, 68)[0]
        self.sector_size = 1 << sector_shift
        self.mini_sector_size = 1 << mini_sector_shift

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

        # Мини-FAT сам адресуется обычными секторами через self.fat, как и
        # директория ниже — только его содержимое читается как цепочка
        # мини-секторов (см. _read_mini_chain).
        minifat_raw = self._read_chain(first_minifat_sector)
        self.minifat = list(struct.unpack_from("<%dI" % (len(minifat_raw) // 4), minifat_raw, 0))

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

        # Потоки короче mini_stream_cutoff (обычно 4096 байт) лежат не в
        # обычных секторах, а в "мини-стриме" — единой цепочке из секторов
        # Root Entry, разбитой на мини-сектора (обычно 64 байта). Без этого
        # чтения короткие потоки (например "$$Lib_structure$$" в маленьких
        # библиотеках) читались бы как случайный мусор из середины файла.
        root_start_sec, root_size = self.index.get("Root Entry", (ENDOFCHAIN, 0))
        self.ministream = self._read_chain(root_start_sec, root_size)

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

    def _read_mini_chain(self, start_sec, size):
        chunks = []
        sec = start_sec
        seen = set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < 0xFFFFFFF0:
            if sec in seen:
                break
            seen.add(sec)
            off = sec * self.mini_sector_size
            chunks.append(self.ministream[off: off + self.mini_sector_size])
            sec = self.minifat[sec] if sec < len(self.minifat) else ENDOFCHAIN
        return b"".join(chunks)[:size]

    def get_stream(self, name):
        if name not in self.index:
            return None
        start_sec, size = self.index[name]
        if name != "Root Entry" and size < self.mini_stream_cutoff:
            return self._read_mini_chain(start_sec, size)
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

# Маркировка стандартной готовой заготовки рядом со штрихкодом на эскизе,
# например "R061" — скрытый штрихкод-текст вида *R061*, по аналогии с
# *ORD######*. На такие детали программу не пишут и в таблицу их не заносят,
# поэтому они, как сборочные чертежи и материалы 3мм, идут со снятой
# галочкой копирования по умолчанию.
STANDARD_BLANK_RE = re.compile(r"\*R\d+\*")


def is_standard_blank(text):
    return bool(STANDARD_BLANK_RE.search(text))

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


# Пометка "Смотри ДОП. Эскиз" посреди листа — деталь начерчена на отдельном
# "дополнительном" эскизе, не на этом листе. Обычно следующей строкой идёт
# логин конструктора, например "a.rohin" (в остальном тексте отчёта тоже
# встречаются такие логины, например "e.orhoyan"). Логин захватывается,
# только если следующая строка похожа на него (начинается с буквы) — иначе
# это уже нумерованный список кромок ("1 -..."), и имя не подставляется.
EXTRA_SKETCH_RE = re.compile(r"Смотри\s+ДОП\.\s*Эскиз(?:\s*\n\s*([A-Za-zА-Яа-яЁё][\w.\-]*))?")


def extract_extra_sketch_note(text: str):
    m = EXTRA_SKETCH_RE.search(text)
    if not m:
        return None
    name = m.group(1)
    return name if name else "Смотри ДОП. Эскиз"


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
            extra_sketch = extract_extra_sketch_note(text) if text.strip() else None
            thickness_mm = material_thickness_mm(material)
            is_too_thin = thickness_mm in NOT_MACHINABLE_THICKNESS_MM
            is_excluded_name = is_excluded_part_name(part_name)
            is_blank = is_standard_blank(text)

            if order:
                order_votes[order] = order_votes.get(order, 0) + 1

            results.append({
                "order_from_content": order,
                "part_code": part,
                "description": part_name or "",
                "material": material,
                "extra_sketch": extra_sketch,
                "source_dir": "",
                "raw_name": os.path.basename(pdf_path),
                "auto_exclude": is_too_thin or is_excluded_name or is_blank,
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
            f"Найдено деталей, которые не идут в работу (толщина {thickness_list} мм, "
            f"особое название вроде «Рейка на щит», либо маркировка стандартной заготовки "
            f"вроде «R061»): {n_not_machinable} — показаны в таблице для проверки, галочка "
            f"копирования снята по умолчанию."
        )

    if not results:
        warnings.append("В этом PDF не найдено ни одной страницы.")

    return order_number, results, warnings


# ---------------------------------------------------------------------------
# Интерфейс
# ---------------------------------------------------------------------------

THEMES = {
    "dark": dict(
        bg="#131316",           # фон окна
        card="#1b1c22",         # фон карточек/полей
        input="#212228",        # фон полей ввода (чуть светлее card)
        border="#2a2b33",       # обводка карточек и полей
        text="#e6e6ea",         # основной текст
        muted="#8b8d98",        # второстепенный текст (подписи, статус)
        accent="#0b63f6",       # фирменный синий — акцентные кнопки
        accent_hover="#3c82f8", # акцент при наведении
        accent_text="#ffffff",  # текст на акцентных кнопках
        danger="#e5484d",       # деструктивные/отменяющие кнопки (пока не используется)
        danger_hover="#c93d41",
    ),
    "light": dict(
        bg="#f4f5f8",
        card="#ffffff",
        input="#eef1f6",
        border="#dde2ea",
        text="#15171c",
        muted="#6b7280",
        accent="#2563eb",       # чуть темнее, чем в тёмной теме — иначе выцветает на белом
        accent_hover="#1d4ed8",
        accent_text="#ffffff",
        danger="#dc2626",
        danger_hover="#b91c1c",
    ),
}

# Лог (успех/ошибка) — фиксированные цвета, не завязаны на тему.
SUCCESS_COLOR = "#22c55e"
ERROR_COLOR = "#ef4444"

ctk.set_appearance_mode("dark")  # стартовая тема; переключение — вручную, см. toggle_theme()


class MessageDialog(ctk.CTkToplevel):
    """Свой диалог вместо tkinter.messagebox — тот рисуется ОС, а не Tk, и
    полностью игнорирует тёмную/светлую тему приложения (на Windows всегда
    светлый). Только текст + кнопка OK, без иконки-бейджа (кружок с "✕"/"i"
    пробовали — выглядело как лишняя некрасивая кнопка, убрали)."""

    def __init__(self, master, colors, icon_images, is_dark, title, message):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.configure(fg_color=colors["bg"])
        self.transient(master)
        _apply_window_icon(self, icon_images)
        set_windows_dark_titlebar(self, dark=is_dark)

        content = ctk.CTkFrame(
            self, fg_color=colors["card"], corner_radius=16,
            border_width=1, border_color=colors["border"],
        )
        content.grid(row=0, column=0, padx=16, pady=16)

        ctk.CTkLabel(
            content, text=message, text_color=colors["text"],
            wraplength=360, justify="center",
        ).grid(row=0, column=0, padx=32, pady=(28, 16))

        ctk.CTkButton(
            content, text="OK", command=self.destroy, width=110, height=32,
            corner_radius=20, fg_color=colors["accent"],
            hover_color=colors["accent_hover"], text_color=colors["accent_text"],
        ).grid(row=1, column=0, pady=(0, 24))

        self.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        self.wait_window()


class SketchExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("Парсер учёта эскизов")
        root.geometry("1300x1020")

        self._icon_imgs = [
            tk.PhotoImage(data=ICON_PNG_32),
            tk.PhotoImage(data=ICON_PNG_64),
        ]
        _apply_window_icon(root, self._icon_imgs)

        self.theme = "dark"
        # Реестр CTk-виджетов для мгновенного переключения темы без мигания
        # (см. apply_theme) — не вызываем ctk.set_appearance_mode() на
        # каждое переключение, это перерисовывает всё окно целиком.
        # CustomTkinter сам подхватывает смену цвета только для ПРЯМЫХ
        # детей обычного (не-CTk) окна — вложенные CTk-в-CTk виджеты сами
        # не перекрашиваются, поэтому bg_color всех таких виджетов тоже
        # приходится обновлять вручную через реестр (см. apply_theme).
        self._themed = []
        t = THEMES[self.theme]  # стартовые цвета — тема "dark" по умолчанию

        # --- Шапка: заголовок приложения + переключатель темы ---
        top_bar = ctk.CTkFrame(root, fg_color=t["bg"])
        top_bar.pack(fill="x", padx=8, pady=(8, 4))
        self._reg(top_bar, "plain_frame", surface="bg")

        title_label = ctk.CTkLabel(
            top_bar, text="Парсер учёта эскизов", font=ctk.CTkFont(size=16, weight="bold"),
        )
        title_label.pack(side="left")
        self._reg(title_label, "label", surface="bg")

        self.theme_toggle_btn = ctk.CTkButton(
            top_bar, text="Тёмная", command=self.toggle_theme,
            height=32, width=100, corner_radius=20,
        )
        self.theme_toggle_btn.pack(side="right")
        self._reg(self.theme_toggle_btn, "secondary_button", surface="bg")

        # --- Карточка "Основное": файл, тип эскиза, дата запуска ---
        main_card = ctk.CTkFrame(root, corner_radius=16, border_width=1, fg_color=t["card"], border_color=t["border"])
        main_card.pack(fill="x", padx=8, pady=4)
        self._reg(main_card, "card", surface="bg")

        card_title = ctk.CTkLabel(main_card, text="Основное", font=ctk.CTkFont(size=13, weight="bold"))
        card_title.pack(anchor="w", padx=16, pady=(14, 4))
        self._reg(card_title, "label")

        label_width = 190

        file_row = ctk.CTkFrame(main_card, fg_color=t["card"])
        file_row.pack(fill="x", padx=16, pady=4)
        self._reg(file_row, "plain_frame")
        file_label = ctk.CTkLabel(file_row, text="Файл (.bln или .pdf):", width=label_width, anchor="w")
        file_label.pack(side="left")
        self._reg(file_label, "label")
        self.path_entry = ctk.CTkEntry(file_row)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._reg(self.path_entry, "entry")
        fix_clipboard_shortcuts(self.path_entry)
        self.path_entry.bind("<Return>", lambda event: self.run_parse())
        if HAS_DND:
            self.path_entry.drop_target_register(DND_FILES)
            self.path_entry.dnd_bind("<<Drop>>", self.on_drop)
        browse_btn = ctk.CTkButton(
            file_row, text="Обзор...", command=self.browse_file, height=32, width=90, corner_radius=20,
        )
        browse_btn.pack(side="left")
        self._reg(browse_btn, "secondary_button")

        type_row = ctk.CTkFrame(main_card, fg_color=t["card"])
        type_row.pack(fill="x", padx=16, pady=4)
        self._reg(type_row, "plain_frame")
        type_label = ctk.CTkLabel(type_row, text="Тип эскиза:", width=label_width, anchor="w")
        type_label.pack(side="left")
        self._reg(type_label, "label")
        self.type_var = tk.StringVar(value="")
        self.type_combo = ctk.CTkComboBox(
            type_row, variable=self.type_var, values=list(SOURCE_TYPES),
            state="disabled", width=180, command=self.on_type_change,
        )
        self.type_combo.pack(side="left")
        self._reg(self.type_combo, "combobox")

        date_row = ctk.CTkFrame(main_card, fg_color=t["card"])
        date_row.pack(fill="x", padx=16, pady=(4, 14))
        self._reg(date_row, "plain_frame")
        date_label = ctk.CTkLabel(date_row, text="Дата запуска:", width=label_width, anchor="w")
        date_label.pack(side="left")
        self._reg(date_label, "label")
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%d.%m.%Y"))
        self.date_entry = ctk.CTkEntry(date_row, width=110, textvariable=self.date_var)
        self.date_entry.pack(side="left", padx=(0, 8))
        self._reg(self.date_entry, "entry")
        fix_clipboard_shortcuts(self.date_entry)
        self.date_var.trace_add("write", lambda *args: self.refresh_date_column())
        self._install_date_validation(self.date_entry)
        self.calendar_btn = None
        if HAS_TKCALENDAR:
            self.calendar_btn = ctk.CTkButton(
                date_row, text="📅", command=self.open_calendar_picker, height=32, width=40, corner_radius=20,
            )
            self.calendar_btn.pack(side="left", padx=(0, 8))
            self._reg(self.calendar_btn, "secondary_button")
        self.today_btn = ctk.CTkButton(
            date_row, text="Сегодня", command=self.set_today, height=32, width=90, corner_radius=20,
        )
        self.today_btn.pack(side="left")
        self._reg(self.today_btn, "secondary_button")

        # --- Статус-строка ---
        self.status_var = tk.StringVar(value="")
        self.status_label = ctk.CTkLabel(root, textvariable=self.status_var, anchor="w")
        self.status_label.pack(fill="x", padx=16, pady=(4, 6))
        self._reg(self.status_label, "muted_label", surface="bg")

        # --- Карточка с таблицей результатов (сам Treeview — не CustomTkinter,
        # в библиотеке нет виджета-таблицы; только карточка вокруг перекрашена) ---
        self.tree_frame = tree_frame = ctk.CTkFrame(root, corner_radius=16, border_width=1, fg_color=t["card"], border_color=t["border"])
        tree_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._reg(tree_frame, "card", surface="bg")

        columns = ("include", "page", "date", "order", "part", "description", "material", "edge", "extra_sketch", "type")
        self.columns = columns
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=16)
        headers = {
            "include": "Копир.",
            "page": "Страница",
            "date": "Дата запуска",
            "order": "№ заказа",
            "part": "№ детали",
            "description": "Описание",
            "material": "Материал",
            "edge": "Кромка",
            "extra_sketch": "Доп. эскиз",
            "type": "Тип эскиза",
        }
        widths = {
            "include": 55, "page": 70, "date": 90, "order": 80, "part": 100,
            "description": 220, "material": 260, "edge": 60, "extra_sketch": 150, "type": 100,
        }
        for c in columns:
            self.tree.heading(c, text=headers[c], anchor="center")
            self.tree.column(c, width=widths[c], anchor="center")
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<space>", self.on_tree_space)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(14, 0), pady=(14, 0))
        vsb.grid(row=0, column=1, sticky="ns", pady=(14, 0), padx=(0, 14))
        hsb.grid(row=1, column=0, sticky="ew", padx=(14, 0), pady=(0, 14))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # --- Кнопки действий ---
        btn_frame = ctk.CTkFrame(root, fg_color=t["bg"])
        btn_frame.pack(fill="x", padx=8, pady=6)
        self._reg(btn_frame, "plain_frame", surface="bg")
        self.COPY_BTN_TEXT = "Скопировать строки"
        self.copy_btn = ctk.CTkButton(
            btn_frame, text=self.COPY_BTN_TEXT, command=self.copy_for_table,
            height=32, width=180, corner_radius=20,
        )
        self.copy_btn.pack(side="left")
        self._reg(self.copy_btn, "accent_button", surface="bg")
        self._copy_btn_reset_job = None

        self.clear_btn = ctk.CTkButton(
            btn_frame, text="Очистить", command=self.clear_all, height=32, width=110, corner_radius=20,
        )
        self.clear_btn.pack(side="left", padx=(8, 0))
        self._reg(self.clear_btn, "secondary_button", surface="bg")

        # --- Карточка "Журнал" (сворачиваемая, шеврон в шапке той же карточки) ---
        self.log_expanded = False
        self.log_count = 0

        log_card = ctk.CTkFrame(root, corner_radius=16, border_width=1, fg_color=t["card"], border_color=t["border"])
        log_card.pack(fill="x", padx=8, pady=(0, 8))
        self._reg(log_card, "card", surface="bg")

        self.log_header_var = tk.StringVar(value="⌄ Журнал")
        self.log_header = ctk.CTkLabel(
            log_card, textvariable=self.log_header_var, anchor="w",
            cursor="hand2", font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.log_header.pack(fill="x", padx=16, pady=10)
        self.log_header.bind("<Button-1>", lambda event: self.toggle_log())
        self._reg(self.log_header, "label")

        self.log_body = ctk.CTkFrame(log_card, fg_color=t["card"])
        self._reg(self.log_body, "plain_frame")
        self.log_text = ctk.CTkTextbox(self.log_body, height=120, corner_radius=16, border_width=1, font=("Consolas", 12))
        self.log_text.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.log_text.tag_config("warn", foreground=ERROR_COLOR)
        self.log_text.tag_config("success", foreground=SUCCESS_COLOR)
        self.log_text.configure(state="disabled")
        fix_clipboard_shortcuts(self.log_text)
        self._reg(self.log_text, "textbox")
        # log_body остаётся не упакованным — журнал стартует свёрнутым.

        self.current_rows = []  # список dict-ов с результатами разбора
        self.current_kind = None  # "bazis" или "pdf"

        self.apply_theme()

    def _reg(self, widget, kind, surface="card"):
        # surface — какой фон стоит НЕПОСРЕДСТВЕННО за виджетом ("bg" — фон
        # окна, "card" — фон карточки): нужно, чтобы вручную обновлять
        # bg_color виджетов, вложенных в другой CTk-виджет — CustomTkinter
        # сам подхватывает смену цвета только для прямых детей обычного
        # tk-окна, дальше по вложенности — нет (см. apply_theme).
        self._themed.append((widget, kind, surface))
        return widget

    def toggle_log(self, expanded=None):
        self.log_expanded = (not self.log_expanded) if expanded is None else expanded
        if self.log_expanded:
            self.log_body.pack(fill="both", expand=True, pady=(0, 4))
            arrow = "⌃"  # раскрыто — клик свернёт
        else:
            self.log_body.pack_forget()
            arrow = "⌄"  # свёрнуто — клик раскроет
        suffix = f" ({self.log_count})" if self.log_count else ""
        self.log_header_var.set(f"{arrow} Журнал{suffix}")

    def toggle_theme(self):
        self.theme = "light" if self.theme == "dark" else "dark"
        self.theme_toggle_btn.configure(text="Тёмная" if self.theme == "dark" else "Светлая")
        self.apply_theme()

    def apply_theme(self):
        t = THEMES[self.theme]
        self.root.configure(bg=t["bg"])
        set_windows_dark_titlebar(self.root, dark=(self.theme == "dark"))

        for widget, kind, surface in self._themed:
            surface_color = t["bg"] if surface == "bg" else t["card"]
            if kind == "label":
                widget.configure(text_color=t["text"], bg_color=surface_color)
            elif kind == "muted_label":
                widget.configure(text_color=t["muted"], bg_color=surface_color)
            elif kind == "card":
                widget.configure(fg_color=t["card"], border_color=t["border"], bg_color=surface_color)
            elif kind == "plain_frame":
                widget.configure(fg_color=surface_color)
            elif kind == "entry":
                widget.configure(fg_color=t["input"], border_color=t["border"], text_color=t["text"], bg_color=surface_color)
            elif kind == "combobox":
                widget.configure(
                    fg_color=t["input"], border_color=t["border"], text_color=t["text"],
                    button_color=t["input"], button_hover_color=t["border"],
                    dropdown_fg_color=t["card"], dropdown_text_color=t["text"],
                    dropdown_hover_color=t["input"], bg_color=surface_color,
                )
            elif kind == "accent_button":
                widget.configure(fg_color=t["accent"], hover_color=t["accent_hover"], text_color=t["accent_text"], bg_color=surface_color)
            elif kind == "secondary_button":
                widget.configure(
                    fg_color=t["card"], hover_color=t["input"], text_color=t["text"],
                    border_width=1, border_color=t["border"], bg_color=surface_color,
                )
            elif kind == "textbox":
                widget.configure(
                    fg_color=t["card"], border_color=t["border"], text_color=t["text"],
                    scrollbar_button_color=t["border"], scrollbar_button_hover_color=t["muted"],
                    bg_color=surface_color,
                )

        # Кнопка копирования — акцентная, но пока идёт "вспышка" успеха
        # (см. flash_copy_button), не возвращаем её к акценту досрочно.
        if self._copy_btn_reset_job is None:
            self.copy_btn.configure(fg_color=t["accent"], hover_color=t["accent_hover"], text_color=t["accent_text"])

        # ttk-виджеты (Treeview, Scrollbar) — не CustomTkinter, стилизуются через ttk.Style.
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview", background=t["card"], fieldbackground=t["card"],
            foreground=t["text"], borderwidth=0, rowheight=26,
        )
        style.map("Treeview", background=[("selected", t["accent"])], foreground=[("selected", t["accent_text"])])
        style.configure("Treeview.Heading", background=t["bg"], foreground=t["text"], relief="flat")
        style.map("Treeview.Heading", background=[("active", t["bg"])])
        style.configure(
            "TScrollbar", background=t["input"], troughcolor=t["bg"],
            bordercolor=t["border"], arrowcolor=t["accent"],
        )
        style.map("TScrollbar", arrowcolor=[("pressed", t["accent"]), ("active", t["accent"])])
        self.tree.tag_configure("assembly", background=t["input"], foreground=t["muted"])

    def show_message(self, title, message):
        MessageDialog(self.root, THEMES[self.theme], self._icon_imgs, self.theme == "dark", title, message)

    def log(self, message):
        tag = "warn" if message.startswith("⚠") else "success"
        self.log_text.configure(state="normal")
        prefix = "\n" if self.log_text.get("1.0", "end-1c") else ""
        self.log_text.insert(tk.END, prefix + message + "\n", tag)
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
        idx = self.columns.index("date")
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            vals[idx] = new_date
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
        _apply_window_icon(dialog, self._icon_imgs)
        set_windows_dark_titlebar(dialog, dark=(self.theme == "dark"))

        cal = Calendar(
            dialog, date_pattern="dd.mm.yyyy", selectmode="day",
            year=initial_date.year, month=initial_date.month, day=initial_date.day,
            normalbackground=t["input"], normalforeground=t["text"],
            weekendbackground=t["bg"], weekendforeground=t["text"],
            othermonthbackground=t["bg"], othermonthforeground=t["border"],
            othermonthwebackground=t["bg"], othermonthweforeground=t["border"],
            headersbackground=t["bg"], headersforeground=t["text"],
            selectbackground=t["accent"], selectforeground=t["accent_text"],
            bordercolor=t["border"], background=t["bg"], foreground=t["text"],
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

    def on_type_change(self, choice=None):
        new_type = self.type_var.get()
        idx = self.columns.index("type")
        for iid, row in zip(self.tree.get_children(), self.current_rows):
            row["type"] = new_type
            vals = list(self.tree.item(iid, "values"))
            vals[idx] = new_type
            self.tree.item(iid, values=vals)

    def run_parse(self):
        path = self.path_entry.get().strip()
        if not path:
            self.show_message("Нет файла", "Выберите файл .bln или .pdf.")
            return
        if not os.path.isfile(path):
            self.show_message("Ошибка", f"Файл не найден:\n{path}")
            return

        ext = os.path.splitext(path)[1].lower()
        if ext == ".bln":
            kind = "bazis"
        elif ext == ".pdf":
            kind = "pdf"
        else:
            self.show_message(
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
            self.show_message("Не удалось разобрать файл", str(e))
            self.status_var.set("Ошибка при разборе файла.")
            return
        except Exception as e:
            self.show_message("Непредвиденная ошибка", str(e))
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
                "extra_sketch": item.get("extra_sketch") or "-",
                "page": str(item["page"]) if "page" in item else "-",
                "is_assembly": is_assembly,
                "include": not auto_exclude,  # сборочные чертежи и материалы 3мм по умолчанию не копируем
            }
            self.tree.insert("", tk.END, values=(
                "☑" if row["include"] else "☐",
                row["page"], row["date"], row["order"], row["part"],
                row["description"], row["material"], row["edge"], row["extra_sketch"],
                row["type"],
            ), tags=("assembly",) if auto_exclude else ())
            self.current_rows.append(row)

        self.status_var.set(f"Найдено эскизов: {len(results)}")

    def copy_for_table(self):
        if not self.current_rows:
            self.show_message("Нечего копировать", "Сначала разберите файл.")
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
            self.show_message(
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
        self.copy_btn.configure(fg_color=SUCCESS_COLOR, hover_color=SUCCESS_COLOR, text="Скопировано ✓")
        self._copy_btn_reset_job = self.root.after(900, self.reset_copy_button)

    def reset_copy_button(self):
        t = THEMES[self.theme]
        self.copy_btn.configure(fg_color=t["accent"], hover_color=t["accent_hover"], text=self.COPY_BTN_TEXT)
        self._copy_btn_reset_job = None


if __name__ == "__main__":
    RootClass = TkinterDnD.Tk if HAS_DND else tk.Tk
    root = RootClass()
    app = SketchExtractorApp(root)
    root.mainloop()
