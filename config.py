"""Configuration values for the Word merge tool.

Keep operational defaults, Word COM numeric constants, and sample document
content here so app.py can stay focused on workflow code.
"""

from __future__ import annotations


SUPPORTED_EXTENSIONS = {".doc", ".docx", ".docs"}
LEGACY_DOC_EXTENSION = ".doc"
DOCX_EXTENSION = ".docx"

DEFAULT_KYRGYZ_DIR = "kyrgyz_documents"
DEFAULT_RUSSIAN_DIR = "russian_documents"
DEFAULT_OUTPUT_DIR = "merged_documents"
PREPARED_DOCX_TEMP_PREFIX = "word_merge_prepared_"
KYRGYZ_LANGUAGE_KEY = "kyrgyz"
RUSSIAN_LANGUAGE_KEY = "russian"

BACKEND_AUTO = "auto"
BACKEND_WORD = "word"
BACKEND_DOCX = "docx"
BACKEND_CHOICES = (BACKEND_AUTO, BACKEND_WORD, BACKEND_DOCX)
DEFAULT_BACKEND = BACKEND_AUTO
LOG_FORMAT = "%(levelname)s: %(message)s"

MERGED_OUTPUT_SUFFIX = "_merged.docx"
DEFAULT_SAFE_OUTPUT_STEM = "merged"
SAFE_OUTPUT_NAME_PATTERN = r"[^A-Za-zА-Яа-яЁё0-9_.-]+"

WORD_APPLICATION_NAME = "Word.Application"
WORD_FORMAT_DOC = 0
WORD_FORMAT_XML_DOCUMENT = 12
WORD_DO_NOT_SAVE_CHANGES = 0
WORD_DISPLAY_ALERTS_NONE = 0
WORD_HEADER_FOOTER_INDEXES = (1, 2, 3)
WORD_INFO_WITHIN_TABLE = 12
WORD_AUTOFIT_FIXED = 0
WORD_PREFERRED_WIDTH_POINTS = 3
WORD_NUMBER_ALL_NUMBERS = 3

HEADER_STORY = "header"
FOOTER_STORY = "footer"
HEADER_FOOTER_SEPARATOR = " / "
MULTI_SECTION_TEXT_SEPARATOR = " | "

MERGED_TABLE_ROWS = 1
MERGED_TABLE_COLUMNS = 2
TWO_COLUMN_WIDTH_RATIO = 0.5
WORD_TABLE_ROW = 1
WORD_KYRGYZ_COLUMN = 1
WORD_RUSSIAN_COLUMN = 2
DOCX_TABLE_ROW = 0
DOCX_KYRGYZ_COLUMN = 0
DOCX_RUSSIAN_COLUMN = 1
TABLE_GRID_STYLE = "Table Grid"

DUMMY_KYRGYZ_DOCX = "test.docx"
DUMMY_RUSSIAN_DOCX = "test.docx"
DUMMY_LEGACY_DOC = "legacy.doc"

DUMMY_KYRGYZ = {
    "header": "Кыргыз документ башы",
    "footer": "Кыргыз документ аягы",
    "title": "Кыргызча үлгү документ",
    "lead": "Бул абзац майда форматтоону, калың жана курсив текстти текшерет.",
    "table_values": (("Аталышы", "Маани"), ("Текшерүү", "Ийгиликтүү")),
    "accent_rgb": (0x1F, 0x6F, 0x5C),
}

DUMMY_RUSSIAN = {
    "header": "Заголовок русского документа",
    "footer": "Нижний колонтитул русского документа",
    "title": "Русский тестовый документ",
    "lead": "Этот абзац проверяет сохранение жирного, курсивного и цветного текста.",
    "table_values": (("Название", "Значение"), ("Проверка", "Успешно")),
    "accent_rgb": (0x7A, 0x2E, 0x2E),
}
