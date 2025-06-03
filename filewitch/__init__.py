"""
FileWitch - A Python library for converting files between different formats.
"""

from .convert import (
    csv_to_xlsx,
    xlsx_to_csv,
    txt_to_docx,
    docx_to_txt,
    txt_to_pdf,
    docx_to_pdf,
    pptx_to_docx,
    copy_file,
    ConversionError
)

__version__ = "0.2.1"
__all__ = [
    'csv_to_xlsx',
    'xlsx_to_csv',
    'txt_to_docx',
    'docx_to_txt',
    'txt_to_pdf',
    'docx_to_pdf',
    'pptx_to_docx',
    'copy_file',
    'ConversionError'
] 