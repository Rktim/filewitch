"""
FileWitch - A Python library for converting and merging files between different formats.
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

from .merge import (
    merge_files,
    merge_pdfs,
    merge_csvs,
    merge_excels,
    merge_txts,
    merge_docxs
)

__version__ = "0.3.0"
__all__ = [
    'csv_to_xlsx',
    'xlsx_to_csv',
    'txt_to_docx',
    'docx_to_txt',
    'txt_to_pdf',
    'docx_to_pdf',
    'pptx_to_docx',
    'copy_file',
    'ConversionError',
    'merge_files',
    'merge_pdfs',
    'merge_csvs',
    'merge_excels',
    'merge_txts',
    'merge_docxs'
]
