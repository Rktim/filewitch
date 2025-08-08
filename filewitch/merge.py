"""
Command-line interface for FileWitch.
"""

import sys
import os
from pathlib import Path
import click
from .convert import (
    csv_to_xlsx,
    xlsx_to_csv,
    txt_to_docx,
    docx_to_txt,
    txt_to_pdf,
    docx_to_pdf,
    pptx_to_docx,
    pptx_to_pdf,
    copy_file,
    ConversionError
)
from .merge import merge_files


def get_extension(filename: str) -> str:
    """Get file extension without the dot."""
    return Path(filename).suffix[1:].lower()

def get_output_path(input_file: str, target_format: str) -> str:
    """Generate output file path with new extension."""
    return str(Path(input_file).with_suffix(f".{target_format}"))

def validate_conversion(ext: str, target_format: str) -> bool:
    """Validate if the conversion is supported."""
    supported_conversions = {
        ("csv", "xlsx"),
        ("xlsx", "csv"),
        ("txt", "docx"),
        ("docx", "txt"),
        ("txt", "pdf"),
        ("docx", "pdf"),
        ("pptx", "docx"),
        ("pptx", "pdf")
    }
    return (ext, target_format) in supported_conversions or ext == target_format

def convert(input_file: str, target_format: str) -> bool:
    """Convert file from one format to another."""
    if not os.path.exists(input_file):
        click.echo(f"❌ Error: Input file '{input_file}' does not exist.", err=True)
        return False

    ext = get_extension(input_file)
    output_file = get_output_path(input_file, target_format)

    if not validate_conversion(ext, target_format):
        click.echo(f"❌ Error: Conversion from {ext} to {target_format} is not supported.", err=True)
        click.echo("\nSupported conversions:")
        click.echo("  - CSV ↔ Excel")
        click.echo("  - Text ↔ Word")
        click.echo("  - Text → PDF")
        click.echo("  - Word → PDF")
        click.echo("  - PowerPoint → Word")
        click.echo("  - PowerPoint → PDF")
        click.echo("  - Same format copying (e.g., txt to txt, pdf to pdf)")
        return False

    try:
        if ext == target_format:
            copy_file(input_file, output_file)
        else:
            match (ext, target_format):
                case ("csv", "xlsx"):
                    csv_to_xlsx(input_file, output_file)
                case ("xlsx", "csv"):
                    xlsx_to_csv(input_file, output_file)
                case ("txt", "docx"):
                    txt_to_docx(input_file, output_file)
                case ("docx", "txt"):
                    docx_to_txt(input_file, output_file)
                case ("txt", "pdf"):
                    txt_to_pdf(input_file, output_file)
                case ("docx", "pdf"):
                    docx_to_pdf(input_file, output_file)
                case ("pptx", "docx"):
                    pptx_to_docx(input_file, output_file)
                case ("pptx", "pdf"):
                    pptx_to_pdf(input_file, output_file)
        
        click.echo(f"✅ Successfully converted: {output_file}")
        return True
        
    except ConversionError as e:
        click.echo(f"❌ Error during conversion: {str(e)}", err=True)
        return False
    except Exception as e:
        click.echo(f"❌ Unexpected error: {str(e)}", err=True)
        return False


@click.group()
def cli():
    """FileWitch - Convert or merge files between different formats."""
    pass


@cli.command(name='convert')
@click.argument('input_file', type=click.Path(exists=True))
@click.argument('target_format')
def convert_command(input_file: str, target_format: str):
    """Convert a file to another format."""
    if not convert(input_file, target_format):
        sys.exit(1)


@cli.command(name='merge')
@click.argument('output_file', type=click.Path())
@click.argument('input_files', nargs=-1, type=click.Path(exists=True))
def merge_command(output_file: str, input_files: list):
    """Merge multiple files into one."""
    if not input_files or len(input_files) < 2:
        click.echo("❌ Please provide at least two input files to merge.", err=True)
        sys.exit(1)
    try:
        merge_files(output_file, *input_files)
        click.echo(f"✅ Successfully merged files into: {output_file}")
    except Exception as e:
        click.echo(f"❌ Error during merging: {str(e)}", err=True)
        sys.exit(1)


def main():
    """Main entry point for the CLI."""
    cli()
