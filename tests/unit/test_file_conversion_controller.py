"""Tests for the file conversion controller."""

import pathlib

import docx
import pytest
from docx.text import run

from ctk_functions.functions.file_conversion import controller


def is_run_font_color_red(run: run.Run) -> bool:
    """Checks if the font color of a run is red."""
    return '<w:color w:val="FF0000"' in str(run.font.color.element.xml)


def test_mark_warnings_as_red(tmp_path: pathlib.Path) -> None:
    """Tests marking warning labels as red."""
    filename = tmp_path / "test.docx"
    doc = docx.Document()
    doc.add_paragraph("This is a {{!WARNING-TEXT}}.")
    doc.add_paragraph("Another {{!WARNING-TEXT-2}} here.")
    doc.save(str(filename))

    controller.mark_warnings_as_red(filename)
    modified_doc = docx.Document(str(filename))

    assert modified_doc.paragraphs[0].runs[1].text == "{{!WARNING-TEXT}}"
    assert modified_doc.paragraphs[1].runs[1].text == "{{!WARNING-TEXT-2}}"
    assert not is_run_font_color_red(modified_doc.paragraphs[0].runs[0])
    assert is_run_font_color_red(modified_doc.paragraphs[0].runs[1])
    assert is_run_font_color_red(modified_doc.paragraphs[1].runs[1])


@pytest.mark.asyncio
async def test_markdown2docx(tmp_path: pathlib.Path) -> None:
    """Tests the conversion of Markdown to docx."""
    markdown = "# Header\n\nThis is a paragraph.\n\nThis is a {{!WARNING-TEXT}}."

    docx_bytes = await controller.markdown2docx(markdown)
    filename = tmp_path / "test.docx"
    with open(filename, "wb") as file:
        file.write(docx_bytes)
    doc = docx.Document(str(filename))

    assert doc.paragraphs[0].text == "Header"
    assert doc.paragraphs[1].text == "This is a paragraph."
    assert doc.paragraphs[2].runs[1].text == "{{!WARNING-TEXT}}"
    assert not is_run_font_color_red(doc.paragraphs[2].runs[0])
    assert is_run_font_color_red(doc.paragraphs[2].runs[1])
