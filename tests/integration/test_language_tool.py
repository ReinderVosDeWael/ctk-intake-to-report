"""Tests the language tool."""

import pytest

from ctk_functions.microservices import language_tool


@pytest.mark.asyncio
async def test_lanugage_correcter() -> None:
    """Tests the language correcter."""
    correcter = language_tool.LanguageCorrecter()
    text = "They is good."

    corrections = await correcter.check(text)

    assert len(corrections) == 1
    assert corrections[0].rule.id == "PERS_PRONOUN_AGREEMENT"
