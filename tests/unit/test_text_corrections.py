"""Tests for the text corrections module."""

import pytest

from ctk_functions.text import corrections


@pytest.fixture(scope="module")
def correcter() -> corrections.LanguageCorrecter:
    """Fixture for the LanguageCorrecter class.

    The initialization is slow, so it is performed at the module level.
    """
    return corrections.LanguageCorrecter()


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "input_text, expected",
    [
        ("They is going to the store.", "They are going to the store."),
        ("they is going to the store.", "They are going to the store."),
        (
            "He is going to the store. they are walking the dog.",
            "He is going to the store. They are walking the dog.",
        ),
        ("The store to which they has gone.", "The store to which they have gone."),
        (
            (
                "A very long sentence because the context is clipped when the sentence"
                " is too long but the results should still be correct. They is going"
                " to the store."
            ),
            (
                "A very long sentence because the context is clipped when the sentence"
                " is too long, but the results should still be correct. They are going"
                " to the store."
            ),
        ),
        ("", ""),
    ],
)
async def test_text_corrections(
    correcter: corrections.LanguageCorrecter, input_text: str, expected: str
) -> None:
    """Tests the entrypoint of the TextCorrections class."""
    actual = await correcter.run(input_text)

    assert actual == expected
