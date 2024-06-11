"""Utilities for correcting grammar and syntax."""

import asyncio
from typing import Sequence

import cmi_docx
import docx
import spacy
from docx import document

from ctk_functions.text import corrections

NLP = spacy.load("en_core_web_sm")

# c.f. https://community.languagetool.org/rule/list?lang=en for a list of rules.
# These are the rules that have been tested on existing intake reports and are
# used for the dynamic text. Rules applying only to static template text
# should be fixed in the template itself.

DEFAULT_LANGUAGE_RULES = [
    "BASE_FORM",
    "CONSECUTIVE_SPACES",
    "PERS_PRONOUN_AGREEMENT",
    "NON3PRS_VERB",
    "THE_US",
    "UPPERCASE_SENTENCE_START",
    "WEEK_HYPHEN",
]

ruleIds: set[str] = set()


class DocumentCorrections:
    """Corrects verb conjugations associated with 'they' in a Word document."""

    def __init__(
        self,
        document: document.Document,
        enabled_rules: Sequence[str] | None = DEFAULT_LANGUAGE_RULES,
    ) -> None:
        """Initializes the corrector with a document.

        Args:
            document: The docx document to correct.
            enabled_rules: The rules to enable for the correction. If None, all rules
                are enabled.

        """
        self.document = document
        self.correcter = corrections.LanguageCorrecter()
        self.enabled_rules = set(enabled_rules) if enabled_rules else set()

    async def correct(self) -> None:
        """Makes corrections based on the enabled and disabled rules."""
        for paragraph in self.document.paragraphs:
            await self._correct_paragraph(paragraph)

    async def _correct_paragraph(
        self, paragraph: docx.text.paragraph.Paragraph
    ) -> None:
        """Corrects conjugations in a single paragraph.

        Args:
            paragraph: The paragraph to correct.
        """
        sentences = list(NLP(paragraph.text).sents)
        new_sentences = await asyncio.gather(
            *[
                self.correcter.run(
                    sentence.text,
                    enabled_rules=self.enabled_rules,
                )
                for sentence in sentences
            ]
        )
        extended_pargraph = cmi_docx.ExtendParagraph(paragraph)
        for old, new in zip(sentences, new_sentences):
            if old.text != new:
                extended_pargraph.replace(old.text, new)
