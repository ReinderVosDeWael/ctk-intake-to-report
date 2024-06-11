"""Module for syntax and grammatical correctionss of text."""

from typing import Collection

import spacy

from ctk_functions.microservices import language_tool

NLP = spacy.load("en_core_web_sm")


class LanguageCorrecter:
    """Corrects the grammar and syntax of text."""

    def __init__(self) -> None:
        """Initializes the language tool."""
        self.language_tool = language_tool.LanguageCorrecter()

    async def run(
        self,
        text: str,
        enabled_rules: Collection[str] | None = None,
    ) -> str:
        """Corrects the text following the object's settings.

        Initializing the language tool takes a while, so it is performed at
        the class level. This, however, creates a race condition when multiple
        threads try to access the language tool at the same time. To prevent
        this, a lock is used to ensure that only one thread can access the
        language tool at a time.

        Args:
            text: The text to correct.
            enabled_rules: The rules to enable for the correction. If None, all rules
                are enabled.

        Returns:
            The corrected text.
        """

        async def get_corrections(text: str) -> list[language_tool.Correction]:
            corrections = await self.language_tool.check(
                text, enabled_rules=enabled_rules
            )

            return corrections

        corrections = await get_corrections(text)
        while corrections:
            text = self._apply_correction(corrections[-1], text)
            corrections = await get_corrections(text)
        return text

    @classmethod
    def _apply_correction(
        cls, correction: language_tool.Correction, full_text: str
    ) -> str:
        """Applies a correction to a text.

        Args:
            correction: The language tool Match.
            full_text: The full text to correct.
        """
        if len(correction.replacements) == 1:
            return (
                full_text[: correction.offset]
                + correction.replacements[0].value
                + full_text[correction.offset + correction.length :]
            )

        if correction.rule.id == "PERS_PRONOUN_AGREEMENT":
            return cls._resolve_pers_pronoun_agreement(correction, full_text)
        else:
            raise ValueError(f"Cannot resolve replacement {correction}.")

    @classmethod
    def _resolve_pers_pronoun_agreement(
        cls,
        correction: language_tool.Correction,
        full_text: str,
    ) -> str:
        """Resolves personal pronoun corrections with multiple replacements.

        Consult https://www.nltk.org/book/ch05.html for more information on NLTK tokens.

        Args:
            correction: The language tool Match.
            full_text: The full text to correct.

        Returns:
            The corrected text.
        """
        verb_tense = cls._get_verb_tense(full_text, correction.offset)

        # Example correction message: "Use a third-person plural verb with ‘they’.""

        subject = correction.message.split("‘")[1].split("’")[0].lower()

        if subject == "they":
            target_tense = "VBP" if verb_tense == "VBZ" else verb_tense
        elif verb_tense == "VBP":
            target_tense = "VBZ"
        else:
            target_tense = verb_tense

        for replacement in correction.replacements:
            new_sentence = (
                full_text[: correction.offset]
                + replacement.value
                + full_text[correction.offset + correction.length :]
            )
            new_verb_tense = cls._get_verb_tense(new_sentence, correction.offset)
            if new_verb_tense == target_tense:
                return new_sentence

        raise ValueError(f"Could not find a suitable replacement for {correction}.")

    @classmethod
    def _get_verb_tense(cls, sentence: str, verb_offset: int) -> str:
        """Gets the tense of a verb."""
        doc = NLP(sentence)
        verb = next(word for word in doc if word.idx == verb_offset)
        return verb.tag_
