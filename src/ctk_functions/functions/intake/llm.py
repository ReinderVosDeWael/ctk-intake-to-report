"""Large language model functionality for the intake module."""

import dataclasses
import enum
import uuid
from typing import Awaitable, Sequence

from ctk_functions.functions.intake.utils import string_utils
from ctk_functions.microservices import azure


@dataclasses.dataclass
class LlmPlaceholder:
    """Represents a placeholder for large language model input in the report.

    Attributes:
        id: The unique identifier for the placeholder. Will be inserted into
            the report as {{id}}.
        replacement: The replacement text for the placeholder. Provided as
            an awaitable to allow for asynchronous processing.
    """

    id: str
    replacement: Awaitable[str]


class Prompts(str, enum.Enum):
    """Prompts for the large language model."""

    parent_input = """
You will receive an excerpt of a clinical report with part of the text replaced
by a placeholder, a response by a parent, and the name and pronouns of the
child. Your task is to insert the parent's response into the excerpt. You should
return the excerpt in full with the placeholder replaced by the parent's
response. Do not use quotations; make sure the response is
integrated into the text. If the response is not provided, please write that the
response was not provided. If the response is not applicable, please write that
the information is not applicable. Parents' responses may be terse,
grammatically incorrect, or incomplete. Make sure that the response is clear, uses
correct grammar, avoids abbreviations and repetition, and is consistent with the
tone of a clinical report written by a medical professional. For lists that contain
recurring elements, feel free to merge the elements into a single unit if it makes
grammatical sense. Report dates in the format "DD Month YYYY" (e.g., 15 June 2022).
"""
    edit = """
You will receive an excerpt of a clinical report. Your task is to edit the text
to improve its clarity, grammar, and style. You should return the excerpt in
full with the necessary edits. Ensure that the tone is appropriate for a
clinical report written by a doctor, i.e. professional and objective. Do not use
quotations; make sure the response is integrated into the text. Do not alter the
content of the text; ONLY EDIT THE TEXT FOR CLARITY, GRAMMAR, AND STYLE. Make sure
that the response is clear, uses correct grammar, avoids abbreviations and
repetition, and is consistent with the tone of a clinical report written by a
medical professional. Report dates in the format "DD Month YYYY" (e.g., 15 June 2022).
"""


class Llm:
    """Class to represent the interface to a large language model.

    Each run_* method will return a placeholder that can be inserted into the
    report. All placeholders as well as their replacements will be stored in the
    placeholders attribute. The replacements are awaitables to allow for
    asynchronous processing.

    Attributes:
        client: The client to use for the language model.
        child_name: The name of the child in the report.
        child_pronouns: The pronouns of the child in the report.
        placeholders: The placeholders and their replacements.
    """

    def __init__(self, child_name: str, child_pronouns: Sequence[str]) -> None:
        """Initializes the language model.

        Args:
            model: The model to use for the language model.
            child_name: The name of the child in the report.
            child_pronouns: The pronouns of the child in the report.
        """
        self.client = azure.AzureLlm()
        self.child_name = child_name
        self.child_pronouns = child_pronouns
        self.placeholders: list[LlmPlaceholder] = []

    def run_text_with_parent_input(self, text: str, parent_input: str) -> str:
        """Creates a placeholder for an LLM edit of an excerpt with parent input.

        Args:
            text: The excerpt to edit.
            parent_input: The parent input to insert into the excerpt.

        Returns:
            The placeholder for the LLM edit.
        """
        user_prompt = f"""
            Excerpt: {text}.
            Parent Input: {parent_input}.
            Child name: {self.child_name}.
            Child pronouns: {string_utils.join_with_oxford_comma(self.child_pronouns)}.
        """
        user_prompt = string_utils.remove_excess_whitespace(user_prompt)
        replacement = self.client.run(Prompts.parent_input, user_prompt)
        id = str(uuid.uuid4())
        self.placeholders.append(LlmPlaceholder(id, replacement))
        return id

    def run_edit(self, text: str) -> str:
        """Creates a placeholder for an LLM edit of an excerpt.

        Args:
            text: The excerpt to edit.

        Returns:
            The placeholder for the LLM edit.
        """
        user_prompt = string_utils.remove_excess_whitespace(text)
        replacement = self.client.run(Prompts.edit, user_prompt)
        id = str(uuid.uuid4())
        self.placeholders.append(LlmPlaceholder(id, replacement))
        return id
