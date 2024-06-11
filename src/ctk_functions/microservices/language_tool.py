"""Provides a class for correcting text using the LanguageTool API."""

from typing import Collection

import aiohttp
import pydantic

from ctk_functions import config

settings = config.get_settings()

LANGUAGE_TOOL_ENDPOINT = str(settings.LANGUAGE_TOOL_ENDPOINT)


@pydantic.dataclasses.dataclass
class Url:
    """Represents a URL from the LanguageTool API."""

    value: str


@pydantic.dataclasses.dataclass
class Category:
    """Represents a category from the LanguageTool API."""

    id: str
    name: str


@pydantic.dataclasses.dataclass
class Replacement:
    """Represents a replacement from the LanguageTool API."""

    value: str
    shortDescription: str | None = None


@pydantic.dataclasses.dataclass
class Context:
    """Represents a context from the LanguageTool API."""

    text: str
    offset: int
    length: int


@pydantic.dataclasses.dataclass
class Rule:
    """Represents a rule from the LanguageTool API."""

    id: str
    description: str
    issueType: str
    category: Category
    urls: list[Url] | None = None
    subId: str | None = None
    sourceFile: str | None = None


@pydantic.dataclasses.dataclass
class Correction:
    """Represents a correction from the LanguageTool API."""

    message: str
    shortMessage: str
    replacements: list[Replacement]
    rule: Rule
    offset: int
    length: int
    context: Context
    sentence: str
    ignoreForIncompleteSentence: bool
    contextForSureMatch: int


class LanguageCorrecter:
    """Corrects text using the LanguageTool API."""

    def __init__(self, url: str = LANGUAGE_TOOL_ENDPOINT) -> None:
        """Initializes the correcter with the LanguageTool API URL."""
        self.url = url

    async def check(
        self,
        text: str,
        localization: str = "en-US",
        enabled_rules: Collection[str] | None = None,
    ) -> list[Correction]:
        """Corrects the text using the LanguageTool API.

        Args:
            text: The text to check.
            localization: The localization of the text. Defaults to "en-US".
            enabled_rules: The rules to enable for the correction.

        Returns:
            The suggested corrections.
        """
        data = {
            "text": text,
            "language": localization,
        }
        if enabled_rules:
            data["enabledRules"] = ",".join(enabled_rules)
            data["enabledOnly"] = "true"

        async with aiohttp.ClientSession() as session:
            async with session.post(self.url, data=data) as response:
                response.raise_for_status()
                results = await response.json()
        return [Correction(**result) for result in results["matches"]]
