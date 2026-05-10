"""Text model client wrapper.

The Streamlit page owns UI state and error display. This module owns the
provider-specific Anthropic call shape so new writing features can reuse it
without duplicating SDK details.
"""

from typing import Iterable

import anthropic
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential


class MissingTextModelCredentials(RuntimeError):
    """Raised when the text model API key is not configured."""


def build_anthropic_client(api_key: str, base_url: str) -> anthropic.Anthropic:
    if not api_key:
        raise MissingTextModelCredentials("ANTHROPIC_AUTH_TOKEN is not configured")
    normalized_base_url = base_url.rstrip("/") if base_url else ""
    return anthropic.Anthropic(
        api_key=api_key,
        base_url=normalized_base_url if normalized_base_url != "https://api.anthropic.com" else None,
    )


def extract_text_blocks(content_blocks: Iterable[object]) -> str:
    return "".join(
        block.text
        for block in content_blocks
        if getattr(block, "type", "") == "text" and hasattr(block, "text")
    )


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    retry=retry_if_exception_type((anthropic.APITimeoutError, anthropic.InternalServerError)),
)
def call_text_model(
    *,
    prompt: str,
    api_key: str,
    base_url: str,
    model: str,
    timeout: int = 300,
    max_tokens: int = 8192,
    temperature: float = 0.2,
) -> str:
    client = build_anthropic_client(api_key, base_url)
    message = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        temperature=temperature,
        timeout=timeout,
        messages=[{"role": "user", "content": prompt}],
    )
    return extract_text_blocks(message.content)
