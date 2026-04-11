"""
Groq API client factory for the ImageToExcel pipeline.

Centralizes client creation and API key resolution so that
Streamlit UI, CLI, and tests all share the same logic.
"""

from __future__ import annotations

import logging
import os
from typing import Optional

from groq import Groq

from core.exceptions import APIKeyMissingError

logger = logging.getLogger(__name__)


def _resolve_api_key(api_key: Optional[str] = None) -> str:
    """Resolve the Groq API key from multiple sources.

    Priority order:
        1. Explicit `api_key` argument (e.g. from UI input)
        2. Streamlit secrets (for Streamlit Cloud deployment)
        3. Environment variable (for local / CLI usage)

    Args:
        api_key: An optional explicit API key string.

    Returns:
        The resolved API key.

    Raises:
        APIKeyMissingError: If no key is found in any source.
    """
    if api_key:
        return api_key.strip().strip("'\"")

    # 2. Try Environment variable (loaded from .env or shell)
    key = os.environ.get("GROQ_API_KEY", "")
    if key and "your_key" not in key.lower():
        return key.strip().strip("'\"")

    # 3. Try Streamlit secrets (only available inside a Streamlit runtime or if dummy file exists)
    try:
        import streamlit as st
        key = st.secrets.get("GROQ_API_KEY", "")
        if key and "your_key" not in key.lower():
            return key.strip().strip("'\"")
    except (ImportError, Exception):
        pass

    raise APIKeyMissingError()


def create_groq_client(api_key: Optional[str] = None) -> Groq:
    """Create and return a configured Groq client.

    Args:
        api_key: Optional explicit API key. If not provided,
                 the key is resolved from Streamlit secrets or env vars.

    Returns:
        A ready-to-use Groq client instance.

    Raises:
        APIKeyMissingError: If no API key is found.
    """
    resolved_key = _resolve_api_key(api_key)
    logger.debug("Groq client created successfully")
    return Groq(api_key=resolved_key)
