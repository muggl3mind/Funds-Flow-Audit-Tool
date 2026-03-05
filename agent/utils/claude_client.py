"""
Thin wrapper around the Anthropic API.
Handles JSON extraction, retry logic, and logging.
"""
from __future__ import annotations

import json
import re
import time
from typing import Any, Optional

import anthropic

from agent.utils.logging_utils import RunLogger


class ClaudeClient:
    def __init__(self, api_key: str, model: str, logger: Optional[RunLogger] = None):
        self.model = model
        self.logger = logger
        self._client = anthropic.Anthropic(api_key=api_key)

    def call(
        self,
        prompt: str,
        system: str = "You are an expert PE/M&A finance analyst.",
        max_tokens: int = 2048,
        temperature: float = 0.0,
        stage: str = "",
    ) -> str:
        t0 = time.time()
        response = self._client.messages.create(
            model=self.model,
            max_tokens=max_tokens,
            temperature=temperature,
            system=system,
            messages=[{"role": "user", "content": prompt}],
        )
        latency_ms = int((time.time() - t0) * 1000)
        text = response.content[0].text
        if self.logger:
            self.logger.api_call(
                model=self.model,
                prompt_tokens=response.usage.input_tokens,
                completion_tokens=response.usage.output_tokens,
                latency_ms=latency_ms,
                stage=stage,
            )
        return text

    def call_json(
        self,
        prompt: str,
        system: str = "You are an expert PE/M&A finance analyst. Always respond with valid JSON only.",
        max_tokens: int = 2048,
        stage: str = "",
        retries: int = 3,
    ) -> Any:
        for attempt in range(retries):
            raw = self.call(prompt, system=system, max_tokens=max_tokens, stage=stage)
            parsed = _extract_json(raw)
            if parsed is not None:
                return parsed
            if self.logger:
                self.logger.warn(
                    f"JSON parse failed (attempt {attempt + 1}/{retries}), retrying",
                    stage=stage,
                    detail=raw[:200],
                )
            time.sleep(1.5 ** attempt)
        raise ValueError(f"Could not extract valid JSON from Claude response after {retries} attempts")


def _extract_json(text: str) -> Optional[Any]:
    """Try to extract JSON from a response that may have prose around it."""
    # Try direct parse first
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError:
        pass

    # Try to find a JSON block (``` or raw {...} / [...])
    patterns = [
        r"```json\s*([\s\S]+?)\s*```",
        r"```\s*([\s\S]+?)\s*```",
        r"(\{[\s\S]+\})",
        r"(\[[\s\S]+\])",
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                continue
    return None
