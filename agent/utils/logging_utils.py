"""
Structured JSON run logger. Each entry is a newline-delimited JSON record.
"""
from __future__ import annotations

import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional


class RunLogger:
    def __init__(self, log_path: Optional[Path] = None, verbose: bool = True):
        self.log_path = log_path
        self.verbose = verbose
        self._file = open(log_path, "w") if log_path else None
        self.start_time = time.time()

    def _write(self, record: dict[str, Any]):
        record["ts"] = datetime.now(timezone.utc).isoformat()
        record["elapsed_s"] = round(time.time() - self.start_time, 2)
        line = json.dumps(record)
        if self._file:
            self._file.write(line + "\n")
            self._file.flush()
        if self.verbose:
            level = record.get("level", "INFO")
            msg = record.get("msg", "")
            detail = record.get("detail", "")
            tag = f"[{record.get('stage', '')}] " if record.get("stage") else ""
            suffix = f"  — {detail}" if detail else ""
            print(f"  {level:7s} {tag}{msg}{suffix}", file=sys.stderr)

    def info(self, msg: str, stage: str = "", **kwargs):
        self._write({"level": "INFO", "stage": stage, "msg": msg, **kwargs})

    def warn(self, msg: str, stage: str = "", **kwargs):
        self._write({"level": "WARN", "stage": stage, "msg": msg, **kwargs})

    warning = warn  # alias for compatibility

    def error(self, msg: str, stage: str = "", **kwargs):
        self._write({"level": "ERROR", "stage": stage, "msg": msg, **kwargs})

    def api_call(self, model: str, prompt_tokens: int, completion_tokens: int,
                 latency_ms: int, stage: str = ""):
        self._write({
            "level": "API",
            "stage": stage,
            "msg": "Claude API call",
            "model": model,
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "latency_ms": latency_ms,
        })

    def close(self):
        if self._file:
            self._file.close()
