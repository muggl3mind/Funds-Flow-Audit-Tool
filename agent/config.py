from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from dotenv import load_dotenv

load_dotenv()


@dataclass
class DealConfig:
    deal_name: str
    closing_date: str                                    # "YYYY-MM-DD"
    client_role: Literal["buyer", "seller", "both"]
    funds_flow_path: Path
    documents_dir: Path
    output_dir: Path
    fund_allocations: dict[str, float] = field(default_factory=dict)  # {"Fund I": 0.60}
    match_confidence_threshold: float = 0.75
    anthropic_model: str = "claude-opus-4-6"
    anthropic_api_key: str = field(default_factory=lambda: os.environ.get("ANTHROPIC_API_KEY", ""))
    agent_version: str = "1.0.0"

    def __post_init__(self):
        self.funds_flow_path = Path(self.funds_flow_path)
        self.documents_dir = Path(self.documents_dir)
        self.output_dir = Path(self.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if not self.anthropic_api_key:
            raise ValueError("ANTHROPIC_API_KEY not set. Export it or add to .env file.")
