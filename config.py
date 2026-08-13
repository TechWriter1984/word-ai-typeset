"""配置加载：读取 config.yaml + key.txt，返回统一配置对象。"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any, Dict

import yaml


@dataclass
class Config:
    raw: Dict[str, Any] = field(default_factory=dict)
    base_dir: str = ""

    # ---- 便捷访问器 ----
    @property
    def template_docx(self) -> str:
        return self.raw["paths"]["template_docx"]

    @property
    def input_docx(self) -> str:
        return self.raw["paths"]["input_docx"]

    @property
    def output_dir(self) -> str:
        return self.raw["paths"]["output_dir"]

    @property
    def key_file(self) -> str:
        return self.raw["paths"]["key_file"]

    @property
    def api_key(self) -> str:
        if not os.path.exists(self.key_file):
            raise FileNotFoundError(f"API Key 文件不存在: {self.key_file}")
        with open(self.key_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    return line
        raise RuntimeError(f"未在 {self.key_file} 中找到 API Key")

    @property
    def section_marker(self) -> str:
        return self.raw["template"]["section_marker"]

    @property
    def clear_placeholder(self) -> bool:
        return self.raw["template"]["clear_placeholder"]

    @property
    def llm(self) -> Dict[str, Any]:
        return self.raw["llm"]

    @property
    def style_map(self) -> Dict[str, str]:
        return self.raw["style_map"]

    @property
    def typeset(self) -> Dict[str, Any]:
        return self.raw["typeset"]


def load_config(config_path: str | None = None) -> Config:
    """加载 config.yaml。默认从本文件同目录读取。"""
    if config_path is None:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    return Config(raw=raw or {}, base_dir=os.path.dirname(os.path.abspath(config_path)))
