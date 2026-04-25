"""Config management via .env."""

from __future__ import annotations

from dotenv import load_dotenv
import os

load_dotenv()


class Config:
    NEO4J_URI = os.getenv("NEO4J_URI", "bolt://localhost:7687")
    NEO4J_USER = os.getenv("NEO4J_USER", "neo4j")
    NEO4J_PASSWORD = os.getenv("NEO4J_PASSWORD", "")

    DASHSCOPE_API_KEY = os.getenv("DASHSCOPE_API_KEY", "")
    DASHSCOPE_MODEL = os.getenv("DASHSCOPE_MODEL", "qwen-plus")

    @classmethod
    def has_neo4j(cls) -> bool:
        return bool(cls.NEO4J_PASSWORD and cls.NEO4J_PASSWORD != "your_neo4j_password")

    @classmethod
    def has_llm(cls) -> bool:
        return bool(cls.DASHSCOPE_API_KEY and cls.DASHSCOPE_API_KEY != "your_dashscope_api_key")
