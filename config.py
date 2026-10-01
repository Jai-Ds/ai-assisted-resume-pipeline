# -*- coding: utf-8 -*-
"""
config.py — shared config loader for the job-application pipeline.

Looks for config.json next to this file. First run: copy config.example.json
to config.json and fill in your own profile, paths, search categories, and
base resume tags. config.json is gitignored — it holds your personal data.
"""

import json
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_CONFIG_PATH = os.path.join(_HERE, "config.json")
_EXAMPLE_PATH = os.path.join(_HERE, "config.example.json")


def _load_env_file():
    """Load KEY=VALUE pairs from a local .env file (gitignored) into os.environ."""
    env_path = os.path.join(_HERE, ".env")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def load_config():
    if not os.path.exists(_CONFIG_PATH):
        raise RuntimeError(
            f"config.json not found at {_CONFIG_PATH}.\n"
            f"Copy config.example.json to config.json and fill in your own "
            f"profile, paths, search categories, and base resume tags."
        )
    with open(_CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def resolve_path(path_str):
    """Resolve a config path (may be relative to this file) to an absolute path."""
    if os.path.isabs(path_str):
        return path_str
    return os.path.normpath(os.path.join(_HERE, path_str))


_load_env_file()
CONFIG = load_config()

SERPAPI_KEY = os.environ.get("SERPAPI_KEY", "")

PROFILE          = CONFIG["profile"]
PATHS            = {k: resolve_path(v) for k, v in CONFIG["paths"].items()}
BASE_RESUME_TAGS = CONFIG["base_resume_tags"]
CATEGORIES       = CONFIG["categories"]
TITLE_BLACKLIST  = CONFIG["title_blacklist"]
