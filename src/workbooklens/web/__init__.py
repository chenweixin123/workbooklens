"""Local-only FastAPI interface."""

from workbooklens.web.app import create_app
from workbooklens.web.launcher import run_local_ui

__all__ = ["create_app", "run_local_ui"]
