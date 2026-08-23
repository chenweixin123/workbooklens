"""Sanitized, localized public errors with stable support codes."""

from __future__ import annotations

import secrets
from dataclasses import dataclass
from typing import Literal

from workbooklens.exceptions import (
    PatchValidationError,
    StalePlanError,
    UnsafeWorkbookError,
    UsageError,
    WorkbookLensError,
)
from workbooklens.i18n.catalog import Language, require_translation

PUBLIC_ERROR_CODES: dict[str, str] = {
    "security.cross_origin": "WL-SEC-001",
    "security.csrf": "WL-SEC-002",
    "upload.too_large": "WL-UPL-001",
    "upload.empty": "WL-UPL-002",
    "upload.invalid_type": "WL-UPL-003",
    "session.limit": "WL-SES-001",
    "session.not_found": "WL-SES-002",
    "conversion.invalid_input": "WL-CNV-001",
    "conversion.unavailable": "WL-CNV-002",
    "conversion.all_providers_failed": "WL-CNV-003",
    "conversion.output_invalid": "WL-CNV-004",
    "conversion.timeout": "WL-CNV-005",
    "scan.failed": "WL-SCN-001",
    "repair.selection_required": "WL-RPR-001",
    "repair.stale_plan": "WL-RPR-002",
    "repair.validation_failed": "WL-RPR-003",
    "download.not_ready": "WL-DWN-001",
    "request.invalid": "WL-REQ-001",
    "request.baseline_required": "WL-REQ-002",
    "request.baseline_scope_mismatch": "WL-REQ-003",
    "internal.unexpected": "WL-INT-001",
}


@dataclass(frozen=True, slots=True)
class UserFacingError:
    """A path-free error safe for HTML, CLI, and screenshots."""

    key: str
    code: str
    title: str
    message: str
    suggestion: str
    diagnostic_id: str | None = None


def new_diagnostic_id() -> str:
    """Create a non-sensitive correlation ID for a matching local log entry."""

    return "WL-" + secrets.token_hex(6).upper()


def localized_error(
    key: str,
    language: Language | str | None = None,
    *,
    diagnostic_id: str | None = None,
) -> UserFacingError:
    """Build a sanitized public error from its stable classification key."""

    if key not in PUBLIC_ERROR_CODES:
        key = "internal.unexpected"
    return UserFacingError(
        key=key,
        code=PUBLIC_ERROR_CODES[key],
        title=require_translation(f"error.{key}.title", language),
        message=require_translation(f"error.{key}.message", language),
        suggestion=require_translation(f"error.{key}.suggestion", language),
        diagnostic_id=diagnostic_id,
    )


def localize_exception(
    exc: BaseException,
    language: Language | str | None = None,
    *,
    operation: Literal["conversion", "scan", "repair", "download", "request"] = "request",
    diagnostic_id: str | None = None,
) -> UserFacingError:
    """Classify without exposing raw paths, commands, traces, or CLIXML payloads."""

    explicit = getattr(exc, "error_key", None)
    if isinstance(explicit, str) and explicit in PUBLIC_ERROR_CODES:
        key = explicit
    elif isinstance(exc, StalePlanError):
        key = "repair.stale_plan"
    elif isinstance(exc, PatchValidationError):
        key = "repair.validation_failed"
    elif isinstance(exc, UnsafeWorkbookError):
        key = "scan.failed"
    elif isinstance(exc, (UsageError, WorkbookLensError)):
        key = "request.invalid"
    elif operation == "conversion":
        key = "conversion.all_providers_failed"
    elif operation == "scan":
        key = "scan.failed"
    elif operation == "repair":
        key = "repair.validation_failed"
    elif operation == "download":
        key = "download.not_ready"
    else:
        key = "internal.unexpected"
    return localized_error(key, language, diagnostic_id=diagnostic_id)


def assert_error_catalog_complete() -> None:
    """Fail when a stable public error lacks either language or field."""

    for key in PUBLIC_ERROR_CODES:
        for suffix in ("title", "message", "suggestion"):
            require_translation(f"error.{key}.{suffix}", "en")
            require_translation(f"error.{key}.{suffix}", "zh-CN")


__all__ = [
    "PUBLIC_ERROR_CODES",
    "UserFacingError",
    "assert_error_catalog_complete",
    "localize_exception",
    "localized_error",
    "new_diagnostic_id",
]
