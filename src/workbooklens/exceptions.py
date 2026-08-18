"""User-facing exception hierarchy and stable CLI exit codes."""

from __future__ import annotations


class ExitCode:
    """Documented process exit codes used by every CLI command."""

    OK = 0
    FINDINGS_OR_ASSERTIONS = 1
    USAGE = 2
    UNSAFE_INPUT = 3
    STALE_PLAN = 4
    VALIDATION_FAILED = 5
    INTERNAL_ERROR = 10


class WorkbookLensError(Exception):
    """Base class for expected failures safe to show to an end user."""

    exit_code = ExitCode.INTERNAL_ERROR


class UsageError(WorkbookLensError):
    """The requested operation or arguments are unsupported."""

    exit_code = ExitCode.USAGE


class UnsafeWorkbookError(WorkbookLensError):
    """An Office package violated a resource or parsing safety boundary."""

    exit_code = ExitCode.UNSAFE_INPUT


class StalePlanError(WorkbookLensError):
    """A patch plan no longer matches its source workbook or target cells."""

    exit_code = ExitCode.STALE_PLAN


class PatchValidationError(WorkbookLensError):
    """A proposed or completed patch failed closed during validation."""

    exit_code = ExitCode.VALIDATION_FAILED
