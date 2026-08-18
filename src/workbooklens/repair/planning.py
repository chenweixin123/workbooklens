"""Creation and loading of source-bound JSON patch manifests."""

from __future__ import annotations

from pathlib import Path

from pydantic import ValidationError

from workbooklens import __version__
from workbooklens.exceptions import UsageError
from workbooklens.models import PatchPlan
from workbooklens.scanner import ScanResult
from workbooklens.utils import write_json


def build_patch_plan(scan: ScanResult) -> PatchPlan:
    """Create a deterministic plan containing every reviewable proposed patch."""

    return PatchPlan(
        tool_version=__version__,
        source_name=scan.inspection.path.name,
        source_sha256=scan.snapshot.source_sha256,
        patches=scan.patches,
        finding_ids=[finding.id for finding in scan.findings if finding.patch_ids],
        findings=[finding for finding in scan.findings if finding.patch_ids],
    )


def write_patch_plan(path: Path, plan: PatchPlan) -> None:
    """Write a plan using the public JSON schema representation."""

    write_json(path, plan.model_dump(mode="json"))


def load_patch_plan(path: Path, max_bytes: int = 10 * 1024 * 1024) -> PatchPlan:
    """Read and strictly validate a patch plan supplied by a user."""

    try:
        encoded = path.read_bytes()
    except OSError as exc:
        raise UsageError(f"Unable to read patch plan {path}: {exc}") from exc
    if len(encoded) > max_bytes:
        raise UsageError(f"Patch plan exceeds the {max_bytes}-byte limit")
    try:
        return PatchPlan.model_validate_json(encoded)
    except (ValidationError, UnicodeDecodeError) as exc:
        raise UsageError(f"Patch plan does not match schema version 1: {exc}") from exc
