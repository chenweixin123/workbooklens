"""Small deterministic IO and hashing helpers."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from pathlib import Path
from typing import Any


def sha256_bytes(value: bytes) -> str:
    """Return a lowercase SHA-256 digest for bytes."""

    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    """Stream a file into SHA-256 without loading the full workbook in memory."""

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest()


def stable_json_bytes(value: Any) -> bytes:
    """Serialize JSON in a reproducible representation suitable for fingerprints."""

    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    ).encode("utf-8")


def write_json(path: Path, value: Any) -> None:
    """Atomically write human-readable UTF-8 JSON."""

    data = json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True, default=str) + "\n"
    atomic_write_bytes(path, data.encode("utf-8"))


def atomic_write_bytes(path: Path, data: bytes) -> None:
    """Write bytes next to the destination and atomically replace it."""

    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    temporary_path = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        temporary_path.replace(path)
    except BaseException:
        temporary_path.unlink(missing_ok=True)
        raise


def stable_id(prefix: str, *parts: Any, length: int = 16) -> str:
    """Build a stable readable identifier from canonicalized semantic parts."""

    digest = sha256_bytes(stable_json_bytes(parts))[:length]
    return f"{prefix}-{digest}"
