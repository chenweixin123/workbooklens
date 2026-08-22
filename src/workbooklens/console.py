"""Console encoding helpers for frozen Windows entry points."""

from __future__ import annotations

import os
import sys


def configure_utf8_redirected_streams(
    streams: tuple[object | None, ...] | None = None,
) -> None:
    """Emit deterministic UTF-8 when Windows output is redirected or captured."""

    if os.name != "nt":
        return
    selected = (sys.stdout, sys.stderr) if streams is None else streams
    for stream in selected:
        if stream is None:
            continue
        isatty = getattr(stream, "isatty", None)
        reconfigure = getattr(stream, "reconfigure", None)
        if not callable(isatty) or not callable(reconfigure) or isatty():
            continue
        reconfigure(encoding="utf-8", errors="replace")
