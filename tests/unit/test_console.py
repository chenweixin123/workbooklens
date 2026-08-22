from __future__ import annotations

import workbooklens.console as console_support


class _Stream:
    def __init__(self, *, tty: bool) -> None:
        self.tty = tty
        self.calls: list[tuple[str, str]] = []

    def isatty(self) -> bool:
        return self.tty

    def reconfigure(self, *, encoding: str, errors: str) -> None:
        self.calls.append((encoding, errors))


def test_windows_redirected_streams_use_utf8(monkeypatch) -> None:
    redirected = _Stream(tty=False)
    terminal = _Stream(tty=True)
    monkeypatch.setattr(console_support.os, "name", "nt")

    console_support.configure_utf8_redirected_streams((redirected, terminal, None))

    assert redirected.calls == [("utf-8", "replace")]
    assert terminal.calls == []


def test_non_windows_streams_are_unchanged(monkeypatch) -> None:
    redirected = _Stream(tty=False)
    monkeypatch.setattr(console_support.os, "name", "posix")

    console_support.configure_utf8_redirected_streams((redirected,))

    assert redirected.calls == []
