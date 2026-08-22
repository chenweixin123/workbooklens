from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("winreg")

from scripts import smoke_installer_windows as installer_smoke


def test_inno_cleanup_requires_fixed_path_marker_and_reparse_checks() -> None:
    repository_root = Path(__file__).resolve().parents[1]
    source = (repository_root / "packaging" / "windows" / "WorkbookLens.iss").read_text(
        encoding="utf-8"
    )
    install_delete = source.split("[InstallDelete]", 1)[1].split("[UninstallDelete]", 1)[0]
    delete_entries = [
        line.strip() for line in install_delete.splitlines() if line.strip().startswith("Type:")
    ]

    assert delete_entries
    assert all(entry.endswith("; Check: ShouldCleanPreviousPayload") for entry in delete_entries)
    assert "DisableDirPage=yes" in source
    assert "UsePreviousAppDir=no" in source
    assert "function PrepareToInstall(var NeedsRestart: Boolean): String;" in source
    assert "PreviousInstallDir" not in source
    assert "function HasValidOwnershipMarker(Directory: String): Boolean;" in source
    assert "function TreeHasReparsePoint(Path: String): Boolean;" in source
    assert "FileAttributeReparsePoint = $400;" in source
    assert "SamePath(ExpandConstant('{app}'), DefaultInstallDir())" in source
    assert "HasValidOwnershipMarker(ExpandConstant('{app}'))" in source
    assert "not TreeHasReparsePoint(ExpandConstant('{app}'))" in source
    assert "SaveStringToFile(" in source
    assert "[UninstallDelete]" in source
    assert "DirectoryHasEntries(TargetDir)" in source


def test_installed_payload_hashes_exclude_installer_metadata(tmp_path: Path) -> None:
    (tmp_path / "WorkbookLens.exe").write_bytes(b"app")
    (tmp_path / "unins000.exe").write_bytes(b"uninstaller")
    (tmp_path / "unins000.dat").write_bytes(b"uninstaller data")
    (tmp_path / installer_smoke.OWNERSHIP_MARKER_NAME).write_text(
        installer_smoke.OWNERSHIP_MARKER_CONTENT,
        encoding="ascii",
    )

    hashes = installer_smoke._installed_payload_hashes(tmp_path)

    assert set(hashes) == {"WorkbookLens.exe"}


def test_test_uninstall_entry_uses_exact_owned_key(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    recorded: list[tuple[str, object]] = []

    class FakeKey:
        def __enter__(self) -> FakeKey:
            return self

        def __exit__(self, *args: object) -> None:
            return None

    monkeypatch.setattr(installer_smoke.winreg, "CreateKey", lambda *args: FakeKey())
    monkeypatch.setattr(
        installer_smoke.winreg,
        "SetValueEx",
        lambda key, name, reserved, kind, value: recorded.append((name, value)),
    )

    installer_smoke._write_test_uninstall_entry(tmp_path, version="0.0.0-test")

    assert recorded == [
        ("DisplayName", "WorkbookLens"),
        ("DisplayVersion", "0.0.0-test"),
        ("InstallLocation", str(tmp_path)),
    ]


def test_payload_comparison_reports_stale_upgrade_member(tmp_path: Path) -> None:
    internal = tmp_path / "_internal"
    internal.mkdir()
    (internal / "expected.dll").write_bytes(b"expected")
    expected = installer_smoke._installed_payload_hashes(tmp_path)
    (internal / "stale.dll").write_bytes(b"stale")

    with pytest.raises(installer_smoke.InstallerSmokeError, match=r"stale\.dll"):
        installer_smoke._assert_payload_matches(
            tmp_path,
            expected,
            phase="upgrade",
        )


def test_fallback_cleanup_removes_only_test_owned_install_tree(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    scratch = tmp_path / "scratch"
    install_dir = scratch / "install"
    group = tmp_path / "start-menu" / "WorkbookLens"
    desktop_link = tmp_path / "desktop" / "WorkbookLens.lnk"
    install_dir.mkdir(parents=True)
    (install_dir / "partial.bin").write_bytes(b"partial")
    monkeypatch.setattr(installer_smoke, "_owned_uninstall_entry", lambda: None)

    errors = installer_smoke._cleanup_owned_state(
        install_dir=install_dir,
        scratch=scratch,
        group=group,
        desktop_link=desktop_link,
    )

    assert errors == []
    assert not install_dir.exists()


def test_cleanup_never_recursively_removes_nonempty_default_install_dir(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    scratch = tmp_path / "scratch"
    install_dir = tmp_path / "default-install"
    group = tmp_path / "start-menu" / "WorkbookLens"
    desktop_link = tmp_path / "desktop" / "WorkbookLens.lnk"
    foreign_file = install_dir / "foreign-user-data.txt"
    install_dir.mkdir()
    foreign_file.write_bytes(b"must survive")
    monkeypatch.setattr(installer_smoke, "_owned_uninstall_entry", lambda: None)
    monkeypatch.setattr(installer_smoke, "_default_install_dir", lambda: install_dir)

    errors = installer_smoke._cleanup_owned_state(
        install_dir=install_dir,
        scratch=scratch,
        group=group,
        desktop_link=desktop_link,
    )

    assert any("refusing to recursively remove" in error for error in errors)
    assert foreign_file.read_bytes() == b"must survive"


def test_setup_command_uses_default_unless_registered_path_is_explicit(
    tmp_path: Path,
) -> None:
    installer = tmp_path / "setup.exe"
    log = tmp_path / "setup.log"

    fresh = installer_smoke._setup_command(installer, log, install_dir=None)
    upgrade = installer_smoke._setup_command(
        installer,
        log,
        install_dir=tmp_path / "registered-install",
    )

    assert not any(argument.startswith("/DIR=") for argument in fresh)
    assert any(argument.startswith("/DIR=") for argument in upgrade)


def test_cleanup_errors_are_added_to_primary_exception() -> None:
    primary = ValueError("original installer failure")

    installer_smoke._record_cleanup_errors(primary, ["cleanup failed"])

    assert str(primary) == "original installer failure"
    assert primary.__notes__ == ["installer smoke cleanup issues: cleanup failed"]


def test_cleanup_errors_raise_when_no_primary_exception() -> None:
    with pytest.raises(installer_smoke.InstallerSmokeError, match="cleanup failed"):
        installer_smoke._record_cleanup_errors(None, ["cleanup failed"])
