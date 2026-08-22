from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from scripts.build_portable_windows import (
    PortableBuildError,
    collect_distribution_licenses,
    create_deterministic_zip,
    read_wheel_metadata,
    render_version_info,
    sanitized_python_environment,
)


def _wheel(path: Path, *, name: str = "WorkbookLens", version: str = "2.2.1") -> Path:
    metadata = f"Metadata-Version: 2.4\nName: {name}\nVersion: {version}\n\n"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(f"workbooklens-{version}.dist-info/METADATA", metadata.encode())
    return path


def test_reads_workbooklens_wheel_metadata(tmp_path: Path) -> None:
    metadata = read_wheel_metadata(_wheel(tmp_path / "workbooklens.whl"))

    assert metadata.name == "WorkbookLens"
    assert metadata.version == "2.2.1"


def test_rejects_unrelated_wheel(tmp_path: Path) -> None:
    wheel = _wheel(tmp_path / "other.whl", name="other")

    with pytest.raises(PortableBuildError, match="WorkbookLens"):
        read_wheel_metadata(wheel)


def test_renders_windows_version_resource() -> None:
    rendered = render_version_info("2.2.1")

    assert "filevers=(2, 2, 1, 0)" in rendered
    assert "StringStruct('ProductVersion', '2.2.1')" in rendered


def test_collects_dist_info_license_files(tmp_path: Path) -> None:
    site_packages = tmp_path / "site-packages"
    dist_info = site_packages / "sample-1.0.dist-info"
    (dist_info / "licenses").mkdir(parents=True)
    (dist_info / "METADATA").write_text(
        "Metadata-Version: 2.4\nName: sample\nVersion: 1.0\nLicense-Expression: MIT\n\n",
        encoding="utf-8",
    )
    (dist_info / "licenses" / "LICENSE.txt").write_text("MIT license\n", encoding="utf-8")
    output = tmp_path / "LICENSES"
    output.mkdir()

    notices = collect_distribution_licenses(site_packages, output)

    assert notices[0].license_expression == "MIT"
    assert (output / notices[0].copied_files[0]).read_text() == "MIT license\n"


def test_collects_declared_british_license_and_authors_files(tmp_path: Path) -> None:
    site_packages = tmp_path / "site-packages"
    dist_info = site_packages / "sample-1.0.dist-info"
    dist_info.mkdir(parents=True)
    (dist_info / "METADATA").write_text(
        (
            "Metadata-Version: 2.4\n"
            "Name: sample\n"
            "Version: 1.0\n"
            "License-File: LICENCE.rst\n"
            "License-File: AUTHORS.txt\n\n"
        ),
        encoding="utf-8",
    )
    (dist_info / "LICENCE.rst").write_text("British license\n", encoding="utf-8")
    (dist_info / "AUTHORS.txt").write_text("Contributors\n", encoding="utf-8")
    output = tmp_path / "LICENSES"
    output.mkdir()

    notices = collect_distribution_licenses(site_packages, output)

    copied = {(output / path).read_text(encoding="utf-8") for path in notices[0].copied_files}
    assert copied == {"British license\n", "Contributors\n"}


def test_rejects_missing_declared_license_file(tmp_path: Path) -> None:
    site_packages = tmp_path / "site-packages"
    dist_info = site_packages / "sample-1.0.dist-info"
    dist_info.mkdir(parents=True)
    (dist_info / "METADATA").write_text(
        ("Metadata-Version: 2.4\nName: sample\nVersion: 1.0\nLicense-File: LICENCE.rst\n\n"),
        encoding="utf-8",
    )
    output = tmp_path / "LICENSES"
    output.mkdir()

    with pytest.raises(PortableBuildError, match="missing declared License-File"):
        collect_distribution_licenses(site_packages, output)


def test_deterministic_zip_is_reproducible(tmp_path: Path) -> None:
    source = tmp_path / "WorkbookLens-2.2.1-windows-x64"
    source.mkdir()
    (source / "b.txt").write_text("b", encoding="utf-8")
    (source / "a.txt").write_text("a", encoding="utf-8")
    first = tmp_path / "first.zip"
    second = tmp_path / "second.zip"

    create_deterministic_zip(source, first)
    create_deterministic_zip(source, second)

    assert first.read_bytes() == second.read_bytes()


def test_sanitized_python_environment_removes_interpreter_injection() -> None:
    source = {
        "PATH": r"C:\Windows\System32",
        "PYTHONPATH": r"C:\repo\src",
        "pythonhome": r"C:\Python312",
        "VIRTUAL_ENV": r"C:\repo\.venv",
        "CONDA_PREFIX": r"C:\Miniconda",
        "__PYVENV_LAUNCHER__": r"C:\shim.exe",
        "PIP_INDEX_URL": "https://example.invalid/simple",
    }

    result = sanitized_python_environment(source)

    assert result["PATH"] == source["PATH"]
    assert result["PIP_INDEX_URL"] == source["PIP_INDEX_URL"]
    assert result["PYTHONNOUSERSITE"] == "1"
    assert (
        not {
            "PYTHONPATH",
            "pythonhome",
            "VIRTUAL_ENV",
            "CONDA_PREFIX",
            "__PYVENV_LAUNCHER__",
        }
        & result.keys()
    )


def test_windows_launcher_uses_a_dedicated_console() -> None:
    launcher = (
        Path(__file__).parents[1] / "packaging" / "windows" / "Start-WorkbookLens.cmd"
    ).read_text(encoding="utf-8")

    assert 'start "WorkbookLens" /D "%~dp0" "%~dp0WorkbookLens.exe"' in launcher
    assert "Press Ctrl+C there to stop WorkbookLens." in launcher
    assert "Terminate batch job" not in launcher


def test_windows_frozen_runtime_enables_utf8_before_startup() -> None:
    spec = (Path(__file__).parents[1] / "packaging" / "windows" / "WorkbookLens.spec").read_text(
        encoding="utf-8"
    )

    assert '("X utf8", None, "OPTION")' in spec
    assert '"winreg"' in spec
