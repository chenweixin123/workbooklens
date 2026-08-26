from __future__ import annotations

from pathlib import Path

import pytest

from workbooklens.exceptions import UsageError
from workbooklens.rules.profile_quality import MAX_PROFILE_RANGE_CELLS
from workbooklens.testing import load_test_config


def test_yaml_profile_loads_into_stable_scan_mapping(tmp_path: Path) -> None:
    path = tmp_path / "workbooklens.yml"
    path.write_text(
        """\
version: 2
profile:
  infer_semantics: false
  report_trailing_whitespace: true
  review_trailing_whitespace_patches: false
  sheets:
    - sheet: Sales
      range: A1:H500
      header_row: 1
      columns:
        - header: Order ID
          role: identifier
          required: true
          identifier_width: 8
          preserve_leading_zeros: true
        - column: D
          role: category
          allowed_values: [North, South, East, West]
        - header: Email
          role: email
        - header: Amount
          role: currency
        - header: Discount
          role: percentage
""",
        encoding="utf-8",
    )

    config = load_test_config(path)
    payload = config.model_dump(mode="python")

    assert payload["profile"]["infer_semantics"] is False
    sheet = payload["profile"]["sheets"][0]
    assert sheet["sheet"] == "Sales"
    assert sheet["range"] == "A1:H500"
    assert sheet["columns"][0]["identifier_width"] == 8
    assert sheet["columns"][1]["column"] == "D"
    assert sheet["columns"][1]["allowed_values"] == ["North", "South", "East", "West"]


def test_yaml_profile_v3_accepts_column_repair_modes(tmp_path: Path) -> None:
    path = tmp_path / "workbooklens-v3.yml"
    path.write_text(
        "version: 3\nprofile:\n  sheets:\n    - sheet: Data\n      columns:\n"
        "        - header: Amount\n          role: currency\n          repair: auto\n"
        "        - header: Notes\n          role: text\n          repair: review\n",
        encoding="utf-8",
    )

    config = load_test_config(path)

    assert config.version == 3
    assert config.profile is not None
    assert [column.repair for column in config.profile.sheets[0].columns] == ["auto", "review"]


def test_yaml_profile_v2_rejects_column_repair_mode(tmp_path: Path) -> None:
    path = tmp_path / "workbooklens-v2-repair.yml"
    path.write_text(
        "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      columns:\n"
        "        - header: Amount\n          role: currency\n          repair: auto\n",
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="repair requires configuration version 3"):
        load_test_config(path)


@pytest.mark.parametrize(
    "column_block",
    [
        "        - role: email\n",
        "        - column: XFE\n          role: text\n",
        "        - column: 1000000000\n          role: text\n",
        "        - header: Name\n          unknown: true\n",
    ],
)
def test_yaml_profile_rejects_ambiguous_or_unsupported_columns(
    tmp_path: Path,
    column_block: str,
) -> None:
    path = tmp_path / "invalid.yml"
    path.write_text(
        "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      columns:\n" + column_block,
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="configuration is invalid"):
        load_test_config(path)


def test_yaml_profile_rejects_duplicate_column_selectors(tmp_path: Path) -> None:
    path = tmp_path / "duplicates.yml"
    path.write_text(
        """\
version: 2
profile:
  sheets:
    - sheet: Data
      columns:
        - header: Status
          role: category
        - header: status
          role: text
""",
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="duplicate column selectors"):
        load_test_config(path)


@pytest.mark.parametrize(
    "profile_body",
    [
        "      header_row: true\n",
        "      columns:\n        - column: true\n          role: text\n",
        (
            "      columns:\n"
            "        - header: ID\n"
            "          role: identifier\n"
            "          identifier_width: true\n"
        ),
    ],
)
def test_yaml_profile_rejects_boolean_integer_coercion(
    tmp_path: Path,
    profile_body: str,
) -> None:
    path = tmp_path / "boolean-integer.yml"
    path.write_text(
        "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n" + profile_body,
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="configuration is invalid"):
        load_test_config(path)


def test_yaml_version_rejects_boolean_coercion(tmp_path: Path) -> None:
    path = tmp_path / "boolean-version.yml"
    path.write_text("version: true\n", encoding="utf-8")

    with pytest.raises(UsageError, match="configuration is invalid"):
        load_test_config(path)


def test_yaml_profile_range_defaults_header_row_to_first_range_row(tmp_path: Path) -> None:
    path = tmp_path / "range-header.yml"
    path.write_text(
        "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      range: A5:B10\n",
        encoding="utf-8",
    )

    config = load_test_config(path)

    assert config.profile is not None
    assert config.profile.sheets[0].header_row == 5


def test_yaml_profile_rejects_header_row_outside_range(tmp_path: Path) -> None:
    path = tmp_path / "outside-header.yml"
    path.write_text(
        (
            "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n"
            "      range: A5:B10\n      header_row: 1\n"
        ),
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="header_row must be inside"):
        load_test_config(path)


def test_yaml_profile_rejects_dual_column_selector(tmp_path: Path) -> None:
    path = tmp_path / "dual-selector.yml"
    path.write_text(
        (
            "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      columns:\n"
            "        - header: ID\n          column: A\n          role: identifier\n"
        ),
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="exactly one header or column selector"):
        load_test_config(path)


def test_yaml_profile_rejects_duplicate_table_selectors(tmp_path: Path) -> None:
    path = tmp_path / "duplicate-tables.yml"
    path.write_text(
        (
            "version: 2\nprofile:\n  sheets:\n"
            "    - sheet: Data\n      range: A1:B10\n"
            "    - sheet: data\n      range: A20:B30\n"
        ),
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="duplicate sheet table selectors"):
        load_test_config(path)


def test_yaml_profile_requires_configuration_version_2(tmp_path: Path) -> None:
    path = tmp_path / "profile-v1.yml"
    path.write_text(
        "version: 1\nprofile:\n  sheets:\n    - sheet: Data\n",
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="profiles require configuration version 2"):
        load_test_config(path)


def test_yaml_profile_range_uses_the_shared_maximum_cell_limit(tmp_path: Path) -> None:
    accepted = tmp_path / "accepted.yml"
    accepted.write_text(
        "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      range: A1:J100000\n",
        encoding="utf-8",
    )
    assert load_test_config(accepted).profile is not None

    rejected = tmp_path / "rejected.yml"
    rejected.write_text(
        "version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      range: A1:K100000\n",
        encoding="utf-8",
    )
    with pytest.raises(UsageError, match=rf"limit is {MAX_PROFILE_RANGE_CELLS}"):
        load_test_config(rejected)


@pytest.mark.parametrize(
    "range_ref",
    [
        "not-a-range",
        "A1:XFE1",
        "Data!A1:B2",
    ],
)
def test_yaml_profile_rejects_invalid_excel_ranges(
    tmp_path: Path,
    range_ref: str,
) -> None:
    path = tmp_path / "invalid-range.yml"
    path.write_text(
        f"version: 2\nprofile:\n  sheets:\n    - sheet: Data\n      range: {range_ref}\n",
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="configuration is invalid"):
        load_test_config(path)
