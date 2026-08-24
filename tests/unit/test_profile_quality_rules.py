from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from workbooklens.exceptions import UsageError
from workbooklens.models import PatchKind, PatchRisk
from workbooklens.rules.profile_quality import (
    MAX_PROFILE_RANGE_CELLS,
    PROFILE_QUALITY_RULE_IDS,
    PROFILE_QUALITY_RULES,
    ContactFormatRule,
    EnumeratedValueRule,
    LeadingZeroIdentifierRule,
    NumberFormatRoleConflictRule,
    RequiredFieldRule,
    TrailingWhitespaceRule,
    _valid_email,
    parse_workbook_profile,
)
from workbooklens.rules.registry import RuleRegistry
from workbooklens.scanner import ScanResult, scan_workbook


def _save_and_scan(
    workbook: Workbook,
    path: Path,
    *,
    config: dict | None = None,
    rules: tuple[type, ...] = PROFILE_QUALITY_RULES,
) -> ScanResult:
    workbook.save(path)
    workbook.close()
    registry = RuleRegistry(rule_type() for rule_type in rules)
    return scan_workbook(path, config=config, registry=registry)


def _profile(columns: list[dict], **options: object) -> dict:
    return {
        "profile": {
            **options,
            "sheets": [
                {
                    "sheet": "Data",
                    "columns": columns,
                }
            ],
        }
    }


def test_profile_parser_accepts_stable_shape() -> None:
    profile = parse_workbook_profile(
        {
            "profile": {
                "infer_semantics": False,
                "review_trailing_whitespace_patches": True,
                "sheets": [
                    {
                        "sheet": "Data",
                        "columns": [
                            {
                                "header": "Region",
                                "role": "category",
                                "required": True,
                                "allowed_values": ["North", "South"],
                            },
                            {
                                "column": "B",
                                "role": "currency",
                                "preserve_leading_zeros": False,
                            },
                        ],
                    },
                ],
            }
        }
    )

    assert not profile.infer_semantics
    assert profile.review_trailing_whitespace_patches
    assert len(profile.tables) == 1
    table = profile.tables[0]
    assert table.header_row == 1
    assert len(table.columns) == 2
    assert table.columns[0].role == "category"
    assert table.columns[0].required
    assert table.columns[0].allowed_values == ("North", "South")
    assert table.columns[1].column == "B"
    assert table.columns[1].role == "currency"
    assert not table.columns[1].preserve_leading_zeros

    assert parse_workbook_profile({"profile": {}}).tables == ()
    assert parse_workbook_profile(None).tables == ()


@pytest.mark.parametrize(
    ("config", "message"),
    [
        ({"profile": []}, "must be a mapping"),
        ({"profile": {"sheets": "Data"}}, "sheets must be a list"),
        ({"profile": {"sheets": ["Data"]}}, "entry 1 must be a mapping"),
        ({"profile": {"sheets": [{"sheet": ""}]}}, "non-blank sheet name"),
        (
            {"profile": {"sheets": [{"sheet": "Data", "columns": [{"role": "email"}]}]}},
            "exactly one header or column selector",
        ),
        (
            {
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "columns": [{"column": "B", "preserve_leading_zeros": "not-a-bool"}],
                        }
                    ]
                }
            },
            "boolean options",
        ),
        (
            {
                "profile": {
                    "sheets": [
                        {"sheet": "Data", "columns": [{"column": "B", "identifier_width": True}]}
                    ]
                }
            },
            "identifier_width",
        ),
        (
            {"version": 1, "profile": {"sheets": [{"sheet": "Data"}]}},
            "require configuration version 2",
        ),
        (
            {
                "version": 1,
                "workbook_profile": {"tables": [{"sheet": "Data"}]},
            },
            "require configuration version 2",
        ),
        (
            {
                "profile": {"sheets": []},
                "workbook_profile": {"tables": []},
            },
            "cannot define both profile and workbook_profile",
        ),
        (
            {"profile": {"infer_semantcs": False}},
            "unsupported workbook profile keys",
        ),
        (
            {"profile": {"sheets": [{"sheet": "Data", "header_rows": 1}]}},
            "unsupported profile sheet entry 1 keys",
        ),
        (
            {
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "columns": [{"header": "ID", "requried": True}],
                        }
                    ]
                }
            },
            "unsupported profile column entry 1 keys",
        ),
    ],
)
def test_direct_profile_parser_rejects_malformed_contracts(
    config: dict,
    message: str,
) -> None:
    with pytest.raises(ValueError, match=message):
        parse_workbook_profile(config)


def test_profile_rule_exports_are_stable_and_complete() -> None:
    assert {
        "WL036_TRAILING_WHITESPACE",
        "WL037_REQUIRED_FIELD",
        "WL038_ENUM_VALUE",
        "WL039_CONTACT_FORMAT",
        "WL040_LEADING_ZERO_IDENTIFIER",
        "WL041_NUMBER_FORMAT_ROLE_CONFLICT",
    } == PROFILE_QUALITY_RULE_IDS
    assert {rule_type.rule_id for rule_type in PROFILE_QUALITY_RULES} == PROFILE_QUALITY_RULE_IDS


def test_trailing_whitespace_reports_literal_fields_and_only_proposes_opted_in_ascii_patch(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record ID", "Name", "Notes"])
    worksheet.append(["R001", "Alice ", "two internal words"])
    worksheet.append(["R002", "Bob", "Line one\nLine two"])
    worksheet.append(["R003", "Cara\t", "tab suffix"])
    worksheet.append(["R004", "   ", "whitespace-only placeholder"])
    worksheet.append(["R005", '="Dan "', "formula is not a literal field"])

    scan = _save_and_scan(
        workbook,
        tmp_path / "trailing-whitespace.xlsx",
        config=_profile([{"header": "Name", "role": "text", "trim_trailing_whitespace": True}]),
        rules=(TrailingWhitespaceRule,),
    )

    assert {finding.location for finding in scan.findings} == {"B2", "B4"}
    assert len(scan.patches) == 1
    patch = scan.patches[0]
    assert patch.cell == "B2"
    assert patch.kind == PatchKind.SET_TEXT
    assert patch.after == "Alice"
    assert not patch.safe
    assert patch.risk == PatchRisk.LAYOUT_REVIEW
    assert not patch.safe_only_eligible
    assert scan.findings[0].safe_patch_available is False


def test_trailing_whitespace_patch_requires_explicit_review_opt_in(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name"])
    for index, name in enumerate(["Alice ", "Bob", "Cara", "Dan"], start=1):
        worksheet.append([f"R{index:03d}", name])

    scan = _save_and_scan(
        workbook,
        tmp_path / "trailing-no-patch.xlsx",
        config=_profile([{"header": "Name", "role": "text"}]),
        rules=(TrailingWhitespaceRule,),
    )

    assert [finding.location for finding in scan.findings] == ["B2"]
    assert not scan.patches
    assert not scan.findings[0].patch_ids


def test_trailing_whitespace_in_merged_record_anchor_is_report_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name", "Alias"])
    worksheet.append(["R001", "Alice ", None])
    worksheet.merge_cells("B2:C2")

    config = {
        "profile": {
            "review_trailing_whitespace_patches": True,
            "sheets": [
                {
                    "sheet": "Data",
                    "range": "A1:C2",
                    "columns": [
                        {"header": "Name", "role": "text", "trim_trailing_whitespace": True}
                    ],
                }
            ],
        }
    }
    scan = _save_and_scan(
        workbook,
        tmp_path / "merged-trailing-whitespace.xlsx",
        config=config,
        rules=(TrailingWhitespaceRule,),
    )

    assert [finding.location for finding in scan.findings] == ["B2"]
    assert scan.findings[0].evidence.details["merged_cell"] is True
    assert not scan.patches


def test_required_fields_only_flag_populated_records_and_configured_columns(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name", "Comment"])
    worksheet.append(["R001", None, None])
    worksheet.append(["R002", "Bob", None])
    worksheet.append(["R003", "   ", "Name is whitespace"])
    worksheet.append([None, None, None])

    scan = _save_and_scan(
        workbook,
        tmp_path / "required.xlsx",
        config=_profile(
            [
                {"header": "ID", "role": "identifier"},
                {"header": "Name", "role": "text", "required": True},
            ]
        ),
        rules=(RequiredFieldRule,),
    )

    assert {finding.location for finding in scan.findings} == {"B2", "B4"}
    assert all(
        finding.evidence.details["evidence_level"] == "PROVEN_STATIC" for finding in scan.findings
    )
    assert not scan.patches


def test_explicit_profile_scans_a_small_table_without_inference_thresholds(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name"])
    worksheet.append(["R001", None])

    scan = _save_and_scan(
        workbook,
        tmp_path / "small-explicit-profile.xlsx",
        config=_profile([{"header": "Name", "role": "text", "required": True}]),
        rules=(RequiredFieldRule,),
    )

    assert [finding.location for finding in scan.findings] == ["B2"]


def test_explicit_profile_sparse_rows_preserve_required_field_semantics(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name", "Notes"])
    worksheet.append(['="Formula record"', None, None])
    worksheet.append(["Merged record", None, None])
    worksheet.merge_cells("A3:B3")
    worksheet.append(["Single-column record", None, None])
    worksheet.append([None, None, None])

    scan = _save_and_scan(
        workbook,
        tmp_path / "sparse-body-rows.xlsx",
        config={
            "profile": {
                "sheets": [
                    {
                        "sheet": "Data",
                        "range": "A1:C5",
                        "columns": [{"header": "Name", "required": True}],
                    }
                ]
            }
        },
        rules=(RequiredFieldRule,),
    )

    assert [finding.location for finding in scan.findings] == ["B2", "B4"]


def test_direct_profile_config_rejects_ranges_above_shared_limit(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID"])
    worksheet.append(["R001"])

    with pytest.raises(UsageError, match=rf"limit is {MAX_PROFILE_RANGE_CELLS}"):
        _save_and_scan(
            workbook,
            tmp_path / "oversized-profile-range.xlsx",
            config={
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": "A1:K100000",
                            "columns": [{"header": "ID", "required": True}],
                        }
                    ]
                }
            },
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_config_accepts_range_at_shared_limit(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID"])
    worksheet.append(["R001"])

    scan = _save_and_scan(
        workbook,
        tmp_path / "maximum-profile-range.xlsx",
        config={
            "profile": {
                "sheets": [
                    {
                        "sheet": "Data",
                        "range": "A1:J100000",
                        "columns": [{"header": "ID", "required": True}],
                    }
                ]
            }
        },
        rules=(RequiredFieldRule,),
    )

    assert scan.findings == []


def test_direct_profile_parser_rejects_huge_numeric_column_immediately() -> None:
    with pytest.raises(ValueError, match="outside Excel bounds"):
        parse_workbook_profile(
            {
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "columns": [
                                {
                                    "column": 1_000_000_000,
                                    "required": True,
                                }
                            ],
                        }
                    ]
                }
            }
        )


@pytest.mark.parametrize(
    ("range_ref", "message"),
    [
        ("not-a-range", "invalid profile range"),
        ("A1:XFE1", "outside Excel worksheet bounds"),
        ("Data!A1:B2", "must not repeat the sheet name"),
    ],
)
def test_direct_profile_config_rejects_invalid_ranges_instead_of_fallback(
    tmp_path: Path,
    range_ref: str,
    message: str,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID"])
    worksheet.append(["R001"])

    with pytest.raises(UsageError, match=message):
        _save_and_scan(
            workbook,
            tmp_path / f"invalid-profile-range-{range_ref.replace(':', '-')}.xlsx",
            config={
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": range_ref,
                            "columns": [{"header": "ID", "required": True}],
                        }
                    ]
                }
            },
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_range_defaults_header_row_to_first_range_row(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet["A5"] = "ID"
    worksheet["B5"] = "Name"
    worksheet["A6"] = "R001"
    worksheet["A7"] = "R002"
    worksheet["B7"] = "Alice"

    scan = _save_and_scan(
        workbook,
        tmp_path / "range-default-header.xlsx",
        config={
            "profile": {
                "sheets": [
                    {
                        "sheet": "Data",
                        "range": "A5:B7",
                        "columns": [{"header": "Name", "required": True}],
                    }
                ]
            }
        },
        rules=(RequiredFieldRule,),
    )

    assert [finding.location for finding in scan.findings] == ["B6"]


@pytest.mark.parametrize("header_row", [True, 0, 1_048_577, "5"])
def test_direct_profile_rejects_invalid_header_rows(header_row: object) -> None:
    with pytest.raises(ValueError, match="header_row"):
        parse_workbook_profile(
            {
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": "A5:B10",
                            "header_row": header_row,
                        }
                    ]
                }
            }
        )


def test_direct_profile_rejects_header_row_outside_range() -> None:
    with pytest.raises(ValueError, match="inside the configured range"):
        parse_workbook_profile(
            {
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": "A5:B10",
                            "header_row": 1,
                        }
                    ]
                }
            }
        )


def test_direct_profile_rejects_duplicate_table_selectors() -> None:
    with pytest.raises(ValueError, match="duplicate sheet table selectors"):
        parse_workbook_profile(
            {
                "profile": {
                    "sheets": [
                        {"sheet": "Data", "range": "A1:B10"},
                        {"sheet": "data", "range": "A20:B30"},
                    ]
                }
            }
        )


def test_direct_profile_rejects_duplicate_resolved_columns(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name"])
    worksheet.append(["R001", "Alice"])
    worksheet.append(["R002", "Bob"])

    with pytest.raises(UsageError, match="multiple selectors for worksheet column 2"):
        _save_and_scan(
            workbook,
            tmp_path / "duplicate-resolved-column.xlsx",
            config={
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": "A1:B3",
                            "columns": [
                                {"header": "Name", "required": True},
                                {"column": "B", "role": "text"},
                            ],
                        }
                    ]
                }
            },
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_rejects_missing_or_ambiguous_headers(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Name", "Name"])
    worksheet.append(["Alice", "Bob"])

    with pytest.raises(UsageError, match="ambiguous"):
        _save_and_scan(
            workbook,
            tmp_path / "ambiguous-profile-header.xlsx",
            config=_profile([{"header": "Name", "required": True}]),
            rules=(RequiredFieldRule,),
        )

    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Value"])
    worksheet.append(["R001", 1])

    with pytest.raises(UsageError, match="was not found"):
        _save_and_scan(
            workbook,
            tmp_path / "missing-profile-header.xlsx",
            config=_profile([{"header": "Missing", "required": True}]),
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_rejects_missing_worksheet(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID"])
    worksheet.append(["R001"])

    with pytest.raises(UsageError, match="profile worksheet 'Missing' does not exist"):
        _save_and_scan(
            workbook,
            tmp_path / "missing-profile-sheet.xlsx",
            config={
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Missing",
                            "columns": [{"header": "ID", "required": True}],
                        }
                    ]
                }
            },
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_rejects_unresolved_bounds(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"

    with pytest.raises(UsageError, match="table bounds could not be resolved"):
        _save_and_scan(
            workbook,
            tmp_path / "unresolved-profile-bounds.xlsx",
            config=_profile([{"header": "ID", "required": True}]),
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_rejects_blank_header_row(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet["A2"] = "R001"

    with pytest.raises(UsageError, match="does not contain a usable header"):
        _save_and_scan(
            workbook,
            tmp_path / "blank-profile-header.xlsx",
            config={
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": "A1:B2",
                            "columns": [{"column": "A", "required": True}],
                        }
                    ]
                }
            },
            rules=(RequiredFieldRule,),
        )


def test_direct_profile_rejects_column_outside_range(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Name"])
    worksheet.append(["R001", "Alice"])

    with pytest.raises(UsageError, match="column 3 is outside the configured table range"):
        _save_and_scan(
            workbook,
            tmp_path / "outside-profile-column.xlsx",
            config={
                "profile": {
                    "sheets": [
                        {
                            "sheet": "Data",
                            "range": "A1:B2",
                            "columns": [{"column": "C", "required": True}],
                        }
                    ]
                }
            },
            rules=(RequiredFieldRule,),
        )


def test_enumeration_reports_unknown_values_and_only_suggests_unique_near_match(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Region"])
    for row in [
        ["R001", "North"],
        ["R002", "NORTH"],
        ["R003", "Nroth"],
        ["R004", "Other"],
        ["R005", None],
    ]:
        worksheet.append(row)

    scan = _save_and_scan(
        workbook,
        tmp_path / "enumeration.xlsx",
        config=_profile(
            [
                {"header": "ID", "role": "identifier"},
                {
                    "header": "Region",
                    "role": "category",
                    "allowed_values": ["North", "South"],
                },
            ]
        ),
        rules=(EnumeratedValueRule,),
    )

    findings = {finding.location: finding for finding in scan.findings}
    assert set(findings) == {"B4", "B5"}
    assert findings["B4"].evidence.details["near_match"] == "North"
    assert "near_match" not in findings["B5"].evidence.details
    assert not scan.patches


def test_contact_validation_accepts_common_legitimate_forms_and_skips_blanks(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Email", "Phone"])
    worksheet.append(["R001", "alice@example.com", "+86 (10) 1234-5678 ext 9"])
    worksheet.append(["R002", "alice@@example.com", ")1234567("])
    worksheet.append(["R003", None, None])
    worksheet.append(["R004", "bob.smith@sub.example.org", 13800138000])
    worksheet.append(["R005", "first.middle.last+tag@example.co.uk", None])
    worksheet.append(["R006", ".leading@example.com", None])
    worksheet.append(["R007", "trailing.@example.com", None])
    worksheet.append(["R008", "double..dot@example.com", None])

    scan = _save_and_scan(
        workbook,
        tmp_path / "contacts.xlsx",
        config=_profile(
            [
                {"header": "Email", "role": "email"},
                {"header": "Phone", "role": "phone"},
            ]
        ),
        rules=(ContactFormatRule,),
    )

    assert {finding.location for finding in scan.findings} == {
        "B3",
        "C3",
        "B7",
        "B8",
        "B9",
    }
    assert all(float(finding.confidence) == 0.99 for finding in scan.findings)
    assert not scan.patches


def test_email_validation_rejects_ambiguous_dot_input_without_backtracking() -> None:
    assert _valid_email("first.middle.last+tag@example.co.uk")
    assert not _valid_email(".leading@example.com")
    assert not _valid_email("trailing.@example.com")
    assert not _valid_email("double..dot@example.com")
    assert not _valid_email(("!." * 10_000) + "@example.com")

    script = (
        "from workbooklens.rules.profile_quality import _valid_email; "
        "assert not _valid_email(('!.' * 30) + '!')"
    )
    completed = subprocess.run(  # noqa: S603 - fixed interpreter and constant script
        [sys.executable, "-c", script],
        capture_output=True,
        check=False,
        text=True,
        timeout=5,
    )
    assert completed.returncode == 0, completed.stderr


def test_leading_zero_rule_respects_zero_mask_and_non_numeric_identifiers(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Account ID", "Name"])
    worksheet.append([123, "Missing zeros"])
    worksheet.append([456, "Displayed with zeros"])
    worksheet["A3"].number_format = "000000"
    worksheet.append(["000789", "Stored as text"])
    worksheet.append(["000321", "Stored as text"])
    worksheet.append(["000654", "Stored as text"])
    worksheet.append(["000987", "Stored as text"])
    worksheet.append(["0000456", "Extra zero padding"])
    worksheet.append(["A001", "Alphanumeric identifier"])

    scan = _save_and_scan(
        workbook,
        tmp_path / "leading-zero-explicit.xlsx",
        config=_profile(
            [
                {
                    "header": "Account ID",
                    "role": "identifier",
                    "preserve_leading_zeros": True,
                }
            ]
        ),
        rules=(LeadingZeroIdentifierRule,),
    )

    assert [finding.location for finding in scan.findings] == ["A2", "A8"]
    assert {finding.evidence.details["issue_kind"] for finding in scan.findings} == {
        "too_short",
        "extra_leading_zero_padding",
    }
    assert all(finding.evidence.expected["width"] == 6 for finding in scan.findings)
    assert not scan.patches


def test_leading_zero_rule_can_infer_a_dominant_fixed_width_pattern(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record ID", "Name"])
    for identifier in ["0001", "0002", "0003", "0004", 5, "000006"]:
        worksheet.append([identifier, f"Record {identifier}"])

    scan = _save_and_scan(
        workbook,
        tmp_path / "leading-zero-inferred.xlsx",
        rules=(LeadingZeroIdentifierRule,),
    )

    assert [finding.location for finding in scan.findings] == ["A6", "A7"]
    assert {finding.evidence.details["issue_kind"] for finding in scan.findings} == {
        "too_short",
        "extra_leading_zero_padding",
    }
    assert all(
        finding.evidence.details["evidence_level"] == "STRONG_STRUCTURAL"
        for finding in scan.findings
    )


def test_no_profile_does_not_invent_enum_or_variable_width_identifier_constraints(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record ID", "Region"])
    for identifier, region in [
        ("1", "North"),
        ("02", "Nroth"),
        ("003", "Custom"),
        ("AB04", "South"),
        ("50000", "Elsewhere"),
    ]:
        worksheet.append([identifier, region])

    scan = _save_and_scan(
        workbook,
        tmp_path / "no-profile-legitimate-variation.xlsx",
        rules=(EnumeratedValueRule, LeadingZeroIdentifierRule),
    )

    assert not scan.findings
    assert not scan.patches


def test_number_format_rule_reports_clear_role_conflicts_but_not_general_or_valid_formats(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["ID", "Amount", "Rate", "Date"])
    worksheet.append(["R001", 2500, 0.2, datetime(2026, 1, 1)])
    worksheet["B2"].number_format = "0.00%"
    worksheet["C2"].number_format = "$#,##0.00"
    worksheet["D2"].number_format = "0.0%"
    worksheet.append(["R002", 3000, 0.25, datetime(2026, 1, 2)])
    worksheet["B3"].number_format = "¥#,##0.00"
    worksheet["C3"].number_format = "0.0%"
    worksheet["D3"].number_format = "[$-409]mmm d, yyyy"
    worksheet.append(["R003", 3500, 0.3, 46025])

    scan = _save_and_scan(
        workbook,
        tmp_path / "format-roles.xlsx",
        config=_profile(
            [
                {"header": "Amount", "role": "currency"},
                {"header": "Rate", "role": "percentage"},
                {"header": "Date", "role": "date"},
            ]
        ),
        rules=(NumberFormatRoleConflictRule,),
    )

    assert {finding.location for finding in scan.findings} == {"B2", "C2", "D2"}
    assert all(
        finding.evidence.details["evidence_level"] == "PROVEN_STATIC" for finding in scan.findings
    )
    assert not scan.patches


def test_profile_rules_reject_hidden_sheets_when_explicitly_configured(tmp_path: Path) -> None:
    workbook = Workbook()
    visible = workbook.active
    assert visible is not None
    visible.title = "Visible"
    visible.append(["ID", "Name"])
    visible.append(["V001", "Valid"])
    hidden = workbook.create_sheet("Data")
    hidden.sheet_state = "hidden"
    hidden.append(["ID", "Name"])
    hidden.append(["H001", None])

    with pytest.raises(UsageError, match="profile worksheet 'Data' is not visible"):
        _save_and_scan(
            workbook,
            tmp_path / "hidden-profile.xlsx",
            config=_profile([{"header": "Name", "role": "text", "required": True}]),
            rules=(RequiredFieldRule,),
        )
