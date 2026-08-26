from __future__ import annotations

import calendar
import random
from datetime import date, datetime
from decimal import Decimal

import pytest
from hypothesis import example, given
from hypothesis import strategies as st
from openpyxl.utils.datetime import to_excel

from workbooklens.rules.profile_quality import (
    _format_chinese_integer,
    _normalization_candidates,
    _parse_chinese_number,
    _parse_currency_text,
    _parse_date_text,
    _parse_decimal_text,
    _parse_percentage_text,
    _parse_semantic_number,
)

_WINDOWS_EPOCH = datetime(1899, 12, 30)
_ASCII_DIGITS = "0123456789"
_CHINESE_DIGITS = "零一二三四五六七八九"


@given(
    sign=st.sampled_from(("", "+", "-")),
    trailing_digits=st.text(alphabet=_ASCII_DIGITS, min_size=1, max_size=12),
    fractional_digits=st.one_of(
        st.none(),
        st.text(alphabet=_ASCII_DIGITS, min_size=1, max_size=3),
    ),
)
def test_leading_zero_numeric_literals_are_rejected_by_all_strict_parsers(
    sign: str,
    trailing_digits: str,
    fractional_digits: str | None,
) -> None:
    literal = f"{sign}0{trailing_digits}"
    if fractional_digits is not None:
        literal = f"{literal}.{fractional_digits}"

    assert _parse_decimal_text(literal) is None
    assert _parse_currency_text(f"${literal}") is None
    assert _parse_currency_text(f"{literal}元") is None
    assert _parse_percentage_text(f"{literal}%") is None
    assert _parse_semantic_number(f"{literal}倍") is None


@pytest.mark.parametrize(
    ("literal", "expected"),
    [
        ("1.5倍", (1.5, "倍")),
        ("-2倍", (-2, "倍")),
        ("+0.5倍", (0.5, "倍")),
        ("1,000倍", (1000, "倍")),
    ],
)
def test_strict_semantic_multiplier_parser_accepts_unique_numeric_values(
    literal: str,
    expected: tuple[int | float, str],
) -> None:
    assert _parse_semantic_number(literal) == expected


@pytest.mark.parametrize(
    "literal",
    ["01.5倍", "1.5x", "1.5倍倍", "NaN倍", "1,23倍", "倍", "1.5 小时"],
)
def test_strict_semantic_multiplier_parser_rejects_ambiguous_or_unapproved_forms(
    literal: str,
) -> None:
    assert _parse_semantic_number(literal) is None


def test_multiplier_and_percentage_candidates_remain_distinct() -> None:
    multiplier = _normalization_candidates("1.5倍", _WINDOWS_EPOCH)
    percentage = _normalization_candidates("1.5%", _WINDOWS_EPOCH)

    assert [(candidate.kind, candidate.after, candidate.semantic) for candidate in multiplier] == [
        ("semantic_number", 1.5, True)
    ]
    assert [(candidate.kind, candidate.after, candidate.semantic) for candidate in percentage] == [
        ("percentage", 0.015, False)
    ]


@given(
    whole=st.integers(min_value=1_000, max_value=999_999_999),
    cents=st.integers(min_value=0, max_value=99),
    marker=st.sampled_from(
        ("$", "€", "£", "¥", "₹", "₩", "₽", "USD", "EUR", "GBP", "CNY", "RMB", "JPY")
    ),
    placement=st.sampled_from(("prefix", "suffix")),
    separator=st.sampled_from(("", " ")),
)
def test_currency_symbols_preserve_valid_grouped_decimal_values(
    whole: int,
    cents: int,
    marker: str,
    placement: str,
    separator: str,
) -> None:
    body = f"{whole:,}.{cents:02d}"
    literal = (
        f"{marker}{separator}{body}" if placement == "prefix" else f"{body}{separator}{marker}"
    )

    parsed = _parse_currency_text(literal)

    assert parsed is not None
    assert Decimal(str(parsed)) == Decimal(body.replace(",", ""))


@given(scaled_percent=st.integers(min_value=-100_000, max_value=100_000))
@example(scaled_percent=-10_000)
@example(scaled_percent=0)
@example(scaled_percent=1)
@example(scaled_percent=100)
@example(scaled_percent=10_000)
def test_percentage_parser_applies_exactly_one_percent_multiplier(
    scaled_percent: int,
) -> None:
    sign = "-" if scaled_percent < 0 else ""
    magnitude = abs(scaled_percent)
    whole, hundredths = divmod(magnitude, 100)
    literal = f"{sign}{whole}.{hundredths:02d}%"

    parsed = _parse_percentage_text(literal)

    assert parsed is not None
    assert Decimal(str(parsed)) == Decimal(scaled_percent) / Decimal(10_000)


@st.composite
def _valid_date_literals(draw: st.DrawFn) -> tuple[str, date]:
    parsed = draw(
        st.dates(
            min_value=date(1900, 1, 1),
            max_value=date(2099, 12, 31),
        )
    )
    representation = draw(st.sampled_from(("dash", "slash", "dot", "chinese")))
    if representation == "chinese":
        literal = f"{parsed.year}年{parsed.month}月{parsed.day}日"
    else:
        separator = {"dash": "-", "slash": "/", "dot": "."}[representation]
        literal = f"{parsed.year}{separator}{parsed.month}{separator}{parsed.day}"
    return literal, parsed


@given(case=_valid_date_literals())
def test_valid_date_literals_map_to_the_workbook_epoch_serial(
    case: tuple[str, date],
) -> None:
    literal, parsed_date = case

    parsed = _parse_date_text(literal, _WINDOWS_EPOCH)

    assert parsed == int(to_excel(parsed_date, _WINDOWS_EPOCH))  # type: ignore[no-untyped-call]


@st.composite
def _invalid_date_literals(draw: st.DrawFn) -> str:
    year = draw(st.integers(min_value=1900, max_value=2099))
    month = draw(st.integers(min_value=1, max_value=12))
    invalid_day = calendar.monthrange(year, month)[1] + 1
    representation = draw(st.sampled_from(("dash", "slash", "dot", "chinese")))
    if representation == "chinese":
        return f"{year}年{month}月{invalid_day}日"
    separator = {"dash": "-", "slash": "/", "dot": "."}[representation]
    return f"{year}{separator}{month}{separator}{invalid_day}"


@given(literal=_invalid_date_literals())
def test_calendar_invalid_date_literals_are_rejected(literal: str) -> None:
    assert _parse_date_text(literal, _WINDOWS_EPOCH) is None
    assert not any(
        candidate.kind == "date" for candidate in _normalization_candidates(literal, _WINDOWS_EPOCH)
    )


@pytest.mark.parametrize(
    "literal",
    [
        "1899-12-29",
        "2026-01/02",
        "2026/01.02",
        "2026-00-10",
        "2026-13-10",
    ],
)
def test_date_parser_rejects_pre_epoch_and_mixed_separator_boundaries(literal: str) -> None:
    assert _parse_date_text(literal, _WINDOWS_EPOCH) is None


@given(value=st.integers(min_value=0, max_value=999_999_999_999_999))
def test_chinese_digit_sequences_respect_the_fifteen_digit_storage_boundary(
    value: int,
) -> None:
    literal = "".join(_CHINESE_DIGITS[int(digit)] for digit in str(value))

    assert _parse_chinese_number(literal) == (value, "")
    candidates = _normalization_candidates(literal, _WINDOWS_EPOCH)
    assert len(candidates) == 1
    assert candidates[0].after == value
    assert candidates[0].kind == "chinese_number"
    assert candidates[0].semantic


def test_chinese_digit_sequence_beyond_exact_storage_boundary_is_rejected() -> None:
    literal = "一" + ("零" * 15)

    assert _parse_chinese_number(literal) is None
    assert not _normalization_candidates(literal, _WINDOWS_EPOCH)


@given(
    leading_zero=st.sampled_from(("零", "〇")),  # noqa: RUF001
    trailing_digits=st.text(
        alphabet="零〇一二两三四五六七八九",
        min_size=1,
        max_size=14,
    ),
)
def test_multi_digit_chinese_digit_sequences_cannot_start_with_zero(
    leading_zero: str,
    trailing_digits: str,
) -> None:
    literal = f"{leading_zero}{trailing_digits}"

    assert _parse_chinese_number(literal) is None
    assert not _normalization_candidates(literal, _WINDOWS_EPOCH)


@pytest.mark.parametrize(
    "literal",
    [
        "一百百",
        "十十",
        "零一",
        "一万万",
        "一百二",
        "一万三",
        "零十",
        "十零",
    ],
)
def test_noncanonical_chinese_unit_sequences_are_rejected(literal: str) -> None:
    assert _parse_chinese_number(literal) is None
    assert not _normalization_candidates(literal, _WINDOWS_EPOCH)


@pytest.mark.parametrize(
    "literal",
    [
        "两十",
        "一百两十",
        "两三",
        "二两",
        "一百零两",
        "两百零两",
        "一点两",
        "负一点两",
    ],
)
def test_liang_is_rejected_outside_standalone_or_unit_prefix_positions(
    literal: str,
) -> None:
    assert _parse_chinese_number(literal) is None
    assert not _normalization_candidates(literal, _WINDOWS_EPOCH)


@pytest.mark.parametrize(
    ("literal", "expected"),
    [
        ("十", 10),
        ("一百零二", 102),
        ("一万零三", 10_003),
        ("一亿零一万", 100_010_000),
        ("一亿零一万零一", 100_010_001),
        ("一万亿", 1_000_000_000_000),
        ("一万三千亿五千万", 1_300_050_000_000),
        ("一万二千三百四十五亿六千七百八十九万零一百二十三", 1_234_567_890_123),
        ("两千零三", 2_003),
        ("两", 2),
        ("负两", -2),
        ("两点五", 2.5),
        ("两百", 200),
        ("两千", 2_000),
        ("两万", 20_000),
        ("两亿", 200_000_000),
        ("两百零三", 203),
        ("一万两千", 12_000),
        ("两万两千", 22_000),
        ("两亿零两万", 200_020_000),
    ],
)
def test_canonical_chinese_unit_sequences_remain_supported(
    literal: str,
    expected: int,
) -> None:
    assert _parse_chinese_number(literal) == (expected, "")


@given(value=st.integers(min_value=0, max_value=999_999_999_999_999))
def test_canonical_chinese_unit_formatter_round_trips(value: int) -> None:
    literal = _format_chinese_integer(value)

    assert literal is not None
    assert _parse_chinese_number(literal) == (value, "")


def test_chinese_formatter_round_trips_exact_storage_upper_boundary() -> None:
    value = 999_999_999_999_999
    literal = _format_chinese_integer(value)

    assert literal == "九百九十九万九千九百九十九亿九千九百九十九万九千九百九十九"
    assert _parse_chinese_number(literal) == (value, "")


def test_chinese_formatter_round_trips_ten_thousand_seeded_values() -> None:
    random_values = random.Random(0x574C240).sample(range(10**15), 10_000)  # noqa: S311

    for value in random_values:
        literal = _format_chinese_integer(value)

        assert literal is not None
        assert _parse_chinese_number(literal) == (value, "")


def test_numeric_text_with_trailing_spaces_keeps_both_role_candidates() -> None:
    candidates = _normalization_candidates("123  ", _WINDOWS_EPOCH)

    assert {(candidate.kind, candidate.after) for candidate in candidates} == {
        ("number", 123),
        ("trailing_whitespace", "123"),
    }


@given(
    digit=st.integers(min_value=1, max_value=9),
    unit=st.sampled_from(("个", "人", "件", "元", "块", "小时", "天", "次", "公斤", "米")),
)
def test_whitelisted_chinese_units_remain_semantic_review_candidates(
    digit: int,
    unit: str,
) -> None:
    literal = f"{_CHINESE_DIGITS[digit]}{unit}"
    candidates = _normalization_candidates(literal, _WINDOWS_EPOCH)

    assert len(candidates) == 1
    assert candidates[0].after == digit
    assert candidates[0].kind == "chinese_number"
    assert candidates[0].semantic
    assert candidates[0].unit == unit


@pytest.mark.parametrize(
    ("literal", "expected_value", "expected_unit"),
    [
        ("一千克", 1, "千克"),
        ("一万元", 10_000, "万元"),
        ("八万九千元", 89_000, "元"),
    ],
)
def test_chinese_suffixes_starting_with_numeral_units_are_not_consumed(
    literal: str,
    expected_value: int,
    expected_unit: str,
) -> None:
    candidates = _normalization_candidates(literal, _WINDOWS_EPOCH)

    assert len(candidates) == 1
    assert candidates[0].after == expected_value
    assert candidates[0].unit == expected_unit
    assert candidates[0].semantic


@given(
    digit=st.integers(min_value=1, max_value=9),
    measure_word=st.sampled_from(("张", "本", "台", "辆", "盒", "份", "岁", "吨", "页", "册")),
)
def test_non_whitelisted_chinese_measure_words_are_rejected(
    digit: int,
    measure_word: str,
) -> None:
    literal = f"{_CHINESE_DIGITS[digit]}{measure_word}"

    assert _parse_chinese_number(literal) is None
    assert not _normalization_candidates(literal, _WINDOWS_EPOCH)
