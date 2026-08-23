"""English and Simplified Chinese presentation APIs."""

from workbooklens.i18n._rule_catalog import BUILTIN_RULE_TITLES
from workbooklens.i18n.catalog import (
    DEFAULT_LANGUAGE,
    SUPPORTED_LANGUAGES,
    Language,
    assert_catalog_complete,
    message_catalog,
    normalize_language,
    require_translation,
    translate,
)
from workbooklens.i18n.cli import localize_assertion_message
from workbooklens.i18n.errors import (
    PUBLIC_ERROR_CODES,
    UserFacingError,
    assert_error_catalog_complete,
    localize_exception,
    localized_error,
    new_diagnostic_id,
)
from workbooklens.i18n.rules import (
    MissingBuiltinTranslationError,
    canonical_text_translation,
    change_type_label,
    localize_finding,
    localize_patch,
    localize_scan_result,
    patch_kind_label,
    risk_label,
    severity_label,
)

__all__ = [
    "BUILTIN_RULE_TITLES",
    "DEFAULT_LANGUAGE",
    "PUBLIC_ERROR_CODES",
    "SUPPORTED_LANGUAGES",
    "Language",
    "MissingBuiltinTranslationError",
    "UserFacingError",
    "assert_catalog_complete",
    "assert_error_catalog_complete",
    "canonical_text_translation",
    "change_type_label",
    "localize_assertion_message",
    "localize_exception",
    "localize_finding",
    "localize_patch",
    "localize_scan_result",
    "localized_error",
    "message_catalog",
    "new_diagnostic_id",
    "normalize_language",
    "patch_kind_label",
    "require_translation",
    "risk_label",
    "severity_label",
    "translate",
]
