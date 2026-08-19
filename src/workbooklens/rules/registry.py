"""Rule registry with an optional Python entry-point integration surface."""

from __future__ import annotations

from collections.abc import Iterable
from importlib.metadata import entry_points

from workbooklens.rules.base import WorkbookRule


class RuleRegistry:
    """Ordered registry that rejects duplicate stable rule identifiers."""

    def __init__(self, rules: Iterable[WorkbookRule] = ()) -> None:
        self._rules: dict[str, WorkbookRule] = {}
        for rule in rules:
            self.register(rule)

    def register(self, rule: WorkbookRule) -> WorkbookRule:
        """Register and return a rule, enabling decorator-style use."""

        if not rule.rule_id or rule.rule_id in self._rules:
            raise ValueError(f"duplicate or empty rule ID: {rule.rule_id!r}")
        self._rules[rule.rule_id] = rule
        return rule

    def values(self) -> tuple[WorkbookRule, ...]:
        """Return rules sorted by stable identifier for deterministic scans."""

        return tuple(self._rules[key] for key in sorted(self._rules))

    def load_entry_points(self, group: str = "workbooklens.rules") -> None:
        """Load explicitly installed rule plugins from the documented entry-point group."""

        for entry_point in entry_points(group=group):
            loaded = entry_point.load()
            rule = loaded() if isinstance(loaded, type) else loaded
            if not isinstance(rule, WorkbookRule):
                raise TypeError(f"entry point {entry_point.name!r} did not provide a WorkbookRule")
            self.register(rule)


def default_registry() -> RuleRegistry:
    """Construct a fresh registry containing all built-in rules."""

    from workbooklens.rules.builtin import BUILTIN_RULES

    return RuleRegistry(rule_type() for rule_type in BUILTIN_RULES)
