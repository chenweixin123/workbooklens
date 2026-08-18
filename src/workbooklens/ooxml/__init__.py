"""Security-focused OOXML package inspection and patching."""

from workbooklens.ooxml.safety import PackageInspection, PackageLimits, inspect_package

__all__ = ["PackageInspection", "PackageLimits", "inspect_package"]
