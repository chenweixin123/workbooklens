"""WorkbookLens command-line interface and stable exit behavior."""

from __future__ import annotations

from pathlib import Path
from typing import Any, NoReturn

import typer
from rich.console import Console
from rich.table import Table

from workbooklens import __version__
from workbooklens.demo import run_demo
from workbooklens.diff import compare_workbooks, write_diff_report
from workbooklens.exceptions import ExitCode, UsageError, WorkbookLensError
from workbooklens.models import SEVERITY_RANK, PatchRisk, Severity
from workbooklens.ooxml.safety import PackageLimits
from workbooklens.policy import apply_finding_policy, load_baseline, source_scope_for_path
from workbooklens.repair import apply_patch_plan, build_patch_plan
from workbooklens.repair.planning import load_patch_plan, write_patch_plan
from workbooklens.reports import write_scan_report
from workbooklens.scanner import scan_workbook
from workbooklens.testing import TestConfig, evaluate_workbook_tests, load_test_config
from workbooklens.utils import write_json
from workbooklens.web import run_local_ui

app = typer.Typer(
    name="workbooklens",
    help="Lint, test, semantically diff, and safely repair Excel workbooks locally.",
    no_args_is_help=True,
    rich_markup_mode="rich",
)
console = Console()


def _version(value: bool) -> None:
    if value:
        console.print(f"WorkbookLens {__version__}")
        raise typer.Exit(ExitCode.OK)


@app.callback()
def main(
    version: bool = typer.Option(
        False,
        "--version",
        callback=_version,
        is_eager=True,
        help="Show the installed version and exit.",
    ),
) -> None:
    """Deterministic workbook quality tooling with no cloud or Excel dependency."""


def _fail(exc: WorkbookLensError) -> NoReturn:
    console.print(f"[bold red]Error:[/bold red] {exc}", style="red")
    raise typer.Exit(exc.exit_code)


def _internal_fail(exc: Exception) -> NoReturn:
    console.print(
        f"[bold red]Internal error:[/bold red] {type(exc).__name__}: {exc}",
        style="red",
    )
    raise typer.Exit(ExitCode.INTERNAL_ERROR)


def _limits(max_file_mb: int) -> PackageLimits:
    if max_file_mb <= 0:
        raise UsageError("--max-file-mb must be positive")
    return PackageLimits(max_file_bytes=max_file_mb * 1024 * 1024)


def _optional_config(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    return load_test_config(path).model_dump(mode="python")


def _optional_test_config(path: Path | None) -> TestConfig | None:
    return load_test_config(path) if path is not None else None


@app.command()
def scan(
    input_workbook: Path = typer.Argument(..., metavar="INPUT.xlsx", help="Workbook to inspect."),
    out: Path = typer.Option(..., "--out", help="Directory for HTML, JSON, snapshot, and SARIF."),
    config: Path | None = typer.Option(
        None, "--config", help="Optional workbooklens YAML configuration."
    ),
    baseline: Path | None = typer.Option(
        None,
        "--baseline",
        help="Previous findings.json or a JSON finding-ID baseline.",
    ),
    new_only: bool = typer.Option(
        False,
        "--new-only",
        help="Report and gate only unsuppressed findings absent from --baseline.",
    ),
    source_scope: str | None = typer.Option(
        None,
        "--source-scope",
        help="Portable logical workbook path for scoped baselines and SARIF; defaults to cwd-relative.",
    ),
    fail_on: Severity | None = typer.Option(
        None,
        "--fail-on",
        case_sensitive=False,
        help="Exit 1 when this severity or higher is found.",
    ),
    max_file_mb: int = typer.Option(100, min=1, help="Maximum compressed input size."),
) -> None:
    """Scan a workbook and emit a self-contained report, findings JSON, snapshot, and SARIF."""

    try:
        if new_only and baseline is None:
            raise UsageError("--new-only requires --baseline")
        configuration = _optional_test_config(config)
        result = scan_workbook(
            input_workbook,
            config=configuration.model_dump(mode="python") if configuration else {},
            limits=_limits(max_file_mb),
        )
        logical_source = source_scope_for_path(input_workbook, explicit=source_scope)
        baseline_ids = (
            load_baseline(baseline, expected_source_scope=logical_source)
            if baseline is not None
            else frozenset()
        )
        policy = apply_finding_policy(
            result.findings,
            suppressions=configuration.suppressions if configuration else [],
            baseline_ids=baseline_ids,
            baseline_path=baseline,
            source_scope=logical_source,
            new_only=new_only,
        )
        paths = write_scan_report(result, out, policy=policy)
        console.print(
            f"[green]Scanned[/green] {input_workbook}: {len(policy.active_findings)} active, "
            f"{len(policy.suppressed_findings)} suppressed, {len(policy.new_findings)} new; "
            f"report [link=file://{paths['html'].resolve()}]{paths['html']}[/link]"
        )
        if policy.expired_suppression_ids:
            console.print(
                "[yellow]Expired suppressions ignored:[/yellow] "
                + ", ".join(policy.expired_suppression_ids)
            )
        if fail_on is not None and any(
            SEVERITY_RANK[finding.severity] >= SEVERITY_RANK[fail_on]
            for finding in policy.active_findings
        ):
            raise typer.Exit(ExitCode.FINDINGS_OR_ASSERTIONS)
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


@app.command()
def plan(
    input_workbook: Path = typer.Argument(..., metavar="INPUT.xlsx", help="Workbook to scan."),
    out: Path = typer.Option(..., "--out", help="Destination repair-plan.json."),
    config: Path | None = typer.Option(
        None, "--config", help="Optional workbooklens YAML configuration."
    ),
    max_file_mb: int = typer.Option(100, min=1, help="Maximum compressed input size."),
) -> None:
    """Create a source-bound reviewable JSON patch manifest without changing the workbook."""

    try:
        if input_workbook.suffix.lower() == ".xlsm":
            raise UsageError(".xlsm inputs are read-only; no repair plan is created")
        result = scan_workbook(
            input_workbook,
            config=_optional_config(config),
            limits=_limits(max_file_mb),
        )
        patch_plan = build_patch_plan(result)
        write_patch_plan(out, patch_plan)
        safe_count = sum(patch.safe_only_eligible for patch in patch_plan.patches)
        layout_review_count = sum(
            patch.risk == PatchRisk.LAYOUT_REVIEW for patch in patch_plan.patches
        )
        console.print(
            f"[green]Planned[/green] {len(patch_plan.patches)} patches "
            f"({safe_count} safe, {layout_review_count} layout review) → {out}"
        )
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


@app.command()
def apply(
    input_workbook: Path = typer.Argument(..., metavar="INPUT.xlsx", help="Original workbook."),
    repair_plan: Path = typer.Argument(..., metavar="repair-plan.json", help="Reviewed plan."),
    out: Path = typer.Option(..., "--out", help="New .xlsx output; must not already exist."),
    patch_id: list[str] | None = typer.Option(
        None,
        "--patch-id",
        help="Selected patch ID; repeat for multiple patches.",
    ),
    safe_only: bool = typer.Option(
        False,
        "--safe-only",
        help="Apply every safe patch at confidence ≥0.95; excludes layout-review patches.",
    ),
    accept_layout_risk: bool = typer.Option(
        False,
        "--accept-layout-risk",
        help="Allow explicitly selected layout-review patches; never enables other unsafe patches.",
    ),
    config: Path | None = typer.Option(
        None, "--config", help="Optional workbooklens YAML configuration."
    ),
    max_file_mb: int = typer.Option(100, min=1, help="Maximum compressed input size."),
) -> None:
    """Apply selected patches to a validated new workbook copy."""

    try:
        plan_model = load_patch_plan(repair_plan)
        result = apply_patch_plan(
            input_workbook,
            plan_model,
            out,
            selected_ids=patch_id,
            safe_only=safe_only,
            accept_layout_risk=accept_layout_risk,
            config=_optional_config(config),
            limits=_limits(max_file_mb),
        )
        report_path = out.with_suffix(out.suffix + ".apply.json")
        write_json(report_path, result.model_dump(mode="json"))
        console.print(
            f"[green]Applied and validated[/green] {len(result.applied_patch_ids)} patches → {out}"
        )
        console.print(f"Apply report: {report_path}")
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


@app.command()
def diff(
    before: Path = typer.Argument(..., metavar="BEFORE.xlsx"),
    after: Path = typer.Argument(..., metavar="AFTER.xlsx"),
    out: Path = typer.Option(..., "--out", help="Destination self-contained .html report."),
) -> None:
    """Compare values, formulas, styles, visibility, merges, names, validations, and sheet structure."""

    try:
        if out.suffix.lower() != ".html":
            raise UsageError("--out for diff must end in .html")
        result = compare_workbooks(before, after)
        paths = write_diff_report(result, out)
        console.print(
            f"[green]Compared[/green] workbooks: {len(result.cell_changes)} cell and "
            f"{len(result.structural_changes)} structural changes → {paths['html']}"
        )
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


@app.command(name="test")
def test_workbook(
    input_workbook: Path = typer.Argument(..., metavar="INPUT.xlsx"),
    config: Path = typer.Option(..., "--config", help="Version-1 or version-2 WorkbookLens YAML."),
    out: Path | None = typer.Option(None, "--out", help="Optional JSON assertion report."),
    max_file_mb: int = typer.Option(100, min=1, help="Maximum compressed input size."),
) -> None:
    """Evaluate bounded user-defined workbook assertions and finding-count gates."""

    try:
        configuration = load_test_config(config)
        run = evaluate_workbook_tests(input_workbook, configuration, _limits(max_file_mb))
        table = Table(title="Workbook assertions")
        table.add_column("Result")
        table.add_column("Assertion")
        table.add_column("Message")
        for result in run.results:
            table.add_row(
                "[green]PASS[/green]" if result.passed else "[red]FAIL[/red]",
                result.assertion_id,
                result.message,
            )
        console.print(table)
        if run.policy.suppressed_findings:
            console.print(
                f"[yellow]{len(run.policy.suppressed_findings)} findings suppressed by "
                "documented waivers.[/yellow]"
            )
        if run.policy.expired_suppression_ids:
            console.print(
                "[yellow]Expired suppressions ignored:[/yellow] "
                + ", ".join(run.policy.expired_suppression_ids)
            )
        if out is not None:
            write_json(
                out,
                {
                    "passed": run.passed,
                    "results": [result.model_dump(mode="json") for result in run.results],
                    "finding_policy": {
                        "summary": run.policy.summary,
                        "suppressed_findings": [
                            item.model_dump() for item in run.policy.suppressed_findings
                        ],
                        "expired_suppression_ids": list(run.policy.expired_suppression_ids),
                    },
                },
            )
        if not run.passed:
            raise typer.Exit(ExitCode.FINDINGS_OR_ASSERTIONS)
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


@app.command()
def serve(
    port: int = typer.Option(
        8765,
        min=0,
        max=65535,
        help="Local TCP port; use 0 to choose an available port automatically.",
    ),
    max_file_mb: int = typer.Option(100, min=1, help="Maximum upload size."),
    open_browser: bool = typer.Option(
        False,
        "--open-browser/--no-open-browser",
        help="Open the local UI after its health check succeeds.",
    ),
    fallback_port: bool = typer.Option(
        False,
        "--fallback-port/--no-fallback-port",
        help="Choose an automatic port only when the requested port is already in use.",
    ),
) -> None:
    """Start the local-only review UI on 127.0.0.1 (never a public interface)."""

    try:
        # Security invariant implemented by run_local_ui: host="127.0.0.1".
        run_local_ui(
            port=port,
            max_file_bytes=max_file_mb * 1024 * 1024,
            open_browser=open_browser,
            fallback_port=fallback_port,
            status=console.print,
        )
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


@app.command()
def demo(
    out: Path = typer.Option(..., "--out", help="Directory for a real before/after walkthrough."),
) -> None:
    """Generate a flawed workbook and run scan, plan, safe apply, reports, and semantic diff."""

    try:
        result = run_demo(out)
        console.print(f"[green]Demo complete[/green] in {result.directory}")
        console.print(f"Before: {result.before_workbook}")
        console.print(f"After:  {result.after_workbook}")
        console.print(f"Plan:   {result.repair_plan}")
        console.print(f"Diff:   {result.diff_html}")
    except typer.Exit:
        raise
    except WorkbookLensError as exc:
        _fail(exc)
    except Exception as exc:
        _internal_fail(exc)


if __name__ == "__main__":
    app()
