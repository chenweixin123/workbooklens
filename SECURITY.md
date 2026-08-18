# Security policy

Workbook files are untrusted ZIP/XML inputs. Security reports about package parsing, resource
exhaustion, path/relationship handling, temporary files, output preservation, or the local web UI
are especially important.

## Reporting

Use GitHub's private vulnerability-reporting form for the repository. Do not open a public issue
until maintainers have coordinated a fix. Include:

- affected WorkbookLens version or commit;
- operating system and Python version;
- exact command and observed behavior/resource use;
- a minimized synthetic reproducer when safe;
- whether the workbook contains macros, external links, embedded objects, or malformed ZIP/XML.

Never send a sensitive production workbook. Reduce it to the smallest generated package that
reproduces the flaw, or describe how maintainers can create one.

## Supported versions

Until the first stable release, security fixes are made on `main` and included in the next v0.x
release. After 1.0, the project will publish a version-support table here.

## Security model

WorkbookLens v0.1 does not execute formulas, VBA, embedded objects, or external links. It rejects
packages that exceed configured entry, compressed/uncompressed size, compression ratio, member
path, relationship, or XML limits. DTDs and entities are forbidden. `.xlsm` is read-only.

Repairs require a matching source hash, matching target fingerprints, explicit safe patches, an
exact changed-part allowlist, successful reopen, and a clean post-repair scan. Partial outputs are
removed on validation failure. The web server binds to `127.0.0.1` and keeps uploads under a
process-owned temporary directory.

These controls reduce risk but do not prove that every OOXML parser dependency is vulnerability
free. Run WorkbookLens with ordinary user privileges and keep dependencies patched.
