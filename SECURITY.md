# Security policy

Workbook files are untrusted ZIP/XML inputs. Reports about package parsing, resource exhaustion,
path or relationship handling, temporary files, output preservation, dependencies, GitHub Action
behavior, or the local web UI are especially important.

## Report privately

Use the repository's
[private vulnerability reporting form](https://github.com/chenweixin123/workbooklens/security/advisories/new).
Do not disclose exploit details or attach a sensitive workbook to a public issue.

If GitHub private reporting is unavailable, open a minimal
[public issue](https://github.com/chenweixin123/workbooklens/issues/new) that says only that you
need a private security contact and identifies the affected version. Do not include the
vulnerability, proof of concept, workbook-derived logs, or credentials. Maintainers can then
coordinate a private GitHub Security Advisory.

Include privately:

- affected WorkbookLens version or commit;
- operating system and Python version;
- exact command and observed behavior or resource use;
- a minimized synthetic reproducer when safe;
- whether macros, external links, embedded objects, malformed ZIP/XML, or the GitHub Action are
  involved.

Never send a production workbook. Generate or reduce a package to the smallest non-sensitive
reproducer.

## Supported versions

| Version | Security support |
|---|---|
| 2.2.x | Supported |
| 2.1.x and earlier | Upgrade required |

Security fixes land on main and are included in the next supported patch release. A GitHub
Security Advisory may remain private until users have an upgrade path.

## Security model

WorkbookLens 2.2 does not execute formulas, VBA, embedded objects, or external links. It rejects
packages that exceed entry, compressed/uncompressed size, compression-ratio, member-path,
relationship, or XML limits. DTDs and entities are forbidden. .xlsm remains read-only.

Repairs require a matching source hash, matching target fingerprints, explicit safe patches, an
exact changed-part allowlist, successful reopen, and a clean post-repair scan. Partial output is
removed on failure. The web server binds to 127.0.0.1, rejects non-loopback Host headers and
cross-origin or invalid-token form submissions, and sends restrictive browser security headers.
Uploads stay in a process-owned temporary directory that is removed on normal shutdown; an abrupt
process or operating-system stop can leave files until manual or operating-system cleanup.

Release artifacts are built from an explicit allowlist and audited for unexpected roots,
environments, caches, bytecode, secret-like files, and oversized members. These controls reduce
risk but do not prove that every OOXML parser or transitive dependency is vulnerability-free. Run
WorkbookLens with ordinary user privileges and keep dependencies patched.
