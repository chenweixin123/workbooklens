WorkbookLens @VERSION@ - Windows x64 Portable
================================================

This package runs entirely on your Windows computer. Workbook files are sent
only to the local WorkbookLens process listening on 127.0.0.1. They are not
uploaded to a WorkbookLens cloud service.

Requirements
------------

- 64-bit Windows 10 or Windows 11.
- No separate Python installation is required.

Start the desktop application
-----------------------------

1. Extract the complete ZIP archive. Do not run files from inside the ZIP.
2. Double-click WorkbookLens.exe.
3. WorkbookLens opens in its own desktop window. No console or external browser
   is required. Closing the window stops its private local service.

Start-WorkbookLens.cmd remains as a compatibility shortcut and opens the same
desktop application. The application first tries port 8765. If that port is
already in use, it automatically selects a free local port. It binds only to
127.0.0.1 and displays that local service inside the native application window.

Command-line use
----------------

Open Command Prompt in this directory and run:

  WorkbookLensCLI.exe --help
  WorkbookLensCLI.exe --version
  WorkbookLensCLI.exe demo --out demo-output

Configuration
-------------

workbooklens.example.yml is an example configuration. Copy it to a separate,
writable working directory before editing it.

Integrity and Windows warnings
------------------------------

Verify the SHA256 sidecar before extracting the archive. This initial portable
build is not code-signed, so Windows SmartScreen may display an unfamiliar-app
warning. Do not disable antivirus protection to run WorkbookLens.

Licenses
--------

LICENSE covers WorkbookLens. THIRD-PARTY-NOTICES.txt and LICENSES contain the
licenses and notices for the bundled Python runtime and third-party packages.
