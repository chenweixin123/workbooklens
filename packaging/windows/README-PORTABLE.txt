WorkbookLens @VERSION@ - Windows x64 Portable
================================================

This package runs entirely on your Windows computer. Workbook files are sent
only to the local WorkbookLens process listening on 127.0.0.1. They are not
uploaded to a WorkbookLens cloud service.

Requirements
------------

- 64-bit Windows 10 or Windows 11.
- No separate Python installation is required.

Start the local web interface
-----------------------------

1. Extract the complete ZIP archive. Do not run files from inside the ZIP.
2. Double-click Start-WorkbookLens.cmd.
3. A dedicated WorkbookLens console window opens. Your default browser opens
   after the local service is ready. If it does not,
   open the local URL printed in the console window.
4. Keep the WorkbookLens console window open while using WorkbookLens.
5. Press Ctrl+C in that window to stop the local server.

The launcher first tries port 8765. If that port is already in use, WorkbookLens
automatically selects a free local port and opens that exact address. It binds
only to 127.0.0.1 and will not open or trust an unrelated service on port 8765.

Command-line use
----------------

Open Command Prompt in this directory and run:

  WorkbookLens.exe --help
  WorkbookLens.exe --version
  WorkbookLens.exe demo --out demo-output

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
