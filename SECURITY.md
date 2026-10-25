# Security policy

## Reporting a vulnerability

Please report security problems privately through GitHub: open this repository's
**Security** tab and choose **Report a vulnerability**. Please don't open a public issue
for a security problem.

Include the version (the start banner shows it), what you did, what happened and, if you
can, a short way to reproduce it. Reports are answered as soon as possible.

## Supported versions

Only the latest version on `main` gets fixes.

## Scope

The tool runs with the signed-in admin's own Microsoft 365 rights, so the admin's role is
the real limit on what it can do. In scope:

- anything that makes the tool do more than the admin asked or confirmed: bypassing the
  confirmation gates or the "Run any command" scope, or injection through tenant data or
  CSV files;
- secrets reaching the screen, the logs or the exports;
- how local logs, the undo log and exports are stored;
- the launch and module-install path.

Out of scope: what the admin's Microsoft 365 role itself allows, and vulnerabilities in
Microsoft's or PnP's PowerShell modules (report those to their maintainers).
