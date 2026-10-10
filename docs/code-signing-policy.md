# Code signing policy

Free code signing provided by [SignPath.io](https://signpath.io), certificate by [SignPath Foundation](https://signpath.org).

*Status: application in progress. Releases up to 0.9.1 are unsigned; this policy applies from the first signed release.*

Revit Ballet installers and add-in assemblies are signed so that Windows and Revit can show who published them and so that tampering is detectable.

## What is signed

- `installer.exe`
- `revit-ballet.dll` for every supported Revit year, and the script engine assemblies (`revit-ballet-engine-*.dll`, `revit-ballet.agents.abstractions.dll`)

Only binaries built from the source in this repository are signed. Upstream libraries that are bundled (Roslyn, IronPython, pythonnet, the Python runtime, SQLite, ...) are not re-signed; they keep the signatures of their own publishers where they have one.

## How a release is signed

1. A release is built from a tagged commit by the GitHub Actions workflow in `.github/workflows/`. Build scripts and workflow files are part of the source.
2. The workflow uploads the unsigned build to the signing service together with proof of its origin (repository, commit, workflow run).
3. A maintainer approves each signing request by hand. Nothing is signed automatically.
4. The signed files are attached to the GitHub release.

## Roles

| Role | Who | Responsibility |
|---|---|---|
| Authors | [@baleti](https://github.com/baleti) | trusted to change the source without further review |
| Reviewers | [@baleti](https://github.com/baleti) | review changes from contributors who are not authors |
| Approvers | [@baleti](https://github.com/baleti) | approve every signing request |

All team members use multi-factor authentication on GitHub and on the signing service.

## Privacy

Revit Ballet does not transfer any information to other networked systems unless the user or the person installing it specifically asks for it. Its only network listener is a local server bound to `127.0.0.1`, protected by a token, that scripts and tools on the same machine use to talk to a running Revit.

## System changes made by the installer

The installer announces and performs these changes, per user and without administrator rights:

- copies the add-in and its optional script engines to `%APPDATA%\Autodesk\Revit\Addins\<year>\revit-ballet\` and writes `revit-ballet.addin` beside it
- creates runtime data under `%APPDATA%\revit-ballet\`
- merges a set of default shortcuts into Revit's `KeyboardShortcuts.xml` for each installed Revit version
- registers an uninstaller under `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\RevitBallet` and the add-in's trust entries

`installer.exe /uninstall` (or Windows Settings > Apps) removes the add-in again.

## Reporting a problem

Open an issue at https://github.com/baleti/revit-ballet/issues. Reports about misuse of the signing certificate are investigated with priority.
