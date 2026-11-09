# AGENTS.md

## Project description
This windows-scripts repository contains various scripts
for configuring and maintaining Windows workstations and servers.
The scripts are grouped in folders based on the software or topic they are related to.

## Environment
- Target environment: Windows 11 and Windows Server 2025, unless specified otherwise.

## General instructions
- For scripts that can be feasibly tested, create unit tests in `./tests`.
- This repository is public. Do not commit sensitive information in this repository.
- You may commit directly on the `master` branch. Don't push unless the user explicitly asks you to.

## Coding practices
- New code should be compatible with both PowerShell 7 and PowerShell 5.1.
- Scripts and functions should provide comment-based help, at least `.SYNOPSIS`.
- Define types for function parameters and return values where possible.
- Keep the encoding and line endings as cloned by `git` (CRLF on Windows).
- Use `Show-Output` and `Show-Information` from `Utils.ps1` instead of `Write-Host`.
  Note that `Show-Output` writes to the output stream, so use `Show-Information`
  inside functions that return a value.

## Linting
The scripts are linted with [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer),
configured in `PSScriptAnalyzerSettings.psd1`. Run it locally with:
```
pwsh ./Invoke-Lint.ps1
```
It installs PSScriptAnalyzer for the current user if it is missing.
Apply the automatic formatting fixes with:
```
pwsh ./Invoke-Lint.ps1 -Fix
```
New code should produce no findings at all. The build fails only on `Error` findings,
since the older scripts still have known warnings. Raise the threshold with
`-FailOn Warning` once those have been cleaned up.

The CI workflow runs the same script, so a clean local run means a clean CI run.

## Related repositories
If the repositories below are not available locally when you need them,
`git clone` them in the same parent directory as this repository using SSH.
You may edit the local copies as needed, if no other agent sessions are currently working on them,
and if they don't have large uncommitted changes.

- [agx-ai](https://github.com/AgenttiX/agx-ai): `../agx-ai`, personal AI server and client configs
- [linux-scripts](https://github.com/AgenttiX/linux-scripts): `../linux-scripts`
  - The Linux equivalent of this `windows-scripts` repository.
- [openwrt-scripts](https://github.com/AgenttiX/openwrt-scripts): `../openwrt-scripts`, scripts for OpenWRT-based systems
