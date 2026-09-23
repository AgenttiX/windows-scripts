# AGENTS.md

## Project description
This windows-scripts repository contains various scripts
for configuring and maintaining Windows workstations and servers.

## Environment
- Target environment: Windows 11 and Windows Server 2025, unless specified otherwise.

## Coding practices
- New code should be compatible with both PowerShell 7 and PowerShell 5.1.
- Scripts and functions should provide comment-based help, at least `.SYNOPSIS`.
- Define types for function parameters and return values where possible.
- On Windows script files are UTF-8 with a BOM and CRLF line endings (as cloned by `git`).
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
