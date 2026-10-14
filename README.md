# GridCraft skill for Claude

GridCraft (ArtCraft suite, storytold/gridcraft) is an open-source spreadsheet in Rust, a clean-room reimplementation of **Microsoft Excel**: ribbon, formula bar, Name Box, tables, conditional formatting, charts, PivotTables, printing/PDF. XLSX is the native format. Status: **alpha** (roughly half "ready for real work" per the project ROADMAP; 274/290 ribbon commands, 513/523 functions).

This repository packages a **skill** that teaches Claude to drive GridCraft (part of the [ArtCraft suite](https://github.com/storytold/gridcraft) by storytold) and that **knows how to create the MCP server and guide you through it**, on **Windows and macOS**. It is unofficial and not affiliated with storytold.

> Part of the [ArtCraft Claude skills hub](https://github.com/Vinciux99/artcraft-claude-skills): all 13 skills, a picker and a one-command installer.

## What it does
- Operating manual for GridCraft: connection modes, ground rules, every MCP tool, recipes, CLI, control channel, command families and a safe workflow (work on copies, verify, save only with your OK).
- Built-in MCP setup assistant: Claude checks whether the server is connected, then walks you through installing it (automatic script or manual config), the launcher, authentication and verification.
- Troubleshooting table for the usual problems (server "disconnected", ports, macOS quarantine, Windows Store config location...).

## What you can ask Claude
- "Open `budget.xlsx` (a copy), total each category with SUMIFS and add a clustered column chart next to the table."
- "Clean this CSV: remove duplicates, sort by date descending and turn it into a styled table."
- "Add a PMT-based loan schedule for 30 years at 5% and format it as currency."
- "Highlight every value above 1000 in red with conditional formatting."
- "Build a PivotTable of sales by region and product."
- "Convert `data.xlsx` to CSV and check the XLSX round trip."
- "Write a script that fills formulas down column D and bolds the header, then run it."

## Highlights
- Headless and live MCP, 21 tools plus resources, backed by 313 Excel-style commands and 500+ functions.
- Strict parameter rules, number/formula syntax traps and alpha limitations spelled out (import-only formats, lost chart types).
- CLI pipelines: `run --script`, `convert`, `eval`, `cat`, `check` (Excel-fidelity oracle).
- Full control-channel protocol with real mouse/keyboard input and screenshots.
- Built-in guidance for creating and troubleshooting the MCP server on Windows and macOS.

## App-specific notes
- Kit port **7985** (upstream docs use 7979, which VectorCraft uses here); live MCP `mcp --connect 7985`; headless `mcp`.
- No authentication, loopback only; no read/write roots, so always work on copies and save to new files.
- Formulas always use English syntax with commas, even from non-English systems.
- Printing runs without a dialog: the skill asks before `ui.printSubmit`.

## Get the skill

| Where | How |
|---|---|
| **Claude app** (claude.ai / Desktop) | Download [`gridcraft-skill.zip](https://github.com/Vinciux99/gridcraft-claude-skill/releases/latest/download/gridcraft-skill.zip)`, then Customize > Skills > "+" > Create skill > Upload a skill. Needs Settings > Capabilities > "Code execution and file creation" on. |
| **Claude Code, plugin** | `/plugin marketplace add Vinciux99/artcraft-claude-skills` then `/plugin install gridcraft@artcraft` |
| **Claude Code, manual** | copy `skills/gridcraft/` to `~/.claude/skills/gridcraft/` (the setup script below does it) |

## Create the MCP server

Prerequisite: install GridCraft from [upstream releases](https://github.com/storytold/gridcraft/releases) (Windows `.msi`, macOS `.dmg`) and Claude Desktop and/or Claude Code.

**One command (recommended)**

Windows (PowerShell):
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.ps1))) -Apps gridcraft"
```
macOS (Terminal):
```sh
curl -fsSL https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.sh | bash -s -- --apps gridcraft
```
Add `-DryRun` (Windows) or `--dry-run` (macOS) to preview, and `-Mode live|headless|both` / `--mode ...` to choose the mode.

**Or from a clone:** double-click `setup-windows.cmd` (Windows) or `setup-macos.command` (macOS; first time right-click > Open).

**Or just ask Claude:** once the skill is installed, say "set up the GridCraft MCP server" and Claude will guide you step by step.

After setup: fully quit and reopen Claude, open GridCraft with the **"GridCraft (Claude)"** launcher, then ask Claude to use it. If the server shows "disconnected", the app was not running when Claude started: open it with the launcher and re-enable the server, or use headless mode.

What the setup changes (all backed up or reversible): Claude Desktop config (`mcpServers.gridcraft`, backup `<file>.backup-<timestamp>`), the launcher, Claude Code's user-scope MCP list and `~/.claude/skills/gridcraft/`.

## Repository layout
```
skills/gridcraft/SKILL.md      the skill
app.json                   machine-readable settings (ports, paths, MCP args) used by the scripts
scripts/                   artcraft-setup.ps1 (Windows), artcraft-setup.sh (macOS), json-tool.js
setup-windows.cmd          double-click wrapper
setup-macos.command        double-click wrapper
.claude-plugin/plugin.json Claude Code plugin manifest
.github/workflows/ci.yml   validation, setup tests on Windows and macOS, release zip
```

## CI and releases
Every push validates `app.json`, `SKILL.md` and the scripts, then runs the setup scripts against a fake install on **Windows** and **macOS** runners. Pushing a tag `v*` builds `gridcraft-skill.zip` and publishes a GitHub release.

## Status and caveats
- Windows paths and arguments come from a working install. The macOS paths (`/Applications/GridCraft.app`, CLI location and release asset name) follow the upstream packaging but have **not** been verified on a Mac; the skill marks unverified items and the setup script reports what it could not find.
- GridCraft is a young open-source project; behaviour can change between releases. `list_commands` / `gridcraft-cli commands` is always the source of truth.
- Unofficial community packaging, MIT licensed (see `LICENSE`).
