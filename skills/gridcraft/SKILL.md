---
name: gridcraft
description: Read, create, format and export spreadsheets (.xlsx, .csv) with GridCraft (Excel-style, ArtCraft suite) through MCP, the CLI or the control channel, and set up, connect or troubleshoot its MCP server on Windows and macOS.
---

# GridCraft for Claude

GridCraft (ArtCraft suite, storytold/gridcraft) is an open-source spreadsheet in Rust, a clean-room reimplementation of **Microsoft Excel**: ribbon, formula bar, Name Box, tables, conditional formatting, charts, PivotTables, printing/PDF. XLSX is the native format. Status: **alpha** (roughly half "ready for real work" per the project ROADMAP; 274/290 ribbon commands, 513/523 functions).

**Every action is a command** with a dotted id and JSON parameters: `Session::execute(id, params)`. Ids follow Excel's ribbon (`home.bold`, `insert.chart`, `data.sortAscending`) plus a few primitives (`cell.set`, `selection.set`, `range.setValues`). UI, CLI, control channel and MCP all reach the same registry (313 commands). **Programmatic calls never open dialogs**: with empty params a command uses defaults or returns an error. The source of truth is always `list_commands` / `gridcraft-cli commands`.

## Quick check: is the MCP server connected?

1. Look at the tools you have. If tools from the MCP server `gridcraft` are available (for example `list_commands`, `execute` or `inspect_*` tools) the server is connected: skip to "Operating knowledge" below and use them directly.
2. If they are missing, check whether the user installed it under the name `gridcraft` or `gridcraft-headless`. In Claude Desktop the server may show as **disconnected**.
3. **Most common cause of "disconnected":** the live server (`mcp --connect {port}`) exits when GridCraft is not running at the moment Claude starts. Fix: open GridCraft with the **"GridCraft (Claude)" launcher** (it starts the app with its control channel on), then re-enable the server in Claude's settings or fully quit and reopen Claude.
4. If the user does not need to watch the app live, **headless mode** always starts (no window, works on files). See "Live, headless or both".
5. If nothing is configured yet, follow "Set up the MCP server" below. Ask the user which OS they use if you do not know.

## Set up the MCP server (guide the user)

Walk the user through this step by step. Ask for the OS (Windows or macOS) first. If you have a shell on the user's machine you may run the steps yourself, but **ask before editing any config file and always back it up first**. Never print tokens or keys in chat.

### 1. Prerequisites
- Install GridCraft from its upstream releases page: https://github.com/storytold/gridcraft/releases (Windows: the `.msi` installer; macOS: the `.dmg`, then drag the app to `/Applications`).
- Install Claude Desktop and/or Claude Code. For Claude Desktop, fully quit and reopen it after any config change.
- The MCP server is the **CLI** that ships with the app: `gridcraft-cli`. Windows installs it next to the app (`C:\Program Files\GridCraft\gridcraft-cli.exe`, not on `PATH`). On macOS the CLI is a separate release download (asset `gridcraft-cli-<version>-macos-universal.zip`) unless it is bundled in the app (unverified): the setup script below fetches it into `~/.artcraft/bin/` and clears the macOS quarantine flag.
- Upstream docs use port 7979; this kit uses **7985** so that all 13 ArtCraft apps can run together without clashing.

### 2. Option A: automatic setup (recommended)
The setup script finds the app and CLI, creates the launcher, backs up and updates Claude Desktop's config, registers the server in Claude Code (if the `claude` command exists) and installs this skill. Preview first with a dry run: nothing is changed.

Windows (PowerShell):
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.ps1))) -Apps gridcraft"
```
macOS (Terminal):
```sh
curl -fsSL https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.sh | bash -s -- --apps gridcraft
```
Or clone/download this repository and double-click `setup-windows.cmd` (Windows) or `setup-macos.command` (macOS; first time: right-click, Open). Useful options: `-Mode live|headless|both` / `--mode ...` (default `live`; `both` registers `gridcraft` live and `gridcraft-headless`), `-DryRun` / `--dry-run`, `-NoLauncher`, `-NoClaudeDesktop`, `-NoClaudeCode`, `-NoSkill` (macOS: `--no-launcher`, `--no-claude-desktop`, `--no-claude-code`, `--no-skill`).

What it changes: Claude Desktop config (`%APPDATA%\Claude\claude_desktop_config.json` plus the Microsoft Store/MSIX copy under `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\`; macOS `~/Library/Application Support/Claude/claude_desktop_config.json`; backup saved next to it as `<file>.backup-<timestamp>`), the launcher "GridCraft (Claude)" (Windows: Desktop shortcut; macOS: `~/Applications/GridCraft (Claude).app` plus a Desktop alias), Claude Code's user-scope MCP list, and `~/.claude/skills/gridcraft/`.

### 3. Option B: manual setup
**Windows** - edit `%APPDATA%\Claude\claude_desktop_config.json` (if Claude was installed from the Microsoft Store, also the copy under `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\`). Keep other servers; add this inside `mcpServers`. Backslashes must be doubled in JSON.
```json
{
  "mcpServers": {
    "gridcraft": {
      "command": "C:\\Program Files\\GridCraft\\gridcraft-cli.exe",
      "args": [
        "mcp",
        "--connect",
        "7985"
      ]
    },
    "gridcraft-headless": {
      "command": "C:\\Program Files\\GridCraft\\gridcraft-cli.exe",
      "args": [
        "mcp"
      ]
    }
  }
}
```
**macOS** - edit `~/Library/Application Support/Claude/claude_desktop_config.json`. Use the absolute CLI path (replace `YOUR_NAME`; JSON does not expand `~`). The CLI path is unverified if your install differs: check with `ls ~/.artcraft/bin /Applications/GridCraft.app/Contents/MacOS`.
```json
{
  "mcpServers": {
    "gridcraft": {
      "command": "/Users/YOUR_NAME/.artcraft/bin/gridcraft-cli",
      "args": [
        "mcp",
        "--connect",
        "7985"
      ]
    },
    "gridcraft-headless": {
      "command": "/Users/YOUR_NAME/.artcraft/bin/gridcraft-cli",
      "args": [
        "mcp"
      ]
    }
  }
}
```
**Launcher** (starts the app with the control channel; needed before the live server can connect):
- Windows: create a Desktop shortcut named "GridCraft (Claude)" with target `"C:\Program Files\GridCraft\gridcraft.exe" --control 7985`
- macOS: run `open -n -a "/Applications/GridCraft.app" --args --control 7985`, or save it as an app with Script Editor/`osacompile` (the setup script does this for you).

**Claude Code** (any OS; run once):
Windows (PowerShell):
```powershell
claude mcp add --scope user gridcraft -- "C:\Program Files\GridCraft\gridcraft-cli.exe" mcp --connect 7985
claude mcp add --scope user gridcraft-headless -- "C:\Program Files\GridCraft\gridcraft-cli.exe" mcp
```
macOS (Terminal):
```sh
claude mcp add --scope user gridcraft -- "$HOME/.artcraft/bin/gridcraft-cli" mcp --connect 7985
claude mcp add --scope user gridcraft-headless -- "$HOME/.artcraft/bin/gridcraft-cli" mcp
```
Check with `claude mcp list`.

### 4. Authentication
GridCraft needs no token: the control channel listens on loopback (`127.0.0.1`) only and has no authentication. Anyone on the machine can drive the app while it is open, so close it when you are done.

### 5. Finish and verify
1. **Fully quit and reopen Claude** (Desktop: quit from the tray/menu bar, not just close the window).
2. Open GridCraft with the **"GridCraft (Claude)" launcher** before asking Claude to use it.
3. In Claude, confirm the server `gridcraft` is connected and call its `list_commands` tool (or equivalent). A non-empty command list means it works.

### Live, headless or both
- **Live** (`gridcraft`): drives the open app; the user sees every change and screenshots/UI tools work. Needs the launcher first.
- **Headless** (`gridcraft-headless`): in-process engine, no window; always starts; works on files; UI/screenshot tools are not available.
- **Both**: register both and pick per task. Switch with `-Mode` / `--mode`, or by editing the config entries above.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Server shows "disconnected" right after Claude starts | Live server exits when GridCraft is not running | Open GridCraft with the "GridCraft (Claude)" launcher, then re-enable the server or restart Claude; or use headless mode |
| "Port already in use" / app opens but Claude cannot connect | Another app or an old GridCraft instance holds port 7985 | Quit all GridCraft windows, check the launcher uses the kit port, make sure no other ArtCraft app uses the same port |
| `command not found` / spawn error | CLI path wrong or the CLI is not installed (Windows: not on `PATH`; macOS: separate download) | Use the full absolute path in the config; macOS: run the setup script (it downloads the CLI to `~/.artcraft/bin/`) |
| macOS: "cannot be opened because the developer cannot be verified" | Gatekeeper/quarantine on the app or CLI | Right-click the app, Open; for the CLI: `xattr -dr com.apple.quarantine ~/.artcraft/bin/gridcraft-cli` |
| Anyone else can drive the app | The control channel has no authentication (loopback only) | Start the app with the control channel only when needed and quit it when done |
| Config edited but nothing changed (Windows) | Claude installed from the Microsoft Store reads the MSIX copy of the config | Edit/update `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\claude_desktop_config.json` too |
| JSON config error / server not listed | Unescaped backslashes or trailing comma | Double the backslashes on Windows; validate the JSON; restore the `.backup-*` file if needed |
| App crashes on start or window is blank | Graphics backend problem | Windows: set `WGPU_BACKEND=vulkan` (or `dx12`) for one launch; avoid `gl` on some AMD drivers |
| Changes do not appear in the window | You are on the headless server | Use the live server (launcher first) |
| Quotes disappear in PowerShell JSON arguments | Windows PowerShell 5.1 quoting for `.exe` | Use `--script`/files where the CLI supports them, or single-quote the whole JSON |

## Operating knowledge

## App, settings, environment
- **Log:** `<settings>/logs/gridcraft.log` (Windows `%APPDATA%\GridCraft\logs`). At every launch the previous log rotates to `gridcraft.1.log`, then `gridcraft.2.log`; cap 16 MiB. Settings live in the folder that contains `logs`.
- Other app options: `--sample sales|budget|grades`, or a file as a positional argument.
- **Environment variables:**
  - `GRIDCRAFT_CONTROL_PORT=7985` is the same as `--control 7985`.
  - `GRIDCRAFT_NO_PREFS` starts the app without saved settings (logging goes to stderr only).
  - `RUST_LOG` sets log level (e.g. `RUST_LOG=debug` or `warn,gridcraft_engine=debug`).
  - `WGPU_BACKEND` (`dx12`/`vulkan`/`gl`) helps if the app will not open: default is DirectX 12; `vulkan` can crash on some Intel UHD drivers, `gl` on some AMD.
  - `WGPU_DX12_COMPILER=dxc` is for troubleshooting only.
- **Language:** the UI has no Italian (en, ja, zh, ko, pt-BR, ru, de, es...). A system set to another language falls back to English.
- **Working folder:** keep **copies** of the files to process in a dedicated folder such as `Documents/GridCraft` (create it). GridCraft has no safety "root": it reads and writes any path you give it.

## Connecting

| Mode | How to start | Transport | Notes |
|---|---|---|---|
| **MCP headless** (default) | `gridcraft-cli mcp [--in book.xlsx \| --sample budget]` | line-delimited JSON-RPC 2.0 on stdin/stdout; log on stderr | In-process engine session, no window. Starts with an empty `Book1`, or with `--in`/`--sample`. `screenshot` and `ui.*` methods return an error |
| **MCP attached to the live app** | app with `gridcraft --control 7985`, then `gridcraft-cli mcp --connect 7985` | MCP on stdio to TCP JSON lines on `127.0.0.1:7985` | Edits show in the window and `screenshot` works. 30 s I/O timeout; if the app restarted, retry once with a new connection |
| **CLI** | `gridcraft-cli <subcommand>` | single process | `info`, `convert`, `eval`, `cat`, `run`, `commands`, `functions`, `check`, `send`, `mcp`, `version` |
| **Control channel** | `gridcraft --control 7985` (or `GRIDCRAFT_CONTROL_PORT`) | line JSON over TCP `127.0.0.1` | Commands + real mouse/keyboard input, dialogs, screenshots. From the CLI: `gridcraft-cli send 7985 <method> '<json>'` |

- **Port:** upstream docs use 7979, but the kit uses **7985** so it does not collide with VectorCraft (7979).
- MCP protocol `2025-06-18`; it also answers `2025-03-26` and `2024-11-05` clients in their own version.
- **Authentication: none.** No token or token file, and no auth in docs or source. Anyone on the machine can drive the app on the open port (it listens on `127.0.0.1` only). Close the app (or restart it without `--control`) when done.
- **Read/write roots: none.** There is no `--root`; paths are plain filesystem paths, so be careful where you save.

### Choose the channel
1. **Do you have the `gridcraft` MCP tools?** Use them (live: the user sees the result and `screenshot` works).
2. **Do you have a shell on the user's machine?** Use the CLI (`run`, `cat`, `eval`, `convert`) or `send 7985 ...` to the open app.
3. **A cloud/Linux sandbox cannot run the user's `.exe`/app or reach their `127.0.0.1:7985`.** There you can only **prepare files**: `.jsonl` scripts for `run --script`, CSVs to import, data.
4. **Last resort: computer use on the GUI.** Non-US keyboards: `[` and `]` are typed with AltGr, which on Windows arrives as Ctrl+Alt; GridCraft treats it as text, so shortcuts with `[`/`]` (e.g. Ctrl+[) **do not fire**: use commands (`formulas.tracePrecedents`...). Formulas: localized Excel uses `;` as separator, GridCraft **does not**. In English (and so with any non-Spanish system) it wants commas between arguments and a dot for decimals: write `=IF(A1>0,1.5,0)`, not a localized `=SE(...)`. Only with the **Espanol** UI language does the editor accept Spanish names, `;` and decimal comma. Via API/CLI/MCP the syntax is **always canonical English**.

## Basic rules
- **Addresses:** A1 style. `"A1:D10"`, whole columns `"B:B"`, with sheet `"Sheet2!A1:C3"`, multiple areas `"A1:B2,D4"` (in `select`). R1C1 exists in the parser. Limits: 1,048,576 rows x 16,384 columns.
- **Selection:** most commands act on the current selection if you omit `range`. Always pass an explicit `range`, or `select`/`selection.set` first.
- **Sheets:** `inspect_workbook` returns names, used range, tables, charts, merges, filters and freeze. Commands `home.insertSheet`, `home.deleteSheet`, `sheet.rename`, `sheet.move`, `sheet.hide`/`sheet.unhide`, `sheet.tabColor` (params: `list_commands {"search":"sheet"}`). The `Sheet2!` prefix in `select` switches sheet.
- **Input:** `set_cell`/`cell.set` enter "as if typed". `=` means formula. Strings are interpreted (`"12%"`, `"1/2/2024"`, `"$5"`). Number formatting is **en-US only**, so for numbers use JSON values (`1200.5`), not strings with a decimal comma. In `write_range`, `null` clears the cell.
- **Formulas:** 500+ functions, dynamic arrays and spill (`A1#`, `#SPILL!`), LET/LAMBDA/MAP/REDUCE/SCAN/BYROW/BYCOL/MAKEARRAY, structured table references, names, 3-D references. List with `list_functions {"search":"lookup"}`. Test a formula without touching the sheet using `evaluate_formula`.
- **Recalculation:** automatic, dependency graph, single-thread. Volatile functions (`RAND`, `NOW`, `TODAY`...) are always dirty. Commands `formulas.calculationOptions`, `formulas.calculateNow`, `formulas.calculateSheet`.
- **Number formats:** `format_range.numberFormat` takes a name (`General`, `Number`, `Currency`, `Accounting`, `Comma`, `Percentage`, `Short Date`, `Long Date`, `Time`, `Fraction`, `Scientific`, `Text`) or a code (`"#,##0.00"`, `"0%"`, `"yyyy-mm-dd"`). Colors `"#RRGGBB"`, plus `fontColor:"auto"` and `fillColor:"none"`.
- **Undo:** copy-on-write, nearly free snapshots. `write_range` is one step. Tools `undo`/`redo` with `steps`, commands `edit.undo`/`edit.redo`. `format_range` runs one command per property, so it probably counts several undo steps (undocumented; check with `inspect_workbook`).
- **Parameters:** MCP tools are **strict**. An unknown argument, wrong type or out-of-`enum` value returns JSON-RPC error `-32602`. For `execute_command` `params` the strictness on extras is undocumented: copy names from `list_commands` (`params` field).
- **Batch / `$N`:** there are no batch tools and no `$N` references to earlier results. For multiple steps use `gridcraft-cli run --script` (stops at first error) or sequential calls. `engine.journal` returns executed commands in replayable form (passwords excluded).
- **Traps (alpha):**
  - `.xls`, `.xlsb` and `.ods` are **import-only**; in `.xlsb`/`.ods` formulas become values. Save as `.xlsx`.
  - Several chart types do not survive XLSX save: stock, combo, histogram, box, waterfall, funnel, treemap, sunburst.
  - Slicers, timelines and the data model are lost on save.
  - PDF uses only the 14 standard fonts, so non-Latin scripts are not exported.
  - Open bugs: 15-digit equality (#57) and spills not updating after structural edits (#133/#139).
  - Very large files (~100 MB) do not open.
  - CSV saves **displayed** values (with applied format). A file opened as `;` CSV stays `;`.

## MCP tools (21)

| Tool | Arguments | Notes |
|---|---|---|
| `execute_command` | `command`, `params?` | Runs **any** command (`cell.set`, `home.bold`, `data.removeDuplicates`...) |
| `list_commands` | `search?` | id, label, ribbon path, shortcut, `params` doc, `enabled`. Read-only |
| `inspect_workbook` | none | = `document.inspect`: title, path, dirty, sheets, names, selection, calc mode, undo history |
| `read_range` | `range?`, `sheet?`, `formulas?`, `formatted?` | 2-D rows. Empty cells as `null`, errors as `{"error":"#DIV/0!"}`. Without `range` reads the used range |
| `write_range` | `range`, `values` (2-D) | = `range.setValues`, starts at the top-left cell. One undo step |
| `set_cell` / `get_cell` | `cell`, `input` / `cell` | First is `cell.set`. Second returns value, type, formula, display text, raw input, spill, style, merge, comment, link |
| `evaluate_formula` | `formula`, `cell?` | Computes without modifying. `cell` is for relative references and `ROW()` |
| `select` | `range`, `active?` | = `selection.set` |
| `format_range` | `range` + `bold?` `italic?` `underline?` `fontName?` `fontSize?` (1-409) `fontColor?` `fillColor?` `numberFormat?` `horizontalAlign?` (`left`/`center`/`right`/`justify`) `wrapText?` `borders?` | Composes the `home.*` commands. `borders`: a preset or `{preset, style?, color?}` |
| `insert_chart` | `range`, `type`, `title?`, `subtype?`, `at?` | = `insert.chart`, returns the chart id |
| `create_table` | `range`, `style?`, `header?`, `name?` | = `insert.table`. Default style `TableStyleMedium2` |
| `sort_range` | `keys`, `range?`, `header?` | = `data.sort`. `keys`: `"B"`, `"B desc"` or `{column, order:"asc"\|"desc"}` |
| `filter` | `column`, `range?`, `values?`, `custom?`, `top?`, `blanks?`, `clear?` | = `data.filterBy`, turns on AutoFilter if needed |
| `open_workbook` | `path` | Opens `.xlsx .xlsm .xls .xlsb .ods .csv .tsv .txt .json`. Check the `warnings` |
| `save_workbook` | `path?` | Format from extension: `.xlsx` (default), `.csv`/`.tsv` (active sheet only), `.json`, `.html`. Without `path` saves to the current file, but errors if the file was never saved |
| `new_workbook` | `sample?` (`budget`/`sales`/`grades`) | New empty or sample workbook |
| `undo` / `redo` | `steps?` | History |
| `screenshot` | `path` | PNG of the window, only with `--connect` |
| `list_functions` | `search?`, `category?` | E.g. category `"Math & Trig"`, `"Lookup & Reference"` |

- Chart `type` values: `column`, `bar`, `line`, `pie`, `doughnut`, `area`, `scatter`, `bubble`, `radar`, `combo`, `histogram`, `waterfall`, `funnel`, `treemap`, `sunburst`, `boxWhisker`, `stock`. `subtype`: `clustered`, `stacked`, `stacked100`, `markers`.
- Border presets: `bottom`, `top`, `left`, `right`, `none`, `all`, `outside`, `thickOutside`, `thickBottom`, `doubleBottom`, `topBottom`, `topThickBottom`, `topDoubleBottom`, `insideHorizontal`, `insideVertical`, `inside`, `diagonalDown`, `diagonalUp`. Styles: `thin|medium|thick|dashed|dotted|double|hair...`.
- `custom` filter: `{op, value, op2?, value2?, and?}`, `op` among `= <> > >= < <= beginsWith contains...`. `top` filter: `{count, bottom?, percent?}`.
- **Resources:** `gridcraft://workbook` (document.inspect), `gridcraft://commands`, `gridcraft://functions`. `prompts/list` is empty.
- **Errors:** every result is a text block with indented JSON. Failures set `isError: true`. JSON-RPC codes: `-32602` arguments, `-32700` malformed JSON, `-32601` unknown method.

## Recipes (MCP; via channel/CLI they are the same commands)

```text
# Orient on a COPY (first copy the original into your working folder)
open_workbook {"path":"/path/to/work/copy.xlsx"}
inspect_workbook {}
read_range {"range":"A1:F20"}                 # values; then {"formulas":true} or {"formatted":true}

# Write data and formulas (example from docs/mcp.md)
write_range {"range":"A1","values":[["Item","Cost"],["Rent",1200],["Food",400],["Total","=SUM(B2:B3)"]]}
read_range {"range":"B4"}                     # -> [[1600.0]]
set_cell {"cell":"D2","input":"=B2*C2"}
execute_command {"command":"edit.fillDown","params":{"range":"D2:D20"}}
evaluate_formula {"formula":"=XLOOKUP(\"b\",{\"a\",\"b\"},{1,2})"}   # -> 2, no change

# Format
format_range {"range":"A1:B1","bold":true,"fillColor":"#DDEBF7","borders":"bottom"}
format_range {"range":"B2:B4","numberFormat":"#,##0.00","horizontalAlign":"right"}
execute_command {"command":"home.bold","params":{"range":"A1:C1"}}

# Table, sort, filter
create_table {"range":"A1:D20","name":"Sales"}
sort_range {"range":"A1:D20","keys":[{"column":"C","order":"desc"},"A"],"header":true}
filter {"column":"B","values":["North","South"]}
filter {"column":"C","custom":{"op":">","value":"10"}}
execute_command {"command":"data.removeDuplicates","params":{"range":"A1:C20"}}

# Chart
insert_chart {"range":"A1:C6","type":"column","subtype":"clustered","title":"Sales","at":"F2"}

# PivotTable, conditional formatting, sparkline: params are undocumented, discover them
list_commands {"search":"pivot"}              # insert.pivotTable, pivot.addField, pivot.moveField...
list_commands {"search":"conditional"}        # home.conditionalFormat
list_commands {"search":"sparkline"}          # insert.sparkline

# Picture inside a cell (docs/cell-pictures.md)
execute_command {"command":"insert.picture","params":{"path":"photo.png","placement":"cell","at":"B2","alt":"Product photo"}}

# Password-protected file: without the password it errors (never a dialog); saving does not rewrite the encrypted original
execute_command {"command":"file.open","params":{"path":"Secret.xlsx","password":"..."}}

# Export (only with the user's OK, to a NEW file)
save_workbook {"path":"/path/to/work/result.xlsx"}
save_workbook {"path":"/path/to/work/data.csv"}      # active sheet only
list_commands {"search":"pdf"}                # file.exportPdf: check params before use

# See the result (live app only)
screenshot {"path":"/path/to/work/shot.png"}
```

Native printing (desktop app only): `ui.printSettings {}` shows printers and state. With `{"refresh":true}` it starts discovery: poll until `discovering` is false. Set `printer`, `copies` (1-999), `duplex` (`printerDefault|off|longEdge|shortEdge`) and `grayscale`. Then `ui.printSubmit` prints **without a dialog**: always ask first.

## CLI

```text
gridcraft-cli info <file> [--json]
gridcraft-cli convert <in> <out> [--sheet NAME]          # out: .xlsx .csv .tsv .json .html
gridcraft-cli eval <formula> [--in FILE] [--sheet S] [--cell A1] [--json]
gridcraft-cli cat <file> [--range R] [--sheet S] [--formulas] [--csv]
gridcraft-cli check <file-or-dir>... [--json] [--tolerance 1e-9] [--max-report N]
gridcraft-cli run [--in FILE | --sample budget|sales|grades] [--cmd 'id={json}']... [--script steps.jsonl]
                  [--out FILE] [--print RANGE|used] [--sheet S] [--csv] [--formulas] [--quiet]
gridcraft-cli commands [--json] [--search X]
gridcraft-cli functions [--json] [--search X] [--category C]
gridcraft-cli mcp [--connect PORT] [--in FILE | --sample NAME]
gridcraft-cli send <port> <method> [json]
gridcraft-cli version
```

- Errors go to stderr with a non-zero exit code. Import warnings go to stderr, stdout stays clean.
- `run` opens the file (or a sample, or an empty workbook), executes `--cmd` and `--script` **in the order given**, saves `--out` and prints every `--print`. Each non-null result is a JSON line (unless `--quiet`). It stops at the first error.
- `--cmd` forms: `id`, `id={json}` or `id {json}`. A script has one command per line in those forms, or `{"command":"cell.set","params":{...}}`. Blank lines and `#` are ignored.
- `check`: calculation oracle (compares with values saved by Excel) and XLSX round-trip fidelity. Writes nothing. Exit 0 = ok, 1 = differences, 2 = usage error. `--json` uses schema `gridcraft-check/1`.
- PowerShell: inner quotes of a JSON passed to an `.exe` can get lost (especially in Windows PowerShell 5.1). For multiple steps prefer `--script` with a `.jsonl` file.

```powershell
$gc = "C:\Program Files\GridCraft\gridcraft-cli.exe"
& $gc info .\copy.xlsx
& $gc cat  .\copy.xlsx --range A1:F20
& $gc eval '=PMT(5%/12,360,-300000)'
& $gc convert copy.xlsx data.csv --sheet Sales
& $gc run --in copy.xlsx --script steps.jsonl --out copy-edited.xlsx --print A1:D5
& $gc commands --search border
& $gc send 7985 engine.execute '{"command":"home.bold","params":{"range":"A1:C1"}}'
```

Example `steps.jsonl`:
```text
# header and total
home.bold={"range":"A1:C1"}
{"command":"cell.set","params":{"cell":"D2","input":"=B2*C2"}}
edit.fillDown={"range":"D2:D20"}
```

## Control channel (desktop app)

- **Start:** `gridcraft --control 7985` (the "GridCraft (Claude)" launcher) or `GRIDCRAFT_CONTROL_PORT=7985`.
- **Transport:** TCP on `127.0.0.1:<port>`, line JSON. Each request gets **one** response line. **No authentication.**
- **Format:**
  ```json
  {"id": 1, "method": "engine.execute", "params": {"command": "cell.set", "params": {"cell": "B2", "input": "=SUM(A1:A9)"}}}
  {"id": 1, "ok": true, "result": {"cell": "B2", "value": "45"}}
  ```
  Errors: `{"id": ..., "ok": false, "error": "..."}`.
- **Limits:** the MCP bridge has a 30 s timeout per request and accepts responses up to 256 MiB. `ui.screenshot` waits for a presented frame; with the window not visible it fails: `ui.focus` and retry.

| Method | Params | What it does |
|---|---|---|
| `engine.execute` | `{command, params}` | Runs any command |
| `engine.commands` | `{search?}` | All commands with id, label, path, shortcut, params doc, enabled |
| `engine.journal` | none | Executed commands, replayable (passwords excluded) |
| `document.inspect` | none | Workbook summary |
| `sheet.read` | `{range?, sheet?, formulas?, formatted?}` | Values, formulas or display text |
| `cell.get` | `{cell?}` | One cell: value, formula, text, style, merge, comment, link |
| `ui.inspect` | none | Ribbon tab, editor, open dialog, Backstage page, message, grid rect, scroll, perf, open/dirty documents |
| `ui.click` / `ui.drag` / `ui.move` | `{x, y, toX?, toY?, button?, count?, shift?, cmd?, alt?}` | Real pointer, in window points |
| `ui.cell.click` / `ui.cell.drag` | `{cell}` / `{from, to}` | Pointer at a cell's center (also works in point mode while editing) |
| `ui.key` / `ui.text` | `{key, shift?, cmd?, alt?, ctrl?}` / `{text}` | Keyboard. On Windows `cmd` means Ctrl. Keys: `enter`, `esc`, `delete`, `backspace`, arrows, `space`, `tab`, `pageup`/`pagedown`, `home`, `end`, `f2`, `f4`, `f9`... |
| `ui.edit.begin` / `.text` / `.commit` / `.cancel` | `{text?}` / `{move?: down\|right\|up\|left\|none}` | Cell editor |
| `ui.ribbon` | `{tab}` | Switch ribbon tab |
| `ui.dialog` / `.set` / `.confirm` / `.cancel` | `{name}` / `{field, value}` | Dialogs |
| `ui.backstage` (via `engine.execute`) | `{page?: home\|new\|open\|info\|saveAs\|print\|export\|options\|close}` | File view |
| `ui.screenshot` | `{path?}` | PNG (base64 if no `path`), animations settled |
| `ui.set` | `{dark?, formulaBar?, ribbonCollapsed?, animations?}` | UI preferences |
| `ui.resize` / `ui.focus` | `{width, height}` | Window |
| `app.open` / `app.save` / `app.quit` | `{path, password?}` / `{path?}` | Files and quit |

- Any other method name that is a command id (e.g. `home.bold`) runs that command.
- In source (`control.rs`) but not in docs: aliases `command`/`ui.menu.invoke` for `engine.execute`, plus `ui.message.dismiss` and `selection.stats`.
- Theme: `engine.execute {command:"view.theme", params:{mode:"system"|"light"|"dark"}}`. Animations: `view.animations {on:false}`. UI language: `app.language.set {language:"ja"}` (alias `code`), or `app.language.english`/`app.language.japanese`.
- **Element ids:** there are no widget ids. Target by cell (`ui.cell.*`), by coordinates (the `grid` rect from `ui.inspect`), by tab/dialog/field name. Charts have an id returned by `insert_chart`; `selectedChart` appears in `ui.inspect`.

## Commands by family (real ids from `docs/parity-checklist.md`)

- **Primitives:** `cell.set`, `cell.get`, `range.setValues`, `selection.set`, `sheet.read`, `document.inspect`, `edit.undo`, `edit.redo`, `data.filterBy`, `formulas.functions`, `sheet.insert`.
- **file.:** `new`, `open`, `save`, `saveAs`, `close`, `print`, `exportPdf`, `exportCsv`, `exportHtml`, `properties`, `options`.
- **edit.:** `cut`, `copy`, `paste`, `pasteSpecial`, `formatPainter`, `fillDown/Up/Left/Right`, `fillSeries`, `flashFill`, `find`, `replace`, `goTo`, `goToSpecial`, `clearAll/Contents/Formats/Comments/Hyperlinks`.
- **home.** (94):
  - Font: `fontName`, `fontSize`, `increaseFontSize`/`decreaseFontSize`, `bold`, `italic`, `underline`, `doubleUnderline`, `strikethrough`, `subscript`/`superscript`, `fontColor`, `fillColor`, `borders`, `drawBorder`.
  - Alignment: `alignLeft/Center/Right/Top/Middle/Bottom`, `justify`, `orientation`, `wrapText`, `increaseIndent`/`decreaseIndent`, `mergeCenter`/`mergeAcross`/`mergeCells`/`unmergeCells`.
  - Numbers: `numberFormat`, `accounting`, `percent`, `comma`, `increaseDecimal`/`decreaseDecimal`.
  - Styles: `conditionalFormat`, `clearRules`, `manageRules`, `formatAsTable`, `cellStyle`, `newCellStyle`.
  - Cells and sheets: `insertCells/Rows/Columns/Sheet`, `deleteCells/Rows/Columns/Sheet`, `rowHeight`, `columnWidth`, `autofitRowHeight`, `autofitColumnWidth`, `defaultWidth`, `hideRows/Columns`, `unhideRows/Columns`, `formatCells`, `lockCell`, `addIns`.
- **insert.:** `table`, `pivotTable`, `recommendedPivotTables`, `chart`, `recommendedCharts`, `pivotChart`, `sparkline`, `picture`, `shape`, `icons`, `textBox`, `symbol`, `link`, `checkbox`, `slicer`. **Missing:** `smartArt`, `screenshot`, `map`, `timeline`, `wordArt`, `signatureLine`, `object`, `equation`.
- **chart.:** `set`, `changeType`, `selectData`, `switchRowColumn`, `addElement`, `quickLayout`, `changeColors`, `move`, `formatSelection`, `shapeFill`, `shapeOutline`.
- **table.:** `rename`, `resize`, `style`, `headerRow`, `totalRow`, `bandedRows`/`bandedColumns`, `firstColumn`/`lastColumn`, `removeDuplicates`, `convertToRange`, `summarizePivot`, `insertSlicer`.
- **pivot.:** `addField`, `removeField`, `moveField`, `valueSettings`, `calculatedField`, `filter`, `sort`, `group`, `collapse`, `refresh`, `changeSource`, `reportLayout`, `subtotals`, `grandTotals`, `style`, `fieldList`, `delete`.
- **formulas.:**
  - Function library: `autoSum`, `insertFunction`, `financial`, `logical`, `text`, `dateTime`, `lookupReference`, `mathTrig`, `moreFunctions`, `recentlyUsed`.
  - Names: `defineName`, `nameManager`, `createFromSelection`, `useInFormula`.
  - Auditing: `tracePrecedents`, `traceDependents`, `removeArrows`, `showFormulas`, `errorChecking`, `evaluateFormula`, `watchWindow`.
  - Calculation: `calculationOptions`, `calculateNow`, `calculateSheet`.
- **data.:**
  - Sort and filter: `sort`, `sortAscending`, `sortDescending`, `filter`, `clearFilter`, `reapply`, `advancedFilter`.
  - Data tools: `removeDuplicates`, `textToColumns`, `flashFill`, `validation`, `circleInvalid`, `consolidate`.
  - What-if analysis: `goalSeek`, `scenarioManager`, `dataTable`, `solver`, `forecastSheet`, `analysisToolPak`.
  - Outline: `group`, `ungroup`, `subtotal`, `showDetail`, `hideDetail`, `showLevels`.
  - External data: `getData`, `fromTextCsv`, `fromTableRange`, `refreshAll`, `queriesConnections`.
  - **Missing:** `fromWeb`, `stocks`, `geography`, `relationships`, `analyzeData`.
- **review.:** `spelling`, `thesaurus`, `workbookStatistics`, `checkAccessibility`, `newComment`, `deleteComment`, `previousComment`/`nextComment`, `showComments`, `newNote`, `showNote`, `showAllNotes`, `convertToComments`, `protectSheet`, `unprotectSheet`, `protectWorkbook`, `allowEditRanges`. Missing: `smartLookup`, `translate`, `showChanges`.
- **pageLayout.:** `theme`, `themeColors`, `themeFonts`, `themeEffects`, `margins`, `orientation`, `size`, `printArea`, `breaks`, `background`, `printTitles`, `scaleToFit`, `printGridlines`, `printHeadings`, `pageSetup`, `headerFooter`.
- **view.:**
  - Views: `normal`, `pageBreakPreview`, `pageLayout`, `customViews`.
  - Show: `ruler`, `gridlines`, `formulaBar`, `headings`.
  - Zoom: `zoom`, `zoom100`, `zoomToSelection`.
  - Windows and panes: `freezePanes`, `freezeTopRow`, `freezeFirstColumn`, `unfreezePanes`, `split`, `newWindow`, `arrangeAll`, `hideWindow`, `unhideWindow`, `switchWindows`.
  - Other: `macros`, `theme`, `darkMode`, `animations`, `rightToLeft`.
- **sheet.:** `rename`, `move`, `hide`, `unhide`, `tabColor`. **arrange.:** `align`, `bringForward`, `sendBackward`, `group`, `rotate`, `selectionPane`.
- **draw.:** `pen`, `eraser`, `lasso`, `inkToShape`. **automate.:** `newScript`, `allScripts`, `recordActions`.
- **app./ui.** (desktop app only): `app.language.set/english/japanese`, `ui.backstage`, `ui.printSettings`, `ui.printSubmit`.

This list is neither complete nor does it show parameters. Before using a command call `list_commands {"search":"..."}` (or `gridcraft-cli commands --search ...`) and read its `params` and `enabled` fields.

## Recommended workflow

1. **Orient:** `inspect_workbook` (or `gridcraft-cli info`), then `read_range` on the useful areas. With the live app add `ui.inspect` and `screenshot`.
2. **Work on a copy:** copy the original into the working folder, then `open_workbook` on the copy. For experiments from scratch use `new_workbook {"sample":"sales"}`.
3. **Find the command:** `list_commands {"search":"..."}` for id, `params` and `enabled`. Search functions with `list_functions`; try formulas first with `evaluate_formula`.
4. **Edit in small steps:** one range or command at a time, with `range` always explicit. On error read the message, fix, and `undo` if needed.
5. **Verify:** re-read with `read_range` (values, then `formulas:true`, then `formatted:true`) or `get_cell`. Look for `#SPILL!`, `#REF!` and `#VALUE!`. On the live app use `screenshot`, headless use `gridcraft-cli cat`. If the file is going to Excel users, run `gridcraft-cli check` on the result.
6. **Save or export only with the user's OK**, always to a **new** file (`save_workbook {"path":...}` / `convert`). Never overwrite the original. Remember the limits: import-only for xls/xlsb/ods, loss of some charts and slicers, PDF without non-Latin fonts.
7. **Close:** the control channel has no authentication. When finished quit the app (`app.quit`, first asking about unsaved changes: `documents[].dirty` in `ui.inspect`) or relaunch it from the normal shortcut, without `--control`.
