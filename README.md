# SPI Open Access Retro

A public retro reimplementation of a classic integrated office suite, rebuilt as a DOS-style text-mode application in modern .NET.

This project is aimed at:

- retro software preservation
- historical UI recreation
- office-suite experimentation
- database/forms/reporting workflows
- scripting-driven business applications

It is not affiliated with the original vendor. Historic product names are used only for reference, compatibility goals, and preservation context.

## Why This Project Stands Out

Most retro-inspired projects stop at visuals. This one is building an actual modular office environment with a shared shell, persistent workspace state, and feature-specific modules that behave like a real business application.

Current highlights:

- DOS-style full-screen shell with module tabs, status bars, and command workflow
- persistent session state across app restarts
- editable database tables with search, append, and update commands
- interactive spreadsheet state with recalculation and scenario-style commands
- editable word-processing draft state with preview
- editable internal mail draft workflow
- reporting, communications, and scripting modules with dedicated operational views
- automated tests covering module registration, command screens, persistence, and data workflows

## Modules

- `Data`: desktop database, forms, reports, search, append, update
- `Calc`: spreadsheet calculations, goal-seek-style scenarios, print area preview
- `Word`: document draft, preview, merge-oriented workflow
- `Mail`: inbox preview, message open, draft composition, routing rules
- `Comm`: dial, file send, capture session views
- `Rpt`: run, schedule, and design reporting workflows
- `Prog`: PRO-inspired scripting runtime, variable watch, compile/run views

## Screenshots

The application is designed to look and feel like a late DOS business suite inside a terminal window.

Recommended next step for GitHub:

1. capture a few terminal screenshots
2. add them to a `docs/` or `.github/assets/` folder
3. embed them here near the top of the README

That will materially improve click-through and credibility on GitHub.

## Quick Start

Run the app:

```powershell
dotnet run --project .\src\SpiOpenAccess.App
```

Open a module directly:

```powershell
dotnet run --project .\src\SpiOpenAccess.App -- db
```

Run the test suite:

```powershell
dotnet test .\SpiOpenAccess.sln
```

## Example Commands

Inside the app:

```text
menu
use sheet
set Q1 190000
recalc
use word
new letter
title Sales Letter
type Please review the attached quote.
use mail
compose
to FINANCE
subject Cash status
body Need current cash figures.
use db
find CUSTOMERS Bremen
append CUSTOMERS Id=C-1004;Company=Retro Works;City=Berlin;Tier=B
update CUSTOMERS C-1004 City Leipzig
```

## Architecture

- `src/SpiOpenAccess.App`: shell, session handling, persistence, command routing
- `src/SpiOpenAccess.Core`: shared abstractions and workspace state models
- `src/SpiOpenAccess.Infrastructure`: suite composition and default workspace setup
- `src/SpiOpenAccess.Modules.*`: functional office modules
- `tests/SpiOpenAccess.Tests`: automated tests for core behaviors and command flows

## Current State

This repository is already more than a visual prototype, but it is not yet a full historical clone.

What exists today:

- a coherent retro shell
- persistent workspace state
- cross-module command handling
- editable state in multiple modules
- deterministic seeded data
- green automated tests

What still needs to grow:

- true cursor-driven editing with `Console.ReadKey`
- stricter 80x25 layout behavior
- popup dialogs and menu trees
- file-based document/workbook/mailbox management
- deeper report engine logic
- a larger PRO runtime with parser/AST/VM
- richer database CRUD and navigation

## Open Source Position

This repository is intended to be publishable on GitHub as an English-language retro software project.

Please do not contribute:

- original proprietary source code
- original proprietary binaries
- copied historical help files or artwork
- trademark-sensitive assets presented as if they were official

## Contributing

Contributions are welcome, especially around:

- retro UI fidelity
- terminal UX
- module behavior
- persistence
- testing
- historical workflow research

See [CONTRIBUTING.md](C:/tmp/spiopenaccess/CONTRIBUTING.md).

## License

Released under the MIT License. See [LICENSE](C:/tmp/spiopenaccess/LICENSE).
