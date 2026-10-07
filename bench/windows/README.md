# Excel for Windows as a second oracle

The harness in `bench/coverage` takes its answers from Excel for Mac. This one sends the same books to Excel for Windows and stores its answers beside them, in `test/excel/win/`, so the engine can be checked against both and every place where the two Excels part can be seen case by case.

The books are the Mac harness's own: the case sets, the probes, the measurement, the display, spill, workbook and structure sets are written by the same code (`bench/coverage/book.ts`, `case-book.ts`, `display.ts`, `struct.ts`), byte for byte. Only the way Excel is driven differs: COM through PowerShell over SSH instead of AppleScript.

## Where the answers are

`test/excel/win/` has the Mac's layout:

- `coverage/excel-en.json`, `coverage/excel-ru.json`: the measurement (its cases are in `test/excel/coverage/cases.json`).
- `probes.json`, `cases/<set>.json`, `cases/ru-RU/<set>.json`, `legacy/<set>.json`: the probes and the case sets.
- `display/<set>.json`, `display/ru-RU/<set>.json`, `spill/<set>.json`, `workbook/<set>.json`, `struct/<set>.json`.

Each file's `meta` names Excel's version and the settings it computed with: the interface language, the regional format with its decimal, thousands and list separators, the system locale with its code page, and the screen's dpi. A formula Excel for Windows does not take when it opens a book is listed in `rejected` of its set.

`test/excel-win.test.ts` checks the engine against the formula answers (`test/excel-win-cases.ts` has the cases' books and `WIN_KNOWN`, the engine's answers pinned where it knowingly differs from Excel for Windows, each with its reason). `npx tsx bench/windows/compare.ts` prints, for the stored answers, the cases where Windows and the Mac part, with the engine's answer and which of the two it follows, and the cases where both Excels agree and the engine does not.

## Running it

```sh
npx tsx bench/windows/run.ts env            # the machine's settings, and Excel's as it starts with them
npx tsx bench/windows/run.ts all ru         # everything the Mac has answers for in ru-RU
npx tsx bench/windows/run.ts all en         # the same in en-US: the formulas, the display, spill, workbook and struct sets
npx tsx bench/windows/run.ts set char-code ru
npx tsx bench/windows/run.ts display general
npx tsx bench/windows/run.ts struct rows
```

A command takes the machine's lock, puts the machine in the locale it needs, and keeps the stored answers of the cases that did not change (`FCOV_FRESH=1` runs everything again). `XLWIN_HOST` names another machine than `user@windows-host`; `XLWIN_KEEP=1` keeps a job's local files in `bench/windows/out/`.

Changing the system locale needs a restart, which the harness does itself (it waits up to ten minutes for SSH). Run all of one locale before the other: `all ru`, then `all en`.

## The machine

`windows-host` is a virtual machine with Windows 11 Pro (Russian), Office ProPlus 2024 (Excel 16.0.17928, dynamic arrays) with the English language pack, and OpenSSH with key login; its SSH shell is PowerShell 5.1. The harness keeps its files in `C:\xlcalc\`: `bin\` the scripts (copied from here at the start of a command), `work\` the jobs, `lock\` the lock.

What a run depends on, and how the harness sets it (`host.ps1 -do locale`):

- Excel's interface language follows the Windows display language of the logon, and the interface language changes computation (with the Russian interface `"x"&TRUE` is `xИСТИНА`). Office's own settings (`UILanguageTag`, the `UILanguage` policy) do not change it, and `Set-WinUILanguageOverride` alone stays pending until an interactive logon. The harness writes `PreferredUILanguages` under `HKCU\Control Panel\Desktop`, which the next SSH logon reads.
- The regional format: `Set-Culture` (en-US: `.` and `,`; ru-RU: `,` and a space, list separator `;`).
- `CHAR` and `CODE` follow the system code page: `Set-WinSystemLocale` and a restart (en-US 1252, ru-RU 1251).

`host.ps1 -do env` reports all of these and starts Excel to read its interface language (`LanguageSettings.LanguageID(2)`: 1033 or 1049); a run refuses to start when any of them is not the locale's.

## How a job runs

`remote.ts` writes the books here and a `job.json` listing each book's steps, copies the directory to `C:\xlcalc\work\` with `scp`, runs `xl.ps1` on it and copies back `out\`. `xl.ps1` starts one Excel, opens the books one by one, takes their steps (a full calculation, edits, reads of ranges, shown text, a structure operation) and writes each book's answer as soon as it is done; at the end it checks that no book is left open and quits Excel.

- A range is read in one call: `Value2` gives numbers, text, booleans and errors as COM codes (`-2146826281` is `#DIV/0!`, 0x800A0000 plus Excel's error number), so the wrapper and mirror columns the Mac reads are not needed. They stay in the books, which are the Mac's.
- A formula case's book is checked to still hold its formulas (`HasFormula`): a book Excel repaired on opening loses them. A book Excel does not open, opens repaired or hangs on is halved down to the formula at fault, which is stored as not taken.
- A job holds up to 40 books. Each SSH call costs seconds on this machine (a new PowerShell starts), so the harness shares one SSH connection per run and puts as many books in a job as it can.
- A job that runs past its time (Excel waits on a window nobody sees) is stopped with Excel; the books done keep their answers and the rest run again.
- Excel started over SSH runs in session 0. It opens no file ("cannot access the file") unless the system profile has a desktop folder, `C:\Windows\System32\config\systemprofile\Desktop` and its `SysWOW64` twin; `host.ps1 -do clean` makes sure of them, stops Excel processes left in session 0 and removes the jobs' files. An interactive user's Excel is left alone.

## What is not run here

- The xlsx step (Excel saving the books it computed) and the ru-RU formula text snapshot (`bench/locale`): both work through COM (`SaveAs`, `FormulaLocal`) but are not ported.
- Mac-only conditions do not carry over: the spill order runs (`order-runs.json`) record Excel for Mac's varying choice and are no set.

## The first probe

`probe.ps1` is the probe that first showed COM working over SSH, kept as it ran: its Office settings do not change the interface language.

## How it was checked

The first set run through the harness, `char-code` in ru-RU, gave 305 answers; 298 equal the Mac's, and the 7 others are `CODE` of characters outside the code page, 63 on Windows and 95 on the Mac. The engine matches all 305.
