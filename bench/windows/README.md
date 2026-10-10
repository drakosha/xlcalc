# Excel for Windows as a second oracle

The harness in `bench/coverage` takes its answers from Excel for Mac. This one sends the same books to Excel for Windows and stores its answers beside them, in `test/excel/win/`, so the engine is checked against both and every place where the two Excels part shows case by case.

The books are the Mac harness's own: the measurement, the probes, the case sets, the legacy sets, and the display, tables, spill, workbook and structure sets are written by the same code (`bench/coverage/book.ts`, `case-book.ts`, `display.ts`, `tables.ts`, `struct.ts`, `xlsx/filters.ts`, `xlsx/formats.ts`), byte for byte. Only the way Excel is driven differs: COM through PowerShell over SSH instead of AppleScript.

## Where the answers are

`test/excel/win/` has the Mac's layout, en-US's answers at the top of each directory and another locale's in a directory named by its tag (`bench/coverage/locales.ts`):

- `coverage/excel-en.json`, `coverage/excel-ru.json`, `coverage/excel-<tag>.json` for another locale: the measurement (its cases are in `test/excel/coverage/cases.json`).
- `probes.json`, `cases/<set>.json`, `cases/ru-RU/<set>.json`, `legacy/<set>.json`: the probes, the case sets, the formulas of the old kind (each at the row it had in the Mac's run, as implicit intersection takes the row).
- `display/`, `display/ru-RU/`, `tables/`, `tables/ru-RU/`, `spill/`, `workbook/`, `struct/`: the book sets.
- `parts/<set>.json` and `parts/<set>/`: notes, hyperlinks and frozen panes as Excel read them in each case's book, and the files it saved, as the Mac's.
- `cond/<set>.json`, `cond/ru-RU/<set>.json` and their folders: conditional formats in the Mac step's books (`bench/coverage/cond/`), the rules the set `made` makes through COM (`FormatConditions.Add`, `AddTop10`, `AddAboveAverage`, `AddUniqueValues`, `AddColorScale`, `AddDatabar`, `AddIconSetCondition`), what the read cells show by their `DisplayFormat`, the rules each sheet lists, and the files Excel saved.
- `entry/<set>.json`, `entry/ru-RU/<set>.json`: text typed into cells through `Formula2Local` in the Mac step's books (`bench/coverage/entry/`), what each cell shows, its format codes and prefix character, and what the book Excel saved holds; `meta` names Excel's `AutoPercentEntry`.
- `xlsx/filters/` and `xlsx/filters.json`: the filter files as Excel computed and saved them; `xlsx/formats.json` and `xlsx/ru-RU/formats.json` the codes Excel shows for the built-in number formats 1 to 58 (`NumberFormat` and `NumberFormatLocal`, both in the user's language as COM reads them over SSH).

Each file's `meta` names Excel's version and the settings it computed with: the interface language, the regional format with its decimal, thousands and list separators, the system locale with its code page, and the screen's dpi. A formula Excel for Windows does not take when it opens a book would be listed in `rejected` of its set; none is.

## The tests and the differences

Each kind has its test beside the Mac's: `test/excel-win.test.ts` (formulas), `display-win`, `tables-win`, `spill-win`, `workbook-win`, `structure-win`, `legacy-win`, `xlsx-win`. Where the engine knowingly differs from Excel for Windows, the test pins it with the reason: `WIN_KNOWN` in `test/excel-win-known.ts` for formulas (reasons in `WIN_REASONS`, `test/excel-win-cases.ts`, or the Mac's own where both Excels agree), rules by reason for what cells show (`WIN_SHOWN`), and lists by case in the others. Where the two Excels' regional data part, the engine follows Excel for Windows (the order of text, month names, built-in formats by number, the times of ru-RU, `CHAR` and `CODE`); elsewhere it follows Excel for Mac.

`npx tsx bench/windows/compare.ts > bench/windows/differences.md` writes the table of the formula cases where Windows and the Mac part (the Mac's answer, Windows's, the engine's, which of the two the engine follows, and the class: a platform difference with its reason, or noise), and the cases where both Excels agree and the engine does not.

## Running it

A command that takes a locale takes `en`, `ru` or the tag of one of the package's locales (`de-DE`), en-US when it names none.

```sh
npx tsx bench/windows/run.ts env            # the machine's settings, and Excel's as it starts with them
npx tsx bench/windows/run.ts all ru         # everything the Mac has answers for in ru-RU
npx tsx bench/windows/run.ts all en         # the same in en-US
npx tsx bench/windows/run.ts functions en   # the formula cases only; books en: the book sets only
npx tsx bench/windows/run.ts set char-code ru
npx tsx bench/windows/run.ts display general
npx tsx bench/windows/run.ts tables ops
npx tsx bench/windows/run.ts struct rows
XLWIN_CODEPAGE=any npx tsx bench/windows/run.ts parts files
XLWIN_CODEPAGE=any npx tsx bench/windows/run.ts cond looks
npx tsx bench/windows/run.ts xlsx filters
npx tsx bench/windows/run.ts xlsx formats ru
npx tsx bench/windows/run.ts entry typed ru
```

A command takes the machine's lock, puts the machine in the locale it needs, and keeps the stored answers of the cases that did not change (`FCOV_FRESH=1` runs everything again). `XLWIN_HOST` names another machine than `user@windows-host`; `XLWIN_KEEP=1` keeps a job's local files in `bench/windows/out/`. `XLWIN_CODEPAGE=any` sets only the interface language and the regional format and leaves the system locale as it is, with no restart, for a set that reads no code page (no `CHAR`, `CODE`); the answers' `meta` names the code page they were computed with.

Changing the system locale needs a restart, which the harness does itself: it waits until the machine is up from a new boot, ten minutes at most. Run all of one locale before the other, `all en` and then `all ru`, which also leaves the machine in its own locale.

### Another locale

`npx tsx bench/windows/run.ts all de-DE` computes in de-DE what the Mac has answers for in de-DE, into the directories of its tag. The tag must be one of the package's: Excel's interface language id and the system code page it checks come from the locale's file (`regional.lcid`, `regional.codePage`, through `windowsLocale` in `bench/coverage/locales.ts`), and another tag stops before the machine is touched. The machine needs the language: Excel's interface language follows the Windows display language, so a new language needs its Windows language pack and Office's language pack for it, as the English one is installed here. Without them the run stops on its check of the settings (`windows-host is not in de-DE: language ..., Excel interface ...`). That has not been tried with a third language.

## The machine

`windows-host` is a virtual machine with Windows 11 Pro (Russian), Office ProPlus 2024 with the English language pack, and OpenSSH with key login; its SSH shell is PowerShell 5.1. Excel is 16.0.20430.20146 (Current Channel): Office updated itself from 16.0.17928 during the first run, so its automatic updates are off now (the policy `enableautomaticupdates` = 0 under `HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate`, and the scheduled tasks Office Automatic Updates 2.0 and Office Feature Updates disabled), and every answer is from that build. The harness keeps its files in `C:\xlcalc\`: `bin\` the scripts (copied from here at the start of a command), `work\` the jobs, `lock\` the lock.

What a run depends on, and how the harness sets it (`host.ps1 -do locale`):

- Excel's interface language follows the Windows display language of the logon, and the interface language changes computation (with the Russian interface `"x"&TRUE` is `xИСТИНА`). Office's own settings (`UILanguageTag`, the `UILanguage` policy) do not change it, and `Set-WinUILanguageOverride` alone stays pending until an interactive logon. The harness writes `PreferredUILanguages` under `HKCU\Control Panel\Desktop`, which the next SSH logon reads.
- The regional format: `Set-Culture` (en-US: `.` and `,`; ru-RU: `,` and a space, list separator `;`).
- `CHAR` and `CODE` follow the system code page: `Set-WinSystemLocale` and a restart (en-US 1252, ru-RU 1251).

`host.ps1 -do env` reports all of these and starts Excel to read its interface language (`LanguageSettings.LanguageID(2)`: 1033 or 1049, the locale's `lcid`); a run refuses to start when any of them is not the locale's.

## How a job runs

`remote.ts` writes the books here and a `job.json` listing each book's steps, copies the directory to `C:\xlcalc\work\` with `scp`, runs `xl.ps1` on it and copies back `out\`. `xl.ps1` starts one Excel, opens the books one by one, takes their steps (a full calculation, edits, reads of ranges, shown text, a structure or table operation, a save) and writes each book's answer as soon as it is done; at the end it checks that no book is left open and quits Excel.

- A range is read in one call: `Value2` gives numbers, text, booleans and errors as COM codes (`-2146826281` is `#DIV/0!`, 0x800A0000 plus Excel's error number), so the wrapper and mirror columns the Mac reads are not needed. They stay in the books, which are the Mac's.
- A formula case's book is checked to still hold its formulas (`HasFormula`): a book Excel repaired on opening loses them. A book Excel does not open, opens repaired or hangs on is halved down to the formula at fault, which is stored as not taken.
- A job holds up to 40 books. Each SSH call costs seconds on this machine (a new PowerShell starts), so the harness shares one SSH connection per run and puts as many books in a job as it can.
- A job that runs past its time (Excel waits on a window nobody sees) is stopped with Excel; the books done keep their answers and the rest run again. Excel that is still exiting from the previous job is waited for: a new one fails to start meanwhile.
- Excel started over SSH runs in session 0. It opens no file ("cannot access the file") unless the system profile has a desktop folder, `C:\Windows\System32\config\systemprofile\Desktop` and its `SysWOW64` twin; `host.ps1 -do clean` makes sure of them, stops Excel processes left in session 0 and removes the jobs' files. An interactive user's Excel is left alone.
- A number is typed into a table through `Formula` (its text with a decimal point): `Value2` with a number into a table's header failed now and then with an invalid cast.
- A defined name is added with `Names.Add` and the formula `=0`, then given its formula through `RefersTo`: `Names.Add` reads its formula in the user's language (`=0.2` is an error in ru-RU), `RefersTo` in English.
- Notes, hyperlinks and frozen panes are made with `AddComment`, `Hyperlinks.Add` and the active window's `FreezePanes`, which works in session 0. Excel 2024 has no object model for threads of comments: `AddCommentThreaded` answers "Excel VBA Object Model for Threaded Comments is not supported in your version of Excel", and `CommentsThreaded` lists none of a file's threads. A saved book goes to the job's `out\`, a folder below the book's, so a relative hyperlink address comes back one folder up (`../other.xlsx`).

## What is not run here

- The xlsx round trip, the saved books and the colours of `bench/coverage/xlsx.ts`, and the ru-RU formula text snapshot (`bench/locale`): they work through COM (`SaveAs`, `Interior.Color`, `FormulaLocal`) but are not ported; only the filters and the built-in formats are.
- The spill order runs (`order-runs.json`) record Excel for Mac's varying choice and are no set.

## The first probe

`probe.ps1` is the probe that first showed COM working over SSH, kept as it ran: its Office settings do not change the interface language.

## How it was checked

Every set the Mac has answers for was computed by Excel 16.0.20430 in its locale; Excel took every formula and opened every book.

| what | locale | cases | as Excel for Mac | the engine as Windows, or for a pinned reason |
|---|---|---:|---:|---|
| formulas: measurement, probes, 40 case sets | en-US | 34,502 | 34,469 | 34,391, the 111 others in `WIN_KNOWN` |
| formulas: measurement, 8 case sets | ru-RU | 7,297 | 7,136 | 7,291, the 6 others in `WIN_KNOWN` |
| the order of text, the set `collation-more` (Windows alone) | ru-RU | 62 | | 62 |
| formulas of the old kind, 4 sets | en-US | 6,999 | 6,998 (the one other is `NOW()`) | through the Mac's answers |
| cells shown, 8 sets | en-US | 9,819 | 7,929 | 7,979, the others by the rules of `WIN_SHOWN` |
| cells shown, 6 sets | ru-RU | 3,130 | 2,434 | 2,569, the others by the rules of `WIN_SHOWN` |
| tables, 9 sets and 5 | en-US, ru-RU | 133 and 38 | 128 and 34 | all but the 12 the Mac pins too, and `headers-ru-held:tiny` (a number typed as its text takes the format of typed exponents) |
| spill, 15 sets | en-US | 175 | 162 | all but 3 pinned and 3 that vary between runs |
| workbook, 7 sets | en-US | 98 | 95 | all but 5 pinned |
| structure operations, 18 sets | en-US | 317 | 235 (sizes in other units) | all but 16 pinned |
| xlsx filters | en-US | 8 files | | all |
| conditional formats, 7 sets | en-US, ru-RU | 52 books, 4,089 cells | 3,856 of the 3,885 cells the Mac read (28 `#` counted in pixels, and `"b">"в"`, which the Mac takes otherwise) | every cell, but `#` counted in pixels and a colour channel one apart at an exact fraction past a percentile; the rules of every case but the bar lengths and the order the Mac pins too; of the 15 rules Excel made itself (`made`), 203 of 204 cells on the book read from its file |
| notes, hyperlinks and frozen panes, 4 sets | en-US | 47 | 47 (the notes Excel made have its own user as author) | all but the pinned differences the Mac has, and a relative address saved one folder down |
| text typed into cells, 3 sets | en-US, ru-RU | 911 and 911 | all but 34: the formats `m/d/yyyy` (en-US, 12) and `h:mm` (ru-RU, 22), which Windows takes as the built-in formats 14 and 20 | all but 167, pinned in `test/entry-known.ts` (the object model, English names in ru-RU, a reference's format) |
| xlsx built-in formats 1 to 58 | en-US, ru-RU | 58 and 58 | 50 to 58 the short date (14), as the Mac en-US shows them by its own | 1 to 49; 50 to 58 pinned (read as `m/d/yy`) |

The entry sets (`typed`, `typed-held`, `typed-held-2`) ran with `XLWIN_CODEPAGE=any` in both locales. The sets `known-held` and `legacy/ifs-text` (45 formulas), `lookup-held`, `lookup-rules`, `lookup-own-held` and `stat-edges-held` (153) ran with `XLWIN_CODEPAGE=any`, in en-US with the system locale left in ru-RU; all their answers are the Mac's, the statistics on edge data bit for bit. The parts sets (`files`, `struct-held`, `frozen-held`, `threads`) ran the same way.

What Excel for Windows does otherwise than Excel for Mac, by reason, and which of the two the engine follows:

- Its regional data in ru-RU: Latin sorts before Cyrillic (`"a"<"а"` is TRUE), month names have a capital (`Январь`) and short ones three letters (`фев`, `ноя`), the short and long time no leading zero; built-in formats by number take Windows's own codes (en-US 14 is `m/d/yyyy`). The engine follows Windows, but for the long date `[$-F800]`, which has no weekday on Windows and keeps the Mac's here.
- `CHAR` and `CODE` by the Windows code pages, and 63 for a character outside them. The engine follows Windows.
- Trigonometry at its poles and complex functions of large arguments from the 13th digit on (`TAN(PI()/2)`), `WEEKNUM` with return type 22, `WEIBULL` with a scale of 1E-300. The engine follows the Mac.
- A table's name in ru-RU: Excel for Windows makes `ИСТИНА_` of a table renamed `ИСТИНА`, as both Excels refuse such a defined name, and keeps a table renamed `TRUE`; Excel for Mac keeps `ИСТИНА` and makes `TRUE_`. The engine refuses both, following Windows for `ИСТИНА`.
- Layout in pixels: where a value meets the column's edge, the cell shows what the engine shows at another width; row heights in points (3/4 of the Mac's) and widths rounded to whole pixels. The engine lays cells out as the Mac does.

Both Excels agree and the engine does not in 80 formula cases, all of them known differences the Mac's tests pin already (`KNOWN`, `RU_KNOWN`); `bench/windows/differences.md` lists them.
