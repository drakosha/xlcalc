# Locales

Excel depends on its language and regional settings in two ways. It displays and parses formulas in the user's language, and it calculates some values by regional rules: text parsed as a number or a date, numbers and booleans converted to text, format codes, currency. Whatever its locale, a workbook here stores formulas in the invariant form of xlsx files (English function names, `,` between arguments, `.` as the decimal separator). en-US and ru-RU are supported for both formula text and calculation.

## In short

Workbooks store formulas the way xlsx files do: English function names, `,` between arguments, `.` as the decimal separator. A locale changes how formulas are entered and displayed, as Excel displays them in another language, and how a workbook calculates, as Excel calculates with those regional settings. en-US is built in. ru-RU (Excel with the Russian interface) is a separate module, which a bundle includes only when it is imported: `registerLocale` loads it, once, for workbooks in ru-RU and for its tag.

```js
import { LOCALES, registerLocale, toInvariant, toLocal, validateLocalFormula } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(toLocal('IF(B1>0.5,SUM({1,2;3,4}),#N/A)', ruRU).formula) // → ЕСЛИ(B1>0,5;СУММ({1;2:3;4});#Н/Д)
registerLocale(ruRU)
console.log(LOCALES.get('ru-RU') === ruRU, [...LOCALES.keys()]) // → true [ 'en-US', 'ru-RU' ]
console.log(toInvariant('=СУММ(A1:A3;1,5)', 'ru-RU').formula) // → SUM(A1:A3,1.5)
const bad = validateLocalFormula('=СУММ(A1;1.5)', 'ru-RU')
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax Здесь не может стоять «.» 10
console.log(validateLocalFormula('=ОКРУГЛ(Данные!B2*1,2;0)', ruRU, { sheets: ['Данные'] }).canonical) // → ROUND('Данные'!B2*1.2,0)
```

A host that knows a locale only by its tag, from a file or a user's choice, loads it with `loadLocale` from `xlcalc/locales`, which imports the package's module of that tag and loads it as `registerLocale` does, for every call that takes a loaded locale, a pool's threads included. It is a subpath of its own, without the engine, so the bundles of the root and of `xlcalc/xlsx` carry no locale module but en-US; where a bundler splits, the module comes as a chunk of its own, only when it is loaded. A locale loaded already comes back without importing its module again, and a tag that is not one of the package's locales is an `InputError`:

```js
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'
import { loadLocale } from 'xlcalc/locales'

const read = await readXlsx(data)
if (read.book.locale) await loadLocale(read.book.locale)
await writeXlsx({ book: read.book })
```

`LOCALES` lists the loaded locales by tag. Every function that accepts a locale accepts the tag of a loaded locale or the locale object, and an object needs no loading where no workbook is involved: `toLocal`, `toInvariant`, `validateLocalFormula`, `parseEntry`, `displayValue`, `entryWidth`, `builtinFormatCode`, `numberFormatCategory`, `numberFormatChoices`, `entryText`, `needsApostrophe`, `errorText`, `scanReferences`, `refAt`, `callAt`, `nameAt`, `functionSignature`, `functionDescription`, `compareValues`, `valueComparator`, `validateSheetName`, `validateName` and `validateNames`. A workbook's `locale` is a tag, and a workbook in a locale that is not loaded causes an `InputError` that says what to import, wherever the workbook goes (`compute`, sessions, `editStructure`, `fitColumns`, `writeXlsx`, a pool); `readXlsx` keeps a file's locale as it reads it. A thread pool loads in its threads the locales loaded where it runs, including those loaded later. A locale is a plain object (the `Locale` type), so formula text in another language is just another object of the same shape; one without calculation rules parses and displays values as en-US does. Adding a locale to the package takes one module ([docs/locale.md](locale.md#adding-a-locale)).

```js
import { compute } from 'xlcalc'

try {
  compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] })
} catch (e) {
  console.log(e.message) // → book.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU), or await loadLocale('ru-RU') from 'xlcalc/locales' (loaded: en-US)
}
```

`toLocal` and `toInvariant` return `{ ok: true, formula }` or `{ ok: false, code, message, at }` with the position in the given text. They translate function names, TRUE and FALSE, error literals, the decimal separator and the separators of arguments and array constants (`{1;2:3;4}` in ru-RU); strings, defined names, references and whitespace are kept as they are. `toLocal` puts a sheet name in quotes only where Excel does, using `quoteSheetIfNeeded` with the locale's booleans ([Addresses and helpers](addresses.md)), as every text displayed to users does (a viewport's `formulaLocal` and `editText`, the text of `explain`): the stored `'Data'!B2*2` becomes `Data!B2*2` in every locale, en-US included, `'Q1 data'!A1` keeps its quotes, and a sheet named `ИСТИНА` becomes `'ИСТИНА'!A1` in ru-RU. `toInvariant` accepts a name with or without quotes and keeps it as typed; a stored formula quotes every sheet, as in its canonical text. `toInvariant` drops a leading `=` and also accepts English function names, TRUE/FALSE and error texts; a function the locale does not name (an external one, for example) passes through unchanged. Text that would mean something different once translated (`1.5` typed in ru-RU, where `.` is not a decimal separator) is a `syntax` error at its position. `validateLocalFormula(text, locale, ctx?)` is `validateFormula` for typed text: the same result with the canonical stored text, and positions in `text`. The messages of all three are in the locale's language.

The ru-RU data come from Excel itself. `bench/locale/snapshot.ts` opens workbooks with 6,381 formulas (the coverage run's cases, the probes, a call of every function on Microsoft's list that the coverage run lacks, and literals) in Excel 16.113 for Mac with the Russian interface and reads each formula back in English and in Russian; `bench/locale/generate.ts ru-RU` writes the formula text of the ru-RU locale file, `src/locales/ru-RU.ts`, from that text (`test/excel/locale/ru-RU.json`). `test/locale-excel.test.ts` checks that both translations reproduce Excel's text of all 6,381 formulas, and `test/locale-roundtrip.test.ts` that the 34,556 formulas of the coverage run, the case sets and the probes round-trip unchanged through ru-RU. Excel 16.113 for Mac does not know `ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE` and `REGEXTEST`, and removes `CALL` and `REGISTER.ID` from the workbooks it loads, so these keep their English names; so do the error literals newer than `#GETTING_DATA` (`#SPILL!`, `#CALC!` and the rest), which it does not accept as literals in a formula.

A workbook is calculated in its own locale, `BookInput.locale` (the tag of a loaded locale: `'en-US'` by default, or `'ru-RU'` once loaded), as Excel calculates with those regional settings, whoever opens the workbook; a viewer still sees the formulas in their own language through `toLocal`. The locale affects how text is parsed as a number, a date or a boolean, how numbers and booleans are converted to text, the format codes of `TEXT` and what cells display, `DOLLAR` and `FIXED`, error texts, the words and numbers in criteria, and text sort order:

```js
import { compute, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const cells = { A1: { formula: 'VALUE("1,5")+VALUE("1 234,5")' }, A2: { formula: '"a"&1.5&TRUE' }, A3: { formula: 'TEXT(45688.75,"ДД ММММ ГГГГ ч:мм")' }, A4: { formula: 'DOLLAR(-1234.567)' }, A5: { formula: 'VALUE("1.5")' } }
const r = compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells }] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
const v = r.sheets[0].cells
console.log(v.A1.value, v.A2.value, v.A3.value, v.A4.value, v.A5.value) // → 1236 a1,5ИСТИНА 31 Январь 2025 18:00 -1 234,57 ₽ 46143
console.log(r.volatile.map((c) => c.cell).join(',')) // → A5
```

In ru-RU, as Excel 16.113 for Mac calculates with the Russian interface and locale, and as Excel for Windows where the two differ: numbers in text use a decimal comma and a space as the group separator (`1 234,5`) and the ruble sign after the amount (`1 234,5 ₽`); dates are day-month-year (`31.01.2025`, `31/1/25`, `2025-01-31`, `31 янв 2025` with a month named by three or more letters of its nominative name; English month names are not recognized); `ИСТИНА` and `ЛОЖЬ` are the booleans in text (`"TRUE"` is not a boolean there); numbers are converted to text with a decimal comma (`a1,5`), complex numbers too; `TEXT` accepts the Russian codes (`ДД.ММ.ГГГГ`, `ч:мм:сс`, `# ##0,00`, `Основной`, `[Красный]`, `[Цвет10]`) and treats English letters as literals, while cells read their formats as xlsx files store them (`dd/mm/yyyy` displays `31.01.2025`); `DOLLAR` gives `1 234,57 ₽`; month names are capitalized (`Январь`) and short ones have three letters (`фев`); and text sorts Latin before Cyrillic, as Excel for Windows sorts (Excel for Mac puts Cyrillic first). Date text without a year (`1.5` is 1 May, `15 янв`) takes the year of the workbook's clock, as Excel takes the current year, so a cell that parses such text depends on the clock as `TODAY` does: it is listed in `volatile`, and a session recalculates it on `setClock`. The rules were verified case by case in Excel ([docs/locale.md](locale.md)).

Messages for users are in the workbook's locale too: why a typed formula does not parse (`parseEntry`, and `toLocal`, `toInvariant` and `validateLocalFormula` in their locale), why a structure operation is rejected, and the text of `explain`, whose numbers, formulas, booleans and errors are written as the locale writes them, and its references as Excel writes them (`План!B1`). Codes, positions and the fields of `explain`'s tree stay unchanged, so a program reads them the same way in every locale.

```js
import { createSession, parseEntry, registerLocale, editStructure } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = { locale: 'ru-RU', sheets: [{ name: 'План', cells: { A1: { value: 400 }, B1: { formula: 'ROUND(A1*1.2,0)' } } }] }
console.log(createSession(book).explain('План', 'B1').text)
// → План!B1 = 480  ← ОКРУГЛ(A1*1,2;0)
//     План!A1 = 400
console.log(parseEntry('=ОКРУГЛ(A1;2;3)', { locale: 'ru-RU' }).message) // → Функция ОКРУГЛ принимает 2 аргумента, а указано 3
console.log(editStructure(book, { op: 'deleteSheet', sheet: 'План' }).message) // → В книге должен остаться хотя бы один лист
```

`validateFormula`, which has no locale, responds in English, as do the reasons of cell errors and name checks. A locale's messages are part of its module ([docs/locale.md](locale.md#messages)), and a locale object without messages of its own uses those of the loaded locale with its tag, otherwise English.

## Loading a locale

en-US is built in. ru-RU is a separate module, `xlcalc/locales/ru-RU`, so a bundle that has no ru-RU workbook does not carry its 52 KB of formula text, rules and messages. A host imports it and loads it once, before a workbook in that locale is calculated:

```js
import { compute, LOCALES, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = { locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { formula: 'VALUE("1,5")' } } }] }
console.log(compute(book).sheets[0].cells.A1.value, [...LOCALES.keys()])
// → 1.5 [ 'en-US', 'ru-RU' ]
```

- `LOCALES` lists the loaded locales by tag: `en-US`, then each registered locale.
- A workbook's `locale` is the tag of a loaded locale. A workbook in a locale that is not loaded causes an `InputError` that says what to do: `book.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU), or await loadLocale('ru-RU') from 'xlcalc/locales' (loaded: en-US)`. So does a tag passed to a function that accepts a locale. Any other tag gets a message listing the loaded ones, and naming the intended tag when only its case differs (`ru-ru`): tags are compared as written.
- Every function that accepts a locale accepts the tag of a loaded locale or the locale object: `toLocal`, `toInvariant`, `validateLocalFormula`, `parseEntry`, `displayValue`, `entryWidth`, `builtinFormatCode`, `numberFormatCategory`, `numberFormatChoices`, `entryText`, `needsApostrophe`, `errorText`, `scanReferences`, `refAt`, `callAt`, `nameAt`, `functionSignature`, `functionDescription`, `compareValues`, `valueComparator`, `validateSheetName`, `validateName` and `validateNames`. An object needs no loading there, since no workbook is calculated: `toLocal(formula, ruRU)` works in a bundle that never calls `registerLocale`.
- Importing the module has no effect by itself, so a bundler leaves it out of a bundle that does not use it. Loading the same object again does nothing; another object with an already loaded tag causes an `InputError`, as does a locale without the rules a workbook calculates by.
- `readXlsx` keeps a file's locale (`ru-RU`) as it reads it, loaded or not; calculating or writing the workbook requires the locale to be loaded.
- A pool's threads load the locales loaded where the pool runs, including those loaded after it started ([docs/async.md](async.md#locales)).

What matters to workbooks in every locale stays in the core, so an en-US workbook displays the same whatever is loaded: the widths of Cyrillic letters (`src/numfmt/glyphs.ts`) and ru-RU's month and day names, which the format code `[$-419]` produces in a workbook of any locale (`src/locale/lcid.ts`).

## Formula text

A locale is a plain object (`Locale`, `src/locale/types.ts`): its formula text and, for a locale a workbook can calculate in, its calculation rules (`regional`) and messages. en-US's is `enUS` from the root entry, ru-RU's is `ruRU` from its module, and `LOCALES` holds the loaded ones by tag (`LOCALES.get('ru-RU')`). `toLocal`, `toInvariant` and `validateLocalFormula` accept a loaded locale's tag or a locale object, and translate a formula token by token: function names, TRUE and FALSE, error literals, the decimal separator, and the separators of arguments and array constants. Strings, defined names, references and whitespace are copied as they are. A sheet name is displayed in quotes only where Excel puts them, using `quoteSheetIfNeeded` with the locale's booleans: the stored `'Data'!B2` becomes `Data!B2` in every locale, `'ИСТИНА'!A1` stays quoted in ru-RU, and `'TRUE'!A1` becomes `TRUE!A1` there; when parsed back, a name with or without quotes is kept as typed.

The output is parsed again in the target locale, so text that would change its meaning in translation is a syntax error at its position: `1.5` typed in ru-RU, where `.` is not a decimal separator.

In ru-RU on the Mac, the list separator is `;`, the decimal separator `,`, and array constants are written `{1;2:3;4}`, with `;` between columns and `:` between rows.

`toInvariant` also accepts English function names, TRUE, FALSE and error texts inside Russian text. The messages of `toLocal`, `toInvariant` and `validateLocalFormula` are in the locale's language ([Messages](#messages)), with positions in the typed text.

Structured references are translated too: the specifiers by the locale's `tableItems` (in ru-RU `#Все`, `#Данные`, `#Заголовки`, `#Итоги` and `#Эта строка`, which Excel displays as `@`) and the separator between their items (`Sales[[#Заголовки];[Qty]]`). `toLocal` displays them as Excel does, and given the table of the formula's cell, drops its name where Excel drops it ([Tables](tables.md)).

### Where the ru-RU names come from

The Russian names were read from Excel 16.113 for Mac with the Russian interface. `bench/locale/snapshot.ts` wrote 6,381 formulas to workbooks and read each one back both as `formula` and as `formula local`, along with Excel's international settings (`test/excel/locale/ru-RU.json`). `bench/locale/derive.ts` pairs the function calls of each English formula with those of its Russian text, and `bench/locale/generate.ts ru-RU` writes them into the ru-RU locale file, `src/locales/ru-RU.ts`.

513 of the 521 functions on Microsoft's list got a Russian name this way (468 of them differ from the English name). Excel for Mac does not know 6 functions (`ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`) and removes 2 from the workbooks it loads (`CALL`, `REGISTER.ID`); these keep their English names. Excel accepts 8 error literals in formulas; `#SPILL!`, `#CALC!` and newer errors are not literals for it.

## Calculating in a workbook's locale

A workbook carries its locale in `BookInput.locale`, the tag of a loaded locale: `'en-US'` (the default), or `'ru-RU'` once loaded. An xlsx file stores no calculation locale, so the same file gives different values in Excel on a Russian machine and on an English one: `VALUE("1.5")` is 1.5 in one and 46143 in the other. A workbook that carries its locale gives the same values to everyone, while each viewer still sees formulas in their own language through `toLocal`.

A tag that is not loaded causes an `InputError` ([Loading a locale](#loading-a-locale)). A host can build its own `Locale` objects for formula text; one without calculation rules parses and displays values as en-US does, and `registerLocale` rejects it, so a workbook calculates only in a locale with its own rules.

In every locale, a percentage in text keeps the digits written: `VALUE("1,1%")` in ru-RU, like `VALUE("1.1%")` in en-US, is 0.011, the number the text represents once the decimal point moves two places, as Excel parses it in VALUE, operators and numeric arguments, not 1.1 divided by 100 (0.011000000000000001; the difference shows when magnified: `(VALUE("1.1%")-0.011)*1E+20` is 0). `NUMBERVALUE` divides by 100 for each `%` in turn, and the `%` operator divides its operand, as Excel does (`NUMBERVALUE("1.1%%")` is 1.1/100/100). Currency with a percent sign is not a number (`"$1.1%"`, `"1,1 ₽%"`). The `pct-text` set verified these in both versions of Excel and both locales.

The locale is included in the hashes only when it is set and is not en-US: `inputHash` gets a line `L "<tag>"`, and the values hash gets `;locale=<tag>` after the configuration version. Workbooks without a locale and en-US workbooks hash as they did before.

### What changes in ru-RU

| | ru-RU, as Excel 16.113 for Mac calculates, and as Excel for Windows where the two differ |
|---|---|
| numbers in text | `1 234,5` and `1 2345` (a plain space before groups of three digits or more; a no-break space is not a separator), `,5`, `(1,5)`, `1,5%`, `%(1,5)`, `1,5E+3`, `1 1/2`, `1 234,5 ₽` (the ruble sign after the amount). `$`, `р.` and `1.234,5` are not numbers |
| dates in text | day, month, year separated by `.`, `/` or `-`: `31.01.2025`, `1/2/2025` is 1 February, `31.01.25`; also `2025-01-31` and `2025.1.31`. A month given by three or more letters of its nominative name, in any case: `31 январь 2025`, `31.янв.2025`, `15 июн. 2025` (not the genitive `января`, and not English names). `янв 2025` and `янв-25` are the 1st |
| two numbers | a day and a month of the current year (`1.5`, `15 янв`), or a month and a year when the second number is greater than 12 (`1.13` is January 2013, `1.2025`) |
| times | as in en-US, with a decimal comma in seconds (`18:30:15,5`) and the English `AM`/`PM` |
| numbers as text | a decimal comma: `"a"&1.5` is `a1,5`, `1,23456789012345E+20`; complex numbers too (`1,5-2,25i`) |
| booleans | `ИСТИНА` and `ЛОЖЬ` in both directions, in any case when parsed; `"TRUE"` is `#VALUE!` where a boolean is expected |
| TEXT codes | `Д`, `Г`, `ч`, `с` in either case; `М` is the month, or minutes after hours or before seconds; `м` is always minutes; `,` is the decimal point and a space the group separator; `Основной` is General; Latin letters are literals (`yyyy` stays `yyyy`); month names in the nominative, capitalized (`Январь`, `Я` for `МММММ`), short ones with three letters (`янв`, `фев`, `сен`, `ноя`), as Excel for Windows writes them; days `пятница`, `Пт` |
| DOLLAR, FIXED | `1 234,57 ₽`, `-1 234,57 ₽`, `-0,00 ₽` (a minus sign even when rounded to zero), `1 234 567,89` |
| VALUE and NUMBERVALUE | parsed as above; NUMBERVALUE's default separators are `,` and a space, and the rest of its text is parsed as VALUE parses it (`NUMBERVALUE("1.5")` is a date) |
| complex text | each part is parsed as VALUE parses it: `IMSUM("1.5+2i")` is `46143+2i` |
| errors, arrays as text | `#Н/Д` and the other Russian error texts in VALUETOTEXT, ARRAYTOTEXT and criteria; ARRAYTOTEXT writes `{1,5;2:ИСТИНА}` and `1,5; 2` |
| criteria | the locale's numbers, dates and words in criterion text and in cell text: `">1,5"`, `">=01.01.2026"`, `"01.05"`, `"ИСТИНА"` (`"TRUE"` is text) |
| text order | Latin before Cyrillic, as in en-US and as Excel for Windows sorts (Excel for Mac ru-RU puts Cyrillic first) |
| times of built-in formats | the hour without a leading zero (`6:00` for formats 20 and 21 and `[$-F400]`), as Excel for Windows displays them |
| CHAR, CODE | Windows-1251 (en-US: Windows-1252) |

### Dates without a year

Excel completes a date without a year with the current year: `1.5` in ru-RU, `1/31` in en-US. The engine uses the year of the workbook's clock, so a workbook's values remain a function of its input and clock.

Whether a cell reads the year depends on its data (`VALUE(A1)` reads it only when A1 contains such text), so the engine finds out during calculation. The rules read the year through a clock object that sets a flag; the engine records every formula that set it and treats that formula as volatile. Such cells are listed in `ComputeResult.volatile` and recalculated by `setClock` and in live sessions. A criterion or criteria index whose parsing read the year is not cached, so every formula that needs it reads the year itself.

### Text order

Text sorts Latin before Cyrillic in both locales, as Excel for Windows sorts in every locale; Excel for Mac ru-RU puts Cyrillic first, and the engine follows Windows there. The order is case-insensitive and accent-sensitive: `"аб"="АБ"` is TRUE, while `"ё"="е"` and `"й"="и"` are FALSE. An accent and `ё` matter only where the letters are otherwise equal, while `й` is a separate letter after `и`:

- `"ёа"<"еб"` is TRUE, as is `"éa"<"eb"`: the second letter decides before the accent.
- `"шел"<"шёл"` is TRUE: with the letters equal, the unaccented one comes first.
- `"йа"<"иб"` is FALSE: `й` comes after every word that starts with `и`.
- `ъ`, `ы`, `ь` come in that order, after `й` and before `э`, `ю`, `я`.

`Intl.Collator` with the en-US order (`collation: 'en-US'` in both locales) and `sensitivity: 'accent'` reproduces Excel for Windows's order on all 236 tested cases in ru-RU and 174 in en-US. Operators, criteria, lookups and sorting all use it.

Text spelled with é as one code point or as e with a combining accent is equal in operators, criteria and `UNIQUE`, as in Excel, while the exact lookups (`MATCH` and `XMATCH` with match type 0, `XLOOKUP` and `XMATCH` searching from the first or the last, `VLOOKUP` and `HLOOKUP` with `FALSE`) tell the two spellings apart, as Excel's `MATCH` does; the binary searches (`XLOOKUP` and `XMATCH` with search mode 2 or −2) and the approximate ones take them as `=` does, which Excel was not measured on. A soft hyphen (U+00AD), which `Intl.Collator` ignores, still tells two texts apart, as in Excel: `a` and `a` with a soft hyphen after it are not equal.

## Messages

A user working in a workbook's locale reads the engine's messages in that language: why a typed formula does not parse, why a structure operation is rejected, where a value comes from. In ru-RU they are in Russian:

| where | the message in ru-RU |
|---|---|
| `parseEntry('=СУММ(1', { locale: 'ru-RU' })` | `Ожидалась ), а формула закончилась` |
| `parseEntry('=ОКРУГЛ(1;2;3)', { locale: 'ru-RU' })` | `Функция ОКРУГЛ принимает 2 аргумента, а указано 3` |
| `validateLocalFormula('=СУММ(1.5)', ruRU)` | `Здесь не может стоять «.»` |
| `editStructure` deleting rows through a legacy array formula | `Нельзя изменить часть массива S!C1 (C1:C3)` |
| `editStructure` renaming a sheet to a name already in use | `Лист «Итоги» уже есть (имена листов не различают регистр)` |
| `explain` | `План!B2 = 480  ← ОКРУГЛ(B1*1,2;0)` |

### Which locale

- `parseEntry` writes in `opts.locale`, and `toLocal`, `toInvariant` and `validateLocalFormula` in the locale they are given. A locale object uses its own `messages`; without them, those of the loaded locale with its tag, and English for any other tag.
- `editStructure`, and a session's `editStructure` alike, on the same thread, in a pool thread or synchronously, writes in the workbook's `locale`.
- `readXlsx` and `writeXlsx` write their issues in `opts.locale`, or English without it, never in the workbook's locale: a file does not say what language its reader uses ([docs/xlsx.md](xlsx.md#the-language-of-the-messages)). A locale's `messages` hold their texts alongside the engine's; a locale without them gets the English ones, which the `xlcalc/xlsx` subpath carries, so a bundle that reads no files does not include them.
- `explain`'s `text` is written in the workbook's locale: its words (`(2 ячейки)`, `перенос из`, `(ошибка из План!A2)`), numbers with the locale's decimal separator and every digit of the value (`0,30000000000000004`, `1E+21`), text in quotes with a quote doubled as in a formula (`"a""b"`), formulas as `toLocal` displays them, references with a sheet in quotes only where Excel puts them, booleans in the locale's words (`ИСТИНА`), errors as the workbook's cells display them (`#Н/Д`), and an empty cell as `пусто`. The tree's `ref`, `value`, `formula` and `origin` stay as the workbook stores them, for programs.

Codes (`code`) and positions (`at`) are the same in every language, so a host that chose its own wording by code can keep doing so, or display `message` as is.

### What stays in English

- `validateFormula`, which has no locale, and every message of an en-US workbook or call.
- The reasons of cell errors and name issues (`diagnostics[].reason`), and the message of `validateName`.
- `InputError`, which reports a programming error rather than a user's mistake.

### How the Russian is written

As an interface addresses a user: short sentences, a capital letter at the start and no final period, no em dash. Where Excel ru-RU has its own message for the same rejection, the text stays close to it (`Нельзя изменить часть массива`). What comes from the typed formula is written in the locale: function names (`ОКРУГЛ`), separators (`Разделитель «;» вне скобок функции`), and specifiers of structured references (`#Данные`). Names of sheets, tables and defined names stand in «» (`Лист «Итоги» уже есть`), and references as they are (`S!C1`). A number takes its noun in the correct grammatical form: `1 строку`, `2 строки`, `5 строк`.

### The catalog

Each message is a key with parameters (`MessageParams` in `src/locale/messages.ts`): the parser, the checks and the structure operations name the message where they reject something, and its text is produced only then, so messages cost nothing while a workbook calculates. A locale's `messages` produce the text of each key from its parameters and the locale's words (`Words`: function names, specifiers, numbers, values and formulas as the locale writes them). en-US's file has every message, the engine's English text; a locale without the text of a message falls back to the English one.

`test/messages-catalog.test.ts` checks that ru-RU writes every message of en-US and shows every parameter the English text shows, and `test/messages-paths.test.ts` that each way a message reaches a user (a typed formula, each rejection of a structure operation in every kind of session, `explain`) produces it in Russian in ru-RU and in English elsewhere.

## How the engine does it

`src/locale/rules.ts` defines `Rules`, one object per workbook. en-US is the engine's own code. Another locale is built by `regionalRules` (`src/locale/rules/regional.ts`) from its data (`Locale.regional`): separators, the order of dates, the currency and where its symbol goes, month and day names, TEXT codes, the collation tag, the code page of CHAR and CODE, the word for new table columns' names, and the words for TRUE, FALSE and errors.

The engine builds the rules when it creates a workbook's state and passes them to operators and criteria through the runtime, and to every function through `Cx.rules`. The argument kinds in `src/args.ts` convert values with the rules, and functions that read or write text themselves (TEXT, DOLLAR, FIXED, VALUE, the complex-number functions, the lookups and others) use them directly. Code outside a workbook, such as the parser, works in en-US.

## Adding a locale

A locale added to the package lets workbooks in it calculate as Excel does in that language and region, and lets hosts display formulas in its language. It takes one module and three lines:

- the module `src/locales/xx-YY.ts`, which holds everything the engine knows about the locale and exports it as `xxYY`, the tag without its hyphen;
- its tag in `PACKAGE_LOCALES` (`src/locale/registry.ts`), which names the module in the message about a workbook in that locale that is not loaded, and lets `readXlsx` keep it;
- its import in `LOCALE_MODULES` (`src/locale/modules.ts`), through which `loadLocale` loads it and a pool's thread gets its messages, and through which the tests and harnesses load every locale of the package;
- its subpath in `package.json`, `"./locales/xx-YY"`, like ru-RU's.

`test/locale-modules.test.ts` checks that the four agree. No engine code changes, as long as the locale's rules are ones the engine reads from data ([What the data cannot express](#what-the-data-cannot-express)), and neither do the tests and harnesses that compare it with Excel ([Tests and harnesses](#tests-and-harnesses)).

### The locale module

A locale module exports one `Locale` (`src/locale/types.ts`) and has no side effects. ru-RU's, `src/locales/ru-RU.ts`, is the one to copy. It has three parts.

The formula text: the tag, the separators, the words for TRUE and FALSE, error literals, function names and the specifiers of structured references. The file holds most of it as the constant `FORMULA_TEXT` between two markers:

```
// ---- formula text, generated by bench/locale/generate.ts: do not edit ----
...
// ---- end of formula text ----
```

`npx tsx bench/locale/generate.ts xx-YY` rewrites the lines between the markers from Excel's snapshot of the locale (`test/excel/locale/xx-YY.json`) and leaves the rest of the file as it is. The tag and the specifiers are written by hand after the markers: the snapshot's workbooks have no tables, so ru-RU's specifiers (`#Все`, `#Данные`, `#Заголовки`, `#Итоги`, `#Эта строка`) were taken from the `locale` table set.

`regional` holds the calculation rules, a `RegionalData` (`src/locale/rules/data.ts`) written by hand:

| field | what it controls | ru-RU |
|---|---|---|
| `decimal`, `group` | the separators of numbers in text, FIXED and DOLLAR, and NUMBERVALUE's defaults | `,` and a space: `1 234,5` |
| `dateOrder` | the order of day, month and year in date text: `dmy`, `mdy` or `ymd` | `dmy`: `31.01.2025`, and `1.5` is 1 May |
| `currency` | DOLLAR's format, and where VALUE expects the symbol | `₽` after the amount, with a space between, and a minus sign: `-1 234,57 ₽` |
| `months`, `monthsShort`, `days`, `daysShort` | the names TEXT writes; date text accepts three or more letters of a full month name | `Январь`, `янв`, `пятница`, `Пт` |
| `codes`, `colors` | TEXT's letters for dates and times, the name of General, the color names | `ДД.ММ.ГГГГ ч:мм`, `Основной`, `[Красный]` |
| `dateSeparator` | what `/` in a cell's date format displays | `.`: `dd/mm/yyyy` displays `31.01.2025` |
| `lcid`, `longDate`, `longTime`, `longMonths` | the locale code (`[$-419]`), and the long date and time (`[$-F800]`, `[$-F400]`) | `пятница, 31 января 2025 г.` |
| `collation` | the `Intl.Collator` tag that orders text | `en-US`: Latin before Cyrillic, as Excel for Windows sorts |
| `builtin` | the built-in formats the locale displays its own way, as en-US codes | 14 is `dd.mm.yyyy` |
| `codePage` | the Windows code page of CHAR and CODE: the name of its published table and its characters | `windows-1251` |
| `columnWord` | the word in a new table column's name | `Столбец`, which gives `Столбец1` |
| `formats` (optional) | the names in the ribbon's Number Format box and its list's items, and the code its Currency item writes; en-US's if omitted | `Числовой`, `(все форматы)`, `#,##0.00\ "₽"` |

Two kinds of data matter to workbooks in every locale, so they live in the core rather than in the module, and an en-US workbook displays the same whether the locale is loaded or not. The widths of the letters of the locale's script in Calibri 11, for what a cell displays at its column's width, go into `src/numfmt/glyphs.ts`, next to the Latin letters and ru-RU's Cyrillic ones: text in any script may appear in any workbook, and a character that no table covers counts as a digit. The month and day names that the locale's code produces in a format (`[$-419]mmmm`) go into `src/locale/lcid.ts`, which `rules.ts` reads for the code, and the module takes them from there (`...RU_NAMES` in ru-RU's `regional`).

`messages` are the engine's messages for users in the locale's language ([Messages](#messages)): a function for each key of `MessageParams` that produces the text from its parameters and the locale's `Words`. A locale may omit any of them, and those are shown in English; `test/messages-catalog.test.ts` checks ru-RU's against en-US's keys, and a new locale's full set would be added there.

en-US's module holds its formula text and its messages, the English ones every locale falls back to, and is built into the root entry: its rules are the engine's code in `src/locale/rules.ts`, which also contains its code page (`windows-1252`) and the word `Column`.

### Where the data come from

Formula text is read from Excel for Mac in the locale's language: `bench/locale/snapshot.ts` writes the snapshot, Excel's own text of 6,381 formulas and its international settings, and `generate.ts` writes the generated part of the file from it. The Windows harness takes no snapshot of formula text.

Regional data start from Excel's international settings, which the snapshot keeps (separators, the currency), and from case sets calculated in Excel (names, TEXT codes, colors, the long date). They are checked against Excel for Windows too. Where Excel for Mac and Excel for Windows differ, Windows is the reference. ru-RU follows it for text sort order (Latin before Cyrillic), month names (`Январь`, not `январь`), the short and long time, and CHAR and CODE (Windows-1251, not Mac Cyrillic); it keeps one value from the Mac, the long date `[$-F800]`.

### What the engine reads from data

These rules were written for ru-RU and are now read from a locale's data. Each was compared with Excel in ru-RU only; in the other configurations the data allow, they are the engine's own rule, made by analogy with en-US, until a locale's Excel shows otherwise.

#### Dates in text

`dateOrder` positions the numbers: day first in ru-RU (`31.01.2025`, `31/1/25`), month first in `mdy` (`1/31/2025`), year first in `ymd` (`2025/1/31`, `25.1.31`). In every order the numbers are separated by `.`, `/` or `-`, `2025-01-31` is a date, and month names are parsed with the day first (`31 янв 2025`, `янв 2025`). Two numbers are a day and a month in day-first text (`1.5`), or a month and a year when the second is greater than 12 (`1.13` is January 2013); in the other orders they are a month followed by a day or a year, as en-US parses `1/31` and `1/2025`.

#### Currency in number text

VALUE and the other number text parsers accept the symbol where `currency.before` puts it: after the amount in ru-RU (`1 234,5 ₽`), and before it otherwise, with one sign before or after the symbol (`-€5`, `€-5`) and no fraction, as en-US parses `$`.

#### Case of names

Defined names, table names and sheet names are compared in upper case by Unicode's default rules, the same in every locale (`caseKey`, `src/collate.ts`), using a letter's full uppercase form: `straße` is the name `STRASSE`, and `ﬁ` is `FI`, as Excel for Mac and for Windows compare defined names and tables. The dotless `ı` and the long `ſ` remain themselves, not `I` and `S`, as Excel for Windows keeps them (Excel for Mac treats `ſ` as `S`). The words of formula text (function names, `TRUE` and `FALSE`, LET names and a locale's words) are compared letter by letter (`wordKey`): a letter whose uppercase form is two letters stays itself, so `gauß(1)` is not `GAUSS` and `falſe` is not FALSE (the `name-case-held` sets). Excel also compares sheet names letter by letter and accepts a sheet named `ß` next to one named `SS`; the engine compares them as names, a known difference. For Latin and Cyrillic, both keys match what the Russian and English rules give, character for character (`test/case-key.test.ts`). Turkish rules pair `i` with `İ` and `ı` with `I`, while these rules map `i` to `I` and keep `ı` and `İ` distinct; how Excel tr-TR compares such names has not been tested.

#### Number formats

TEXT rejects a point before a digit placeholder where the point is neither the decimal nor the group separator (`0.00` in ru-RU), and the group separator in a fraction format (`# ?/?` in ru-RU, where the group separator is a space).

#### Code pages

CHAR and CODE take codes 128 to 255 from the code page in the locale's data; `test/codepage.test.ts` checks the code page of every listed locale against the published table its data name (`test/fixtures/unicode/`), so a new code page needs its table there.

### What the data cannot express

Some locales need code as well as data:

- Scripts written right to left (Arabic, Hebrew): neither the translation of typed formulas nor what cells display has been checked with them.
- Native digits of a script (Arabic-Indic, Devanagari): number and date text is parsed with the digits 0 to 9 only, and TEXT's `[DBNum1]` and similar codes are not supported.
- Calendars other than the Gregorian and the Hijri of `B2` (Japanese eras, Thai Buddhist years): dates are Gregorian.
- Digit groups other than three (`1,00,000` in hi-IN): FIXED and DOLLAR group by three, and number text accepts groups of three or more digits.
- Double-byte code pages (East Asian locales): CHAR and CODE handle one byte.
- Date text with the month name before the day (`Jan 31, 2025`): parsed only by en-US's code.
- Rules the data do not describe. The parsers of number and date text, TEXT's codes and the rest were verified in ru-RU and en-US; a locale's own Excel may follow a rule neither of them has, and that rule needs code.

### Tests and harnesses

A new locale needs no change to the tests or the harnesses. They load every module in `LOCALE_MODULES` (`test/locales.ts`, `bench/coverage/locales.ts`), and the tests compare the engine with Excel in every loaded locale whose results are stored, and report a locale without them as skipped: `acceptance against Excel xx-YY: no answers of Excel xx-YY to the coverage measurement (test/excel/coverage/excel-xx-YY.json)`. The harnesses take the locale's tag and store its results in files of their own.

#### Where a locale's results are stored

Results are found by tag (`bench/coverage/locales.ts`). en-US's are at the top level of each directory under `test/excel/`, and another locale's in a directory named by its tag:

- the case sets: `cases/<set>.json` in en-US, `cases/ru-RU/<set>.json` in ru-RU, `cases/xx-YY/<set>.json` in xx-YY; the same for displayed text (`display/`) and tables (`tables/`);
- the coverage run: `coverage/excel-en.json` and `coverage/excel-ru.json`, as they were first named, and `coverage/excel-xx-YY.json` for another locale;
- the formula text: `locale/xx-YY.json`;
- Excel for Windows's results: the same layout under `win/`, and the built-in formats in `win/xlsx/xx-YY/formats.json`.

A directory of stored results named by a tag the package does not have fails a test (`bench/coverage/locales.test.ts`), so results are never left unchecked.

#### Collecting the results on the Mac

The steps of `bench/coverage/run.ts` take a locale as their last argument: `en`, `ru` or a tag. The snapshot comes first, since the locale's file is generated from it:

```sh
npx tsx bench/locale/snapshot.ts xx-YY             # the formula text, into test/excel/locale/xx-YY.json
npx tsx bench/coverage/run.ts excel xx-YY          # the coverage measurement, into test/excel/coverage/excel-xx-YY.json
npx tsx bench/coverage/run.ts probes <set> xx-YY   # a case set, into test/excel/cases/xx-YY/<set>.json
npx tsx bench/coverage/run.ts display <set> xx-YY  # what cells show
npx tsx bench/coverage/run.ts tables <set> xx-YY   # tables, with formulas typed and read in the locale's language
```

Excel for Mac calculates in a locale according to its launch arguments. en-US starts with `-AppleLanguages (en) -AppleLocale en_US`, and ru-RU with none, using the Russian of the Mac the results were collected on. Another tag starts with `-AppleLanguages (xx) -AppleLocale xx_YY`; that form has not yet been tried with a third language. The harness recognizes Excel's "We found a problem with some content" dialog by its English and Russian text. In another interface language, a workbook Excel rejects waits out the step's three minutes instead of being declined immediately, until the dialog's text is added to `REPAIR` and `declineRepair` in `bench/coverage/excel.ts`.

#### Collecting the results on Windows

`bench/windows/run.ts` takes the same arguments, for the package's locales only (`npx tsx bench/windows/run.ts all xx-YY`). It sets the machine to the locale's language, regional format and system locale, and checks Excel's interface language against the locale's code (`lcid`) and the system code page against its `codePage`. The machine needs the language: Excel's interface language follows the Windows display language, so a new language needs its Windows language pack, as well as Office's language pack for it, as the English one is installed there.

#### Known differences

A case where the engine knowingly differs from Excel in a locale is pinned with its reason: the test passes on the engine's result, and fails once the case starts matching Excel, until it is removed from the list.

- ru-RU's list is `RU_KNOWN` in `test/excel-ru-known.ts`, with its own reasons in `RU_REASONS`. A new locale's list goes into a separate file next to it, named in `KNOWN_BY_LOCALE` (`test/excel-locale-cases.ts`); its new reasons are added to `LOCALE_REASONS` there, and each has a line in the README's known differences.
- Against Excel for Windows, `WIN_KNOWN` (`test/excel-win-known.ts`) holds every locale's differences, with keys that start with the tag: `ru-RU:locale:format#244`.
- Displayed text and tables keep their lists by set: `MAC_SHOWN` in `test/display-excel.test.ts`, and `KNOWN` in `test/tables-excel.test.ts`.

#### Formula text and code pages

`test/locale-excel.test.ts` checks the formula text of every locale with a snapshot, `test/locale-roundtrip.test.ts` round-trips the case sets' formulas through every locale's text, and `test/codepage.test.ts` checks each locale's code page against its published table. A new code page needs that table in `test/fixtures/unicode/`; the test names the missing file.

### Accepting a locale

A locale is ready when the engine returns Excel's results in it, and every remaining difference is listed with its reason. For ru-RU that means:

- the coverage run calculated in Excel ru-RU (`npx tsx bench/coverage/run.ts excel ru`) and compared by `test/excel-locales.test.ts`;
- the case sets (`npx tsx bench/coverage/run.ts probes <set> ru`), displayed text (`display <set> ru`) and tables (`tables <set> ru`), with their tests, as described in `bench/coverage/README.md`;
- the same workbooks in Excel for Windows (`npx tsx bench/windows/run.ts all ru`, `bench/windows/README.md`);
- the differences: `RU_KNOWN` in `test/excel-ru-known.ts` pins the engine's result for each case that differs from Excel ru-RU, with a reason, and `WIN_KNOWN` does the same against Excel for Windows.

A new locale needs the same: the coverage run and the case sets calculated in its Excel, and its own list of differences ([Tests and harnesses](#tests-and-harnesses)).

## Checked against Excel

Every function's cases and the ru-RU case sets were calculated in Excel ru-RU; both translations reproduce Excel's own text of all 6,381 formulas in ru-RU, and 34,556 formulas round-trip unchanged through every locale.
