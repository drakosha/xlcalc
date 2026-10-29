# Regular expressions in XLOOKUP and XMATCH

Match mode 3 of `XLOOKUP` and `XMATCH` finds a regular expression anywhere in a value's text: `XMATCH("[A-Z][a-z]+y",A2:A9,3)` is the position of the first value with a capital letter followed by lowercase letters and a `y`. Numbers are matched by their text, a blank as empty text, from the end with search mode −1; a binary search (search modes 2 and −2) is `#VALUE!`. Excel for Mac reads the pattern as PCRE2 does, with Unicode classes; Excel for Windows 16.0.20430 has no match mode 3 and returns `#VALUE!`. The engine follows Excel for Mac.

How closely a pattern is read as Excel reads it is the host's choice, between three levels.

## The three levels

| | Built in | `pcreSyntax` | `pcreEngine` |
|---|---|---|---|
| What runs | JavaScript's regular expressions | JavaScript's regular expressions, written from the engine's parse of the pattern | the package's own PCRE2 engine |
| Syntax | JavaScript's, with Unicode `\d \s \w \b`, `\A \z \Z`, `{,m}` and a leading `(?i)` as PCRE2 reads them | Excel's, except recursion, conditions on groups, the verbs other than `(*FAIL)` and `(*MARK)`, `\X`, `\p{Bidi_Class:…}` and `(*NOTEMPTY)` (`#VALUE!`) | Excel's, except `\p{Bidi_Class:…}` (`#VALUE!`) |
| Excel's match limit (`#VALUE!`) | no | no | yes |
| A heavy pattern stopped by `timeoutMs` | no | no | yes |
| Speed against the built-in level | 1 | 1.3 to 5.5 times slower | 1.5 to 17 times slower |
| Bundle, beyond `compute` | 1.4 KB of it | 19.9 KB (7.2 KB gzipped) | 31.2 KB (11.5 KB gzipped); both levels together 34.0 KB (12.5 KB) |

What the levels answer where they differ:

| Formula | Excel | Built in | `pcreSyntax` | `pcreEngine` |
|---|---|---|---|---|
| `XMATCH("^\w+$",A1,3)` over `привет` | 1 | 1 | 1 | 1 |
| `XMATCH("a(?i)BC",A1,3)` over `abc` | 1 | `#VALUE!` | 1 | 1 |
| `XMATCH("(?>a+)a",A1,3)` over `aaa` | `#N/A` | `#VALUE!` | `#N/A` | `#N/A` |
| `XMATCH("b$",A1,3)` over `ab` and a line feed | 1 | `#N/A` | 1 | 1 |
| `XMATCH("^(\((?:[^()]\|(?1))*\))$",A1,3)` over `(a(b)c)` | 1 | `#VALUE!` | `#VALUE!` | 1 |
| `XMATCH("(a+)+$",A1,3)` over 19 letters `a` and a `b` | `#VALUE!` | `#N/A` | `#N/A` | `#VALUE!` |
| `XMATCH("(a+)+$",A1,3)` over 40 letters `a` and a `b` | `#VALUE!` | runs for days | runs for days | `#VALUE!` at once |

Without a level, every other construct is read as JavaScript reads it: a possessive quantifier (`a++`), an atomic group, `(?i)` past the start, `\Q…\E` and POSIX classes are `#VALUE!` or mean something else, and `$` does not match before a final newline. `pcreSyntax` translates them, including `(?x)`, `(?s)`, `(?m)`, `(?n)`, `(?U)`, named groups in every spelling, `\p{…}` by PCRE2's names, `\h \v \R \N`, a condition on an assertion, and lines ended by CR, LF and CRLF as in Excel; it accepts and refuses patterns exactly as the engine does, whose parser it shares. A back reference to a group that took no part in the match matches empty text there, where PCRE2 fails it; a back reference to a number or name that several groups carry (in `(?|…)` or with `(?J)`) needs duplicate group names, which JavaScript has from Node 23, Chrome 125, Firefox 129 and Safari 17 (`#VALUE!` before).

Which one to choose:

- The built-in level, for ordinary workbooks, whose patterns are literals, classes, anchors and alternatives.
- `pcreSyntax`, when patterns are written in Excel's syntax and come from trusted workbooks: it reads them as Excel does at JavaScript's speed.
- `pcreEngine`, when the people using the application write the patterns, and a pattern must not hold a tab or a thread: Excel's match limit ends a heavy pattern by itself, and `timeoutMs` ends the rest.

## Loading a level

```js
import { registerRegex } from 'xlcalc'
import { pcreEngine } from 'xlcalc/regex'

registerRegex(pcreEngine) // or pcreSyntax
```

The same lines serve Node and a browser bundle; the bundler takes `xlcalc/regex` with them. A level applies to the calculations from then on, the last one registered being used; values already calculated stay until they are calculated again. A pool's threads calculate with the level registered where the pool runs, before or after `createCalcPool`: a thread imports the level's module itself (a separate chunk for a bundler) before its next task, as it does for locales ([Locales](async.md#locales)).

## Stopping a heavy pattern

A pattern that backtracks a lot can take time that grows exponentially with the text: `(a+)+$` over 24 letters `a` and a `b` takes 0.15 seconds in JavaScript, over 26 letters 6 seconds, over 40 letters days. Excel ends such a search at its match limit and returns `#VALUE!`. With `pcreEngine`, the engine does the same: `(a+)+$` over 40 letters is `#VALUE!` after about 50 ms. The limit is counted per start position in the text, as PCRE2 counts it, on the engine's own steps, so near the limit a pattern can stop where Excel's does not, or the other way round: `(?:a+)+$` over 19 letters is `#N/A` in Excel and `#VALUE!` here.

Patterns under the limit can still take long over many values: 2,000 values of 18 letters take about two minutes. With `pcreEngine`, `timeoutMs` stops them in a session on the same thread (`createInlineCalc`) and in a pool's thread, with `soft_timeout` within a few milliseconds of the deadline, and the session and the thread go on. A signal is seen before the call and between formulas, not inside one: one formula that runs long sees an `abort()` when it ends ([Deadlines and cancellation](async.md#deadlines-and-cancellation)). A synchronous `compute` or `createSession` has no deadline; Excel's match limit bounds each value there. Without `pcreEngine`, JavaScript's regular expression is one call that nothing interrupts: only a pool's `hardTimeoutMs` ends it, by terminating the thread.

## Speed

`XMATCH` over 10,000 values whose last one matches, milliseconds per formula in a session, the median of 15 rounds in one process:

| Pattern | Built in | `pcreSyntax` | `pcreEngine` |
|---|---|---|---|
| `^abc` | 0.08 | 0.13 | 0.12 |
| `[0-9]{3}-\d+` | 0.10 | 0.54 | 1.5 |
| `(foo\|bar\|baz)$` | 0.09 | 0.14 | 1.4 |
| `(?i)ABC$` | 0.08 | 0.11 | 0.38 |
| `^\w+$` over Cyrillic words | 0.52 | 0.91 | 1.6 |
| `^\d++-` | `#VALUE!` | 0.24 | 0.10 |
| `^(?=.*\d)[a-z]+\d*$` | 0.20 | 0.48 | 2.5 |
| `(a+)+$` over 10 values of 18 letters | 21 | 32 | 350 |

The first use of a pattern compiles it: 3 µs built in for one in JavaScript's own syntax, 35 µs when it has `\d`, `\s` or `\w`, whose Unicode classes JavaScript takes longer to compile, 43 µs with `pcreSyntax`, 6 µs with `pcreEngine`; the 64 patterns used last stay compiled.

## Checked against Excel

Excel for Mac 16.114 calculated 129 patterns with every kind of construct (`xmatch-regex-syntax`) and 68 on what that set left open (`xmatch-regex-syntax-2`): the newlines, `\w` over marks and connectors, caseless letters, the options of newer PCRE2 releases and where the match limit lies. `pcreEngine` returns Excel's answer on all of them but `\p{Bidi_Class:L}`, and on 63 of the 64 cases of two sets predicted before Excel ran them (`xmatch-regex-own-held`, `xmatch-regex-own-held-2`): the match limit for twelve other shapes of heavy patterns, a literal the match requires (missing, `#N/A` at once, in texts of up to 32,000 characters), caseless letters outside ASCII, `\X`, the verbs, recursion. Of the 260 cases that JavaScript's regular expressions calculate in reasonable time (the heavy ones of the predicted sets would take hours), `pcreSyntax` differs from Excel on 41 and the built-in level on 123; all of them are listed in `test/regex-levels.test.ts`. [Development](development.md#regular-expressions) has the details.
