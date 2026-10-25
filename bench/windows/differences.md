# Excel for Windows against Excel for Mac and the engine

Written by `npx tsx bench/windows/compare.ts` from the formula answers in `test/excel/win/` and `test/excel/`. The class of a parted case: a platform difference (the reason in `WIN_REASONS` of `test/excel-win-cases.ts`, or the engine follows Windows) or noise.

## Cases compared

| locale | source | cases | Windows = Mac | engine = Windows | all three agree | not taken by Windows |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| en-US | measurement | 5260 | 5255 | 5256 | 5251 | 0 |
| ru-RU | measurement | 5260 | 5252 | 5256 | 5248 | 0 |
| en-US | probes | 959 | 959 | 959 | 959 | 0 |
| en-US | accrint-ironcalc-held | 42 | 42 | 42 | 42 | 0 |
| en-US | areas-each | 52 | 52 | 52 | 52 | 0 |
| en-US | areas | 1137 | 1137 | 1137 | 1137 | 0 |
| en-US | arity | 62 | 62 | 62 | 62 | 0 |
| en-US | arrays | 2428 | 2428 | 2428 | 2428 | 0 |
| en-US | binom-sumx-held | 44 | 44 | 44 | 44 | 0 |
| en-US | char-code | 305 | 298 | 305 | 298 | 0 |
| en-US | collation | 174 | 174 | 174 | 174 | 0 |
| en-US | complex-text | 156 | 156 | 156 | 156 | 0 |
| en-US | core | 258 | 258 | 258 | 258 | 0 |
| en-US | countif-ironcalc-held | 23 | 23 | 23 | 23 | 0 |
| en-US | criteria-text | 37 | 37 | 37 | 37 | 0 |
| en-US | criteria | 69 | 69 | 69 | 69 | 0 |
| en-US | currency-held | 34 | 34 | 34 | 34 | 0 |
| en-US | database-ironcalc-held | 15 | 15 | 15 | 15 | 0 |
| en-US | date-day | 93 | 93 | 93 | 93 | 0 |
| en-US | date-ironcalc-held | 33 | 33 | 33 | 33 | 0 |
| en-US | date-noyear | 191 | 191 | 191 | 191 | 0 |
| en-US | euro-ironcalc-held-2 | 13 | 13 | 13 | 13 | 0 |
| en-US | euro-ironcalc-held-3 | 16 | 16 | 16 | 16 | 0 |
| en-US | euro-ironcalc-held | 29 | 29 | 29 | 29 | 0 |
| en-US | finance-ironcalc-held-2 | 17 | 17 | 17 | 17 | 0 |
| en-US | finance-ironcalc-held | 32 | 32 | 32 | 32 | 0 |
| en-US | findate-edges | 72 | 72 | 72 | 72 | 0 |
| en-US | findate | 1615 | 1614 | 1611 | 1611 | 0 |
| en-US | fit-offset-held | 43 | 43 | 36 | 36 | 0 |
| en-US | format-scale | 14 | 14 | 14 | 14 | 0 |
| en-US | fuzz-rules | 302 | 302 | 299 | 299 | 0 |
| en-US | fuzz | 4000 | 3999 | 3995 | 3994 | 0 |
| en-US | intersect | 44 | 44 | 44 | 44 | 0 |
| en-US | known-held | 37 | 37 | 37 | 37 | 0 |
| en-US | lead-spaces | 91 | 91 | 91 | 91 | 0 |
| en-US | let-held | 33 | 33 | 33 | 33 | 0 |
| en-US | let | 82 | 82 | 81 | 81 | 0 |
| en-US | logical-text | 51 | 51 | 51 | 51 | 0 |
| en-US | lookup-held | 19 | 19 | 19 | 19 | 0 |
| en-US | lookup-own-held | 9 | 9 | 5 | 5 | 0 |
| en-US | lookup-rules | 14 | 14 | 14 | 14 | 0 |
| en-US | low-values-held | 32 | 32 | 29 | 29 | 0 |
| en-US | math-ironcalc-held | 30 | 30 | 30 | 30 | 0 |
| en-US | matheng | 3163 | 3151 | 3151 | 3151 | 0 |
| en-US | name-case-held | 23 | 16 | 22 | 16 | 1 |
| en-US | near-integer | 4220 | 4220 | 4219 | 4219 | 0 |
| en-US | near-zero-held | 189 | 189 | 189 | 189 | 0 |
| en-US | omit-held | 55 | 55 | 55 | 55 | 0 |
| en-US | omit | 4139 | 4139 | 4139 | 4139 | 0 |
| en-US | pct-text | 76 | 76 | 76 | 76 | 0 |
| en-US | power-bound-held | 7 | 7 | 7 | 7 | 0 |
| en-US | power-edges-held | 28 | 28 | 28 | 28 | 0 |
| en-US | power-neg-held | 56 | 56 | 56 | 56 | 0 |
| en-US | ref-deps | 11 | 11 | 11 | 11 | 0 |
| en-US | refs-ironcalc-held-2 | 10 | 10 | 10 | 10 | 0 |
| en-US | refs-ironcalc-held | 35 | 35 | 35 | 35 | 0 |
| en-US | rounding-ironcalc-held | 32 | 32 | 32 | 32 | 0 |
| en-US | seconds-held | 420 | 420 | 420 | 420 | 0 |
| en-US | sheet-cells | 13 | 13 | 13 | 13 | 0 |
| en-US | single-values | 66 | 66 | 66 | 66 | 0 |
| en-US | sortby-one-held | 19 | 19 | 19 | 19 | 0 |
| en-US | stat-edges-held | 111 | 111 | 72 | 72 | 0 |
| en-US | stat | 2678 | 2676 | 2663 | 2663 | 0 |
| en-US | subtotal | 503 | 503 | 503 | 503 | 0 |
| en-US | sum-zero-held | 34 | 34 | 34 | 34 | 0 |
| en-US | sumproduct-held | 17 | 17 | 17 | 17 | 0 |
| en-US | sumx-cell-held | 34 | 34 | 34 | 34 | 0 |
| en-US | tests-cell-held | 17 | 17 | 17 | 17 | 0 |
| en-US | text-date-ironcalc-held-2 | 12 | 12 | 12 | 12 | 0 |
| en-US | text-date-ironcalc-held-3 | 12 | 12 | 12 | 12 | 0 |
| en-US | text-date-ironcalc-held-4 | 27 | 27 | 27 | 27 | 0 |
| en-US | text-date-ironcalc-held | 19 | 19 | 19 | 19 | 0 |
| en-US | text-format-more | 22 | 22 | 22 | 22 | 0 |
| en-US | text-format | 479 | 474 | 479 | 474 | 0 |
| en-US | text | 1229 | 1229 | 1229 | 1229 | 0 |
| en-US | textafter-ironcalc-held-2 | 10 | 10 | 10 | 10 | 0 |
| en-US | textafter-ironcalc-held | 23 | 23 | 23 | 23 | 0 |
| en-US | time-text | 286 | 286 | 286 | 286 | 0 |
| en-US | xlookup-ironcalc-held-2 | 30 | 30 | 30 | 30 | 0 |
| en-US | xlookup-ironcalc-held | 39 | 39 | 39 | 39 | 0 |
| en-US | xmatch-regex-held-2 | 9 | 0 | 0 | 0 | 0 |
| en-US | xmatch-regex-held-3 | 6 | 0 | 0 | 0 | 0 |
| en-US | xmatch-regex-held | 22 | 3 | 3 | 3 | 0 |
| ru-RU | char-code | 305 | 298 | 305 | 298 | 0 |
| ru-RU | collation | 174 | 84 | 174 | 84 | 0 |
| ru-RU | criteria-text | 37 | 37 | 37 | 37 | 0 |
| ru-RU | lead-spaces | 91 | 91 | 91 | 91 | 0 |
| ru-RU | locale | 1109 | 1061 | 1108 | 1060 | 0 |
| ru-RU | logical-text | 51 | 51 | 51 | 51 | 0 |
| ru-RU | pct-text | 76 | 76 | 76 | 76 | 0 |
| ru-RU | text-format-ru-more | 16 | 13 | 15 | 13 | 0 |
| ru-RU | text-format-ru | 254 | 249 | 254 | 249 | 0 |

## Windows and the Mac part

| locale | source | case | function | formula | Mac | Windows | engine | engine follows | class |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| en-US | measurement | CHAR#3 | CHAR | `CHAR(200)` | "»" | "È" | "È" | Windows | platform: the engine follows Windows |
| en-US | measurement | CODE#3 | CODE | `CODE("Ж")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | measurement | CODE#5 | CODE | `CODE("Привет Мир")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | measurement | CODE#6 | CODE | `CODE("пРИВЕТ мИР")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | measurement | CODE#12 | CODE | `CODE(D!A9)` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | measurement | CHAR#3 | CHAR | `CHAR(200)` | "»" | "И" | "И" | Windows | platform: the engine follows Windows |
| ru-RU | measurement | CODE#3 | CODE | `CODE("Ж")` | 134 | 198 | 198 | Windows | platform: the engine follows Windows |
| ru-RU | measurement | CODE#5 | CODE | `CODE("Привет Мир")` | 143 | 207 | 207 | Windows | platform: the engine follows Windows |
| ru-RU | measurement | CODE#8 | CODE | `CODE(TRUE)` | 136 | 200 | 200 | Windows | platform: the engine follows Windows |
| ru-RU | measurement | CODE#12 | CODE | `CODE(D!A9)` | 143 | 207 | 207 | Windows | platform: the engine follows Windows |
| ru-RU | measurement | SORT#9 | SORT | `INDEX(_xlfn._xlws.SORT(D!F1:F7),1,1)` | "Аренда" | "x*y" | "x*y" | Windows | platform: the engine follows Windows |
| ru-RU | measurement | SORT#10 | SORT | `INDEX(_xlfn._xlws.SORT(D!F1:F7),2,1)` | "аренда" | "Арeнда*" | "Арeнда*" | Windows | platform: the engine follows Windows |
| ru-RU | measurement | SORT#11 | SORT | `INDEX(_xlfn._xlws.SORT(D!F1:F7),7,1)` | "x*y" | "Чай" | "Чай" | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#142 | CODE | `CODE("中")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#143 | CODE | `CODE("😀")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#144 | CODE | `CODE("ā")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#145 | CODE | `CODE("Ā")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#146 | CODE | `CODE("ŋ")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#147 | CODE | `CODE("ğ")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | char-code | CODE#148 | CODE | `CODE("ʃ")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| en-US | findate | WEEKNUM#16 | WEEKNUM | `WEEKNUM(45658,22)` | #NUM | 1 | #NUM | Mac | platform: winWeeknum |
| en-US | fuzz | 1#1263 | DAY | `DAY(WEEKNUM(40977,1E+3)<>2.5)` | 1 | #NUM | #NUM | Windows | noise (the Mac's answer does not repeat) |
| en-US | matheng | IMTAN#9 | IMTAN | `_xlfn.IMTAN(5)` | "-3.38051500624658" | "-3.38051500624659" | "-3.38051500624658" | Mac | platform: winMath |
| en-US | matheng | IMCOS#16 | IMCOS | `IMCOS("134217727")` | "-0.293388403555673" | "-0.293388403557722" | "-0.293388403555673" | Mac | platform: winMath |
| en-US | matheng | IMSIN#16 | IMSIN | `IMSIN("134217727")` | "-0.955993328773299" | "-0.955993328772671" | "-0.955993328773299" | Mac | platform: winMath |
| en-US | matheng | IMTAN#14 | IMTAN | `_xlfn.IMTAN("1.5707963267949")` | "-286411383293069" | "-286413440886091" | "-286411383293069" | Mac | platform: winMath |
| en-US | matheng | IMPOWER#10 | IMPOWER | `IMPOWER("1+i",100)` | "-1125899906842640-0.5529257073122i" | "-1125899906842640-0.551513671875006i" | "-1125899906842640-0.5529257073122i" | Mac | platform: winMath |
| en-US | matheng | IMPOWER#11 | IMPOWER | `IMPOWER("1+i",2000)` | "1.07150860718651E+301-1.72200351302201E+288i" | "1.07150860718651E+301-1.72227227725503E+288i" | "1.07150860718651E+301-1.72200351302201E+288i" | Mac | platform: winMath |
| en-US | matheng | TAN#8 | TAN | `TAN(PI()/2)` | 16331239353195370 | 16324552277619072 | 16331239353195400 | Mac | platform: winMath |
| en-US | matheng | SEC#8 | SEC | `_xlfn.SEC(PI()/2)` | 16331239353195370 | 16324552277619072 | 16331239353195400 | Mac | platform: winMath |
| en-US | matheng | CSC#7 | CSC | `_xlfn.CSC(PI())` | 8165619676597685 | 8162276138809536 | 8165619676597690 | Mac | platform: winMath |
| en-US | matheng | CSC#9 | CSC | `_xlfn.CSC(-PI())` | -8165619676597685 | -8162276138809536 | -8165619676597690 | Mac | platform: winMath |
| en-US | matheng | COT#7 | COT | `_xlfn.COT(PI())` | -8165619676597685 | -8162276138809536 | -8165619676597690 | Mac | platform: winMath |
| en-US | matheng | COT#9 | COT | `_xlfn.COT(-PI())` | 8165619676597685 | 8162276138809536 | 8165619676597690 | Mac | platform: winMath |
| en-US | name-case-held | bool#1 | IF | `falſe` | false | #NAME | #NAME | Windows | platform: the engine follows Windows |
| en-US | name-case-held | bool#2 | IF | `FALſE` | false | #NAME | #NAME | Windows | platform: the engine follows Windows |
| en-US | name-case-held | bool#3 | IF | `TYPE(falſe)` | 4 | 16 | #NAME | neither | platform: winNameValue |
| en-US | name-case-held | bool#4 | IF | `IF(falſe,1,2)` | 2 | #NAME | #NAME | Windows | platform: the engine follows Windows |
| en-US | name-case-held | bool#6 | IF | `NOT(falſe)` | true | #NAME | #NAME | Windows | platform: the engine follows Windows |
| en-US | name-case-held | let#2 | LET | `_xlfn.LET(_xlpm.x,1,falſe)` | false | #NAME | #NAME | Windows | platform: the engine follows Windows |
| en-US | name-case-held | text#1 | IF | `IF("falſe",1,2)` | 2 | #VALUE | #VALUE | Windows | platform: the engine follows Windows |
| en-US | stat | WEIBULL#28 | WEIBULL | `WEIBULL(5,2,1E-300,FALSE)` | #DIV0 | #NUM | #DIV0 | Mac | platform: winWeibull |
| en-US | stat | WEIBULL.DIST#28 | WEIBULL.DIST | `_xlfn.WEIBULL.DIST(5,2,1E-300,FALSE)` | #DIV0 | #NUM | #DIV0 | Mac | platform: winWeibull |
| en-US | text-format | date#78 | TEXT | `TEXT(45688.75,"[$-419]mmmm")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| en-US | text-format | date#79 | TEXT | `TEXT(0,"[$-419]mmmm")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| en-US | text-format | date#82 | TEXT | `TEXT(45688.999999,"[$-419]mmmm")` | "февраль" | "Февраль" | "Февраль" | Windows | platform: the engine follows Windows |
| en-US | text-format | date#83 | TEXT | `TEXT(0.75/86400,"[$-419]mmmm")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| en-US | text-format | date#84 | TEXT | `TEXT(1.5,"[$-419]mmmm")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| en-US | xmatch-regex-held-2 | XMATCH#1 | XMATCH | `_xlfn.XMATCH("1\.5",Ry!A1:A7,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#2 | XMATCH | `_xlfn.XMATCH("E\+20",Ry!A1:A7,3)` | 2 | #VALUE | 2 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#3 | XMATCH | `_xlfn.XMATCH("^-",Ry!A1:A7,3)` | 3 | #VALUE | 3 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#4 | XMATCH | `_xlfn.XMATCH("TRUE",Ry!A1:A7,3)` | 4 | #VALUE | 4 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#5 | XMATCH | `_xlfn.XMATCH("N/A",Ry!A1:A7,3)` | #NA | #VALUE | #NA | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#6 | XMATCH | `_xlfn.XMATCH("^$",Ry!A1:A7,3)` | 5 | #VALUE | 5 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#7 | XMATCH | `_xlfn.XMATCH("d$",Ry!A1:A7,3)` | 7 | #VALUE | 7 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#8 | XMATCH | `_xlfn.XMATCH(1.5,Ry!A1:A7,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-2 | XMATCH#9 | XMATCH | `_xlfn.XMATCH(TRUE,Ry!A1:A7,3)` | 4 | #VALUE | 4 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-3 | XMATCH#1 | XMATCH | `_xlfn.XMATCH("^$",Rz!A1:A3,3)` | 2 | #VALUE | 2 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-3 | XMATCH#2 | XMATCH | `_xlfn.XMATCH("^$",Rz!A1:A3,3,-1)` | 2 | #VALUE | 2 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-3 | XMATCH#3 | XMATCH | `_xlfn.XMATCH("^$",Rz!A:A,3,-1)` | 1048576 | #VALUE | 1048576 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-3 | XMATCH#4 | XMATCH | `_xlfn.XMATCH("^$",Rz!A3:A10,3)` | 2 | #VALUE | 2 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-3 | XMATCH#5 | XMATCH | `_xlfn.XMATCH("x*",Rz!A1:A3,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held-3 | XMATCH#6 | XMATCH | `_xlfn.XMATCH(".",Rz!A2:A2,3)` | #NA | #VALUE | #NA | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#1 | XMATCH | `_xlfn.XMATCH("[A-Z][a-z]+y",Rx!A1:A8,3,1)` | 4 | #VALUE | 4 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#2 | XMATCH | `_xlfn.XMATCH("[A-Z][a-z]+y",Rx!A1:A8,3,-1)` | 8 | #VALUE | 8 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | INDEX#1 | INDEX | `INDEX(Rx!A1:A8,_xlfn.XMATCH("[A-Z][a-z]+y",Rx!A1:A8,3,1))` | "Germany" | #VALUE | "Germany" | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#3 | XMATCH | `_xlfn.XMATCH(42,Rx!A1:A8,3)` | #NA | #VALUE | #NA | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#4 | XMATCH | `_xlfn.XMATCH("an",Rx!A1:A8,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#5 | XMATCH | `_xlfn.XMATCH("^In",Rx!A1:A8,3)` | 2 | #VALUE | 2 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#6 | XMATCH | `_xlfn.XMATCH("ia$",Rx!A1:A8,3,-1)` | 3 | #VALUE | 3 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#7 | XMATCH | `_xlfn.XMATCH("y$",Rx!A1:A8,3)` | 4 | #VALUE | 4 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#8 | XMATCH | `_xlfn.XMATCH("I.a",Rx!A1:A8,3)` | 7 | #VALUE | 7 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#9 | XMATCH | `_xlfn.XMATCH("Ind*",Rx!A1:A8,3)` | 2 | #VALUE | 2 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#10 | XMATCH | `_xlfn.XMATCH("",Rx!A1:A8,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#11 | XMATCH | `_xlfn.XMATCH("germany",Rx!A1:A12,3)` | #NA | #VALUE | #NA | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#12 | XMATCH | `_xlfn.XMATCH("(?i)germany",Rx!A1:A12,3)` | 4 | #VALUE | 4 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#13 | XMATCH | `_xlfn.XMATCH("GERMANY",Rx!A1:A12,3)` | 12 | #VALUE | 12 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#15 | XMATCH | `_xlfn.XMATCH(42,Rx!A10:A11,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XMATCH#16 | XMATCH | `_xlfn.XMATCH("\d",Rx!A10:A12,3)` | 1 | #VALUE | 1 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XLOOKUP#1 | XLOOKUP | `_xlfn.XLOOKUP("^Sp",Rx!A1:A8,Rx!B1:B8,,3)` | 5 | #VALUE | 5 | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XLOOKUP#2 | XLOOKUP | `_xlfn.XLOOKUP("zzz",Rx!A1:A8,Rx!B1:B8,"none",3)` | "none" | #VALUE | "none" | Mac | platform: winRegexMatch |
| en-US | xmatch-regex-held | XLOOKUP#3 | XLOOKUP | `_xlfn.XLOOKUP("a$",Rx!A1:A8,Rx!B1:B8,,3,-1)` | 6 | #VALUE | 6 | Mac | platform: winRegexMatch |
| ru-RU | char-code | CODE#142 | CODE | `CODE("中")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | char-code | CODE#143 | CODE | `CODE("😀")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | char-code | CODE#144 | CODE | `CODE("ā")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | char-code | CODE#145 | CODE | `CODE("Ā")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | char-code | CODE#146 | CODE | `CODE("ŋ")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | char-code | CODE#147 | CODE | `CODE("ğ")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | char-code | CODE#148 | CODE | `CODE("ʃ")` | 95 | 63 | 63 | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#1 | compare | `"a"<"а"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#3 | compare | `"a"<"Б"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#5 | compare | `"a"<"я"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#6 | compare | `"a"<"е"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#7 | compare | `"a"<"ё"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#10 | compare | `"a"<"й"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#11 | compare | `"a"<"и"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#12 | compare | `"a"<"ь"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#13 | compare | `"a"<"ъ"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#15 | compare | `"а"<"b"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#17 | compare | `"а"<"z"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#21 | compare | `"а"<"e"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#22 | compare | `"а"<"é"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#28 | compare | `"b"<"Б"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#30 | compare | `"b"<"я"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#31 | compare | `"b"<"е"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#32 | compare | `"b"<"ё"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#35 | compare | `"b"<"й"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#36 | compare | `"b"<"и"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#37 | compare | `"b"<"ь"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#38 | compare | `"b"<"ъ"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#40 | compare | `"Б"<"z"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#44 | compare | `"Б"<"e"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#45 | compare | `"Б"<"é"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#51 | compare | `"z"<"я"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#52 | compare | `"z"<"е"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#53 | compare | `"z"<"ё"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#56 | compare | `"z"<"й"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#57 | compare | `"z"<"и"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#58 | compare | `"z"<"ь"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#59 | compare | `"z"<"ъ"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#63 | compare | `"я"<"e"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#64 | compare | `"я"<"é"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#71 | compare | `"е"<"e"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#72 | compare | `"е"<"é"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#78 | compare | `"ё"<"e"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#79 | compare | `"ё"<"é"` | true | false | false | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#86 | compare | `"e"<"й"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#87 | compare | `"e"<"и"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#88 | compare | `"e"<"ь"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#89 | compare | `"e"<"ъ"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#91 | compare | `"é"<"й"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#92 | compare | `"é"<"и"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#93 | compare | `"é"<"ь"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | compare#94 | compare | `"é"<"ъ"` | false | true | true | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#4 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),4,1)` | "а" | "A" | "A" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#5 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),5,1)` | "аб" | "ab" | "ab" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#6 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),6,1)` | "абв" | "abc" | "abc" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#7 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),7,1)` | "Абв" | "Abc" | "Abc" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#8 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),8,1)` | "Б" | "b" | "b" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#9 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),9,1)` | "е" | "e" | "e" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#10 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),10,1)` | "ё" | "é" | "é" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#11 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),11,1)` | "Ё" | "ea" | "ea" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#12 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),12,1)` | "ж" | "éa" | "éa" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#13 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),13,1)` | "и" | "x1" | "x1" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#14 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),14,1)` | "й" | "z" | "z" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#15 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),15,1)` | "х1" | "Z" | "Z" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#16 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),16,1)` | "ъ" | "а" | "а" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#17 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),17,1)` | "ь" | "аб" | "аб" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#18 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),18,1)` | "я" | "абв" | "абв" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#19 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),19,1)` | "A" | "Абв" | "Абв" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#20 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),20,1)` | "ab" | "Б" | "Б" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#21 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),21,1)` | "abc" | "е" | "е" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#22 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),22,1)` | "Abc" | "ё" | "ё" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#23 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),23,1)` | "b" | "Ё" | "Ё" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#24 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),24,1)` | "e" | "ж" | "ж" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#25 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),25,1)` | "é" | "и" | "и" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#26 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),26,1)` | "ea" | "й" | "й" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#27 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),27,1)` | "éa" | "х1" | "х1" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#28 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),28,1)` | "x1" | "ъ" | "ъ" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#29 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),29,1)` | "z" | "ь" | "ь" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#30 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30),30,1)` | "Z" | "я" | "я" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#31 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),1,1)` | "z" | "я" | "я" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#32 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),4,1)` | "éa" | "х1" | "х1" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#33 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),7,1)` | "e" | "ж" | "ж" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#34 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),10,1)` | "Abc" | "е" | "е" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#35 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),13,1)` | "я" | "Абв" | "Абв" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#36 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),16,1)` | "х1" | "z" | "z" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#37 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),19,1)` | "ж" | "éa" | "éa" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#38 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),22,1)` | "е" | "e" | "e" | Windows | platform: the engine follows Windows |
| ru-RU | collation | sort#39 | sort | `INDEX(_xlfn._xlws.SORT(O!A1:A30,1,-1),25,1)` | "Абв" | "Abc" | "Abc" | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#1 | criteria | `COUNTIF(O!A1:A30,"<м")` | 14 | 26 | 26 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#2 | criteria | `COUNTIF(O!A1:A30,"<b")` | 22 | 7 | 7 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#3 | criteria | `COUNTIF(O!A1:A30,"<я")` | 17 | 29 | 29 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#4 | criteria | `COUNTIF(O!A1:A30,">z")` | 0 | 15 | 15 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#5 | criteria | `COUNTIF(O!A1:A30,"<ё")` | 9 | 21 | 21 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#6 | criteria | `COUNTIF(O!A1:A30,"<=е")` | 9 | 21 | 21 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#7 | criteria | `COUNTIF(O!A1:A30,"<é")` | 24 | 9 | 9 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#8 | criteria | `COUNTIF(O!A1:A30,">=а")` | 27 | 15 | 15 | Windows | platform: the engine follows Windows |
| ru-RU | collation | criteria#9 | criteria | `COUNTIF(O!A1:A30,"<Б")` | 7 | 19 | 19 | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#7 | format | `TEXT(45688.75,"ММММ")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#8 | format | `TEXT(45688.75,"МММММ")` | "я" | "Я" | "Я" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#11 | format | `TEXT(45688.75,"ДД ММММ ГГГГ")` | "31 январь 2025" | "31 Январь 2025" | "31 Январь 2025" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#12 | format | `TEXT(45688.75,"ММММ ГГГГ")` | "январь 2025" | "Январь 2025" | "Январь 2025" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#267 | format | `TEXT(DATE(2025,1,15),"ММММ")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#269 | format | `TEXT(DATE(2025,1,15),"ДД ММММ")` | "15 январь" | "15 Январь" | "15 Январь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#270 | format | `TEXT(DATE(2025,1,15),"МММММ")` | "я" | "Я" | "Я" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#271 | format | `TEXT(DATE(2025,2,15),"ММММ")` | "февраль" | "Февраль" | "Февраль" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#272 | format | `TEXT(DATE(2025,2,15),"МММ")` | "февр" | "фев" | "фев" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#273 | format | `TEXT(DATE(2025,2,15),"ДД ММММ")` | "15 февраль" | "15 Февраль" | "15 Февраль" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#274 | format | `TEXT(DATE(2025,2,15),"МММММ")` | "ф" | "Ф" | "Ф" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#275 | format | `TEXT(DATE(2025,3,15),"ММММ")` | "март" | "Март" | "Март" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#276 | format | `TEXT(DATE(2025,3,15),"МММ")` | "март" | "мар" | "мар" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#277 | format | `TEXT(DATE(2025,3,15),"ДД ММММ")` | "15 март" | "15 Март" | "15 Март" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#278 | format | `TEXT(DATE(2025,3,15),"МММММ")` | "м" | "М" | "М" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#279 | format | `TEXT(DATE(2025,4,15),"ММММ")` | "апрель" | "Апрель" | "Апрель" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#281 | format | `TEXT(DATE(2025,4,15),"ДД ММММ")` | "15 апрель" | "15 Апрель" | "15 Апрель" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#282 | format | `TEXT(DATE(2025,4,15),"МММММ")` | "а" | "А" | "А" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#283 | format | `TEXT(DATE(2025,5,15),"ММММ")` | "май" | "Май" | "Май" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#285 | format | `TEXT(DATE(2025,5,15),"ДД ММММ")` | "15 май" | "15 Май" | "15 Май" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#286 | format | `TEXT(DATE(2025,5,15),"МММММ")` | "м" | "М" | "М" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#287 | format | `TEXT(DATE(2025,6,15),"ММММ")` | "июнь" | "Июнь" | "Июнь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#288 | format | `TEXT(DATE(2025,6,15),"МММ")` | "июнь" | "июн" | "июн" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#289 | format | `TEXT(DATE(2025,6,15),"ДД ММММ")` | "15 июнь" | "15 Июнь" | "15 Июнь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#290 | format | `TEXT(DATE(2025,6,15),"МММММ")` | "и" | "И" | "И" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#291 | format | `TEXT(DATE(2025,7,15),"ММММ")` | "июль" | "Июль" | "Июль" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#292 | format | `TEXT(DATE(2025,7,15),"МММ")` | "июль" | "июл" | "июл" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#293 | format | `TEXT(DATE(2025,7,15),"ДД ММММ")` | "15 июль" | "15 Июль" | "15 Июль" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#294 | format | `TEXT(DATE(2025,7,15),"МММММ")` | "и" | "И" | "И" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#295 | format | `TEXT(DATE(2025,8,15),"ММММ")` | "август" | "Август" | "Август" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#297 | format | `TEXT(DATE(2025,8,15),"ДД ММММ")` | "15 август" | "15 Август" | "15 Август" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#298 | format | `TEXT(DATE(2025,8,15),"МММММ")` | "а" | "А" | "А" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#299 | format | `TEXT(DATE(2025,9,15),"ММММ")` | "сентябрь" | "Сентябрь" | "Сентябрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#300 | format | `TEXT(DATE(2025,9,15),"МММ")` | "сент" | "сен" | "сен" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#301 | format | `TEXT(DATE(2025,9,15),"ДД ММММ")` | "15 сентябрь" | "15 Сентябрь" | "15 Сентябрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#302 | format | `TEXT(DATE(2025,9,15),"МММММ")` | "с" | "С" | "С" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#303 | format | `TEXT(DATE(2025,10,15),"ММММ")` | "октябрь" | "Октябрь" | "Октябрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#305 | format | `TEXT(DATE(2025,10,15),"ДД ММММ")` | "15 октябрь" | "15 Октябрь" | "15 Октябрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#306 | format | `TEXT(DATE(2025,10,15),"МММММ")` | "о" | "О" | "О" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#307 | format | `TEXT(DATE(2025,11,15),"ММММ")` | "ноябрь" | "Ноябрь" | "Ноябрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#308 | format | `TEXT(DATE(2025,11,15),"МММ")` | "нояб" | "ноя" | "ноя" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#309 | format | `TEXT(DATE(2025,11,15),"ДД ММММ")` | "15 ноябрь" | "15 Ноябрь" | "15 Ноябрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#310 | format | `TEXT(DATE(2025,11,15),"МММММ")` | "н" | "Н" | "Н" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#311 | format | `TEXT(DATE(2025,12,15),"ММММ")` | "декабрь" | "Декабрь" | "Декабрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#313 | format | `TEXT(DATE(2025,12,15),"ДД ММММ")` | "15 декабрь" | "15 Декабрь" | "15 Декабрь" | Windows | platform: the engine follows Windows |
| ru-RU | locale | format#314 | format | `TEXT(DATE(2025,12,15),"МММММ")` | "д" | "Д" | "Д" | Windows | platform: the engine follows Windows |
| ru-RU | locale | criteria#30 | criteria | `COUNTIF(L!A1:A16,"<м")` | 4 | 6 | 6 | Windows | platform: the engine follows Windows |
| ru-RU | locale | criteria#31 | criteria | `COUNTIF(L!A1:A16,">a")` | 2 | 3 | 3 | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru-more | TEXT#11 | TEXT | `TEXT(45688,"Д ММММ ГГГГ")` | "31 январь 2025" | "31 Январь 2025" | "31 Январь 2025" | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru-more | TEXT#12 | TEXT | `TEXT(45688,"ДД ММММ")` | "31 январь" | "31 Январь" | "31 Январь" | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru-more | TEXT#13 | TEXT | `TEXT(45688,"[$-F800]ДДДД, ММММ ДД, ГГГГ")` | "пятница, 31 января 2025 г." | "31 января 2025 г." | "пятница, 31 января 2025 г." | Mac | platform: winLongDate |
| ru-RU | text-format-ru | date#55 | TEXT | `TEXT(45688.75,"[$-419]ММММ")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru | date#56 | TEXT | `TEXT(0,"[$-419]ММММ")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru | date#58 | TEXT | `TEXT(45688.999999,"[$-419]ММММ")` | "февраль" | "Февраль" | "Февраль" | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru | date#59 | TEXT | `TEXT(0.75/86400,"[$-419]ММММ")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |
| ru-RU | text-format-ru | date#60 | TEXT | `TEXT(1.5,"[$-419]ММММ")` | "январь" | "Январь" | "Январь" | Windows | platform: the engine follows Windows |

## Windows and the Mac agree, the engine does not

| locale | source | case | formula | Excel | engine | reason (the Mac's, test/excel-cases.ts) |
| --- | --- | --- | --- | --- | --- | --- |
| en-US | measurement | CHIINV#7 | `CHIINV(0.050001,10000000000)` | #NUM | 10000232617.1965 | gammaInvConvergence |
| en-US | measurement | CHISQ.INV#14 | `_xlfn.CHISQ.INV(0.93,10000000000)` | #NUM | 10000208709.154 | gammaInvConvergence |
| en-US | measurement | CHISQ.INV.RT#15 | `_xlfn.CHISQ.INV.RT(0.050001,10000000000)` | #NUM | 10000232617.1965 | gammaInvConvergence |
| en-US | measurement | GAMMAINV#8 | `GAMMAINV(0.068094,10000000000,2)` | #NUM | 19999701973.2484 | gammaInvConvergence |
| ru-RU | measurement | CHIINV#7 | `CHIINV(0.050001,10000000000)` | #NUM | 10000232617.1965 | gammaInvConvergence |
| ru-RU | measurement | CHISQ.INV#14 | `_xlfn.CHISQ.INV(0.93,10000000000)` | #NUM | 10000208709.154 | gammaInvConvergence |
| ru-RU | measurement | CHISQ.INV.RT#15 | `_xlfn.CHISQ.INV.RT(0.050001,10000000000)` | #NUM | 10000232617.1965 | gammaInvConvergence |
| ru-RU | measurement | GAMMAINV#8 | `GAMMAINV(0.068094,10000000000,2)` | #NUM | 19999701973.2484 | gammaInvConvergence |
| en-US | findate | RATE#20 | `RATE(10,-100,1000,0,1)` | -1.487997633973174e-9 | 4.17387402722136e-11 | rateZeroRoot |
| en-US | findate | RATE#23 | `RATE(5,-200,1000,0,1)` | -1.6302285183735498e-9 | 2.1421887871288e-10 | rateZeroRoot |
| en-US | findate | RATE#24 | `RATE(10,-100,1000,0,0,0.05)` | #NUM | 1.22749297268431e-10 | rateZeroRoot |
| en-US | fit-offset-held | LINEST#1 | `INDEX(LINEST(Edge!A2:A10001,Edge!B2:B10001),1,2)-1E12` | -0.0096435546875 | 0 | linestMean |
| en-US | fit-offset-held | LINEST#4 | `INDEX(LINEST(Edge!A2:A10001,Fit!A2:B10001),1,3)-1E12` | -0.0096435546875 | 0 | linestMean |
| en-US | fit-offset-held | LINEST#11 | `INDEX(LINEST(Edge!A2:A10001,Edge!B2:B10001,TRUE,TRUE),5,1)` | 1.5977911949157715 | 0.668002774232626 | linestMean |
| en-US | fit-offset-held | TREND#1 | `TREND(Edge!A2:A10001,Edge!B2:B10001,3)-1E12` | 0.0203857421875 | 0.030029296875 | linestMean |
| en-US | fit-offset-held | TREND#2 | `TREND(Edge!A2:A10001,Fit!A2:B10001,{3,5})-1E12` | 0.0203857421875 | 0.030029296875 | linestMean |
| en-US | fit-offset-held | GROWTH#2 | `GROWTH(Fit!F2:F13,Fit!D2:D13,1E10+12)` | #NUM | 313.83972099924 | growthRange |
| en-US | fit-offset-held | GROWTH#3 | `INDEX(GROWTH(Fit!F2:F13,Fit!D2:D13),12)` | #NUM | 285.309168434641 | growthRange |
| en-US | fuzz-rules | LINEST#41 | `INDEX(LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE),2,4)` | 0.10319961748964883 | 0.474341649025257 | linestConstantColumn |
| en-US | fuzz-rules | LINEST#42 | `INDEX(LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE),4,1)` | 615.6878650646264 | 56.3333333333333 | linestConstantColumn |
| en-US | fuzz-rules | LINEST#43 | `INDEX(LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE),4,2)` | 1 | 2 | linestConstantColumn |
| en-US | fuzz | 1#234 | `_xlfn.TEXTJOIN(";",FALSE,BESSELK(10,3))` | "2.72527006358477E-05" | "2.72527006358476E-05" | textDigits |
| en-US | fuzz | 1#403 | `_xlfn.TEXTJOIN(";",FALSE,_xlfn.ERFC.PRECISE(7))` | "4.18382560777942E-23" | "4.18382560777941E-23" | textDigits |
| en-US | fuzz | 1#947 | `_xlfn.TEXTJOIN(";",FALSE,RATE(48,-200,8000,4,D!A15,0.5))` | "0.00803735952148241" | "0.00803735952120296" | textDigits |
| en-US | fuzz | 1#1327 | `_xlfn.TEXTJOIN(";",FALSE,IRR(D!I1:I5))` | "0.196309310268599" | "0.1963093102686" | textDigits |
| en-US | fuzz | 2#472 | `_xlfn.TEXTJOIN(";",FALSE,FV(0.005,10,-200,7,TRUE))` | "2048.47532711544" | "2048.47532711546" | textDigits |
| en-US | let | let:file#5 | `_xlfn.LET(_xlpm.TRUE,1,_xlpm.TRUE)` | 1 | #ERROR | letFileName |
| en-US | lookup-own-held | LOOKUP#5 | `Lv!C1` | "" | 0 | lookupReadZone |
| en-US | lookup-own-held | LOOKUP#7 | `Lq!E10` | "F" | 0 | lookupReadZone |
| en-US | lookup-own-held | LOOKUP#8 | `Lp!E10` | "F" | 0 | lookupReadZone |
| en-US | lookup-own-held | LOOKUP#9 | `Lu!F2` | "c5" | 0 | lookupReadZone |
| en-US | low-values-held | unique-text#8 | `V!M3=V!M4` | false | true | textEquivalence |
| en-US | low-values-held | unique-text#9 | `MATCH(V!M2,V!M1:M2,0)` | 2 | 1 | textEquivalence |
| en-US | low-values-held | unique-text#10 | `COUNTIF(V!M1:M2,V!M1)` | 2 | 1 | textEquivalence |
| en-US | near-integer | literal#7 | `(1-FTEST(D!C1:C8,D!C1:C8))*2^52` | 1.5 | 1 | lastBits |
| en-US | stat-edges-held | VAR.P#1 | `_xlfn.VAR.P(Edge!A2:A10001)` | 0.00015977911949157714 | 0.0000668002774232626 | momentRounding |
| en-US | stat-edges-held | VAR.P#2 | `_xlfn.VAR.P(Edge!A:A)` | 0.00015977911949157714 | 0.0000668002774232626 | momentRounding |
| en-US | stat-edges-held | VAR.S#1 | `_xlfn.VAR.S(Edge!A2:A10001)` | 0.0001597950990014773 | 0.0000668069581190745 | momentRounding |
| en-US | stat-edges-held | STDEV.P#1 | `_xlfn.STDEV.P(Edge!A2:A10001)` | 0.012640376556557845 | 0.00817314366833611 | momentRounding |
| en-US | stat-edges-held | STDEV.P#2 | `_xlfn.STDEV.P(Edge!A:A)` | 0.012640376556557845 | 0.00817314366833611 | momentRounding |
| en-US | stat-edges-held | STDEV.S#1 | `_xlfn.STDEV.S(Edge!A2:A10001)` | 0.012641008622791034 | 0.00817355235617137 | momentRounding |
| en-US | stat-edges-held | DEVSQ#1 | `DEVSQ(Edge!A2:A10001)` | 1.5977911949157715 | 0.668002774232626 | momentRounding |
| en-US | stat-edges-held | DEVSQ#2 | `DEVSQ(Edge!A:A)` | 1.5977911949157715 | 0.668002774232626 | momentRounding |
| en-US | stat-edges-held | CORREL#1 | `CORREL(Edge!A2:A10001,Edge!B2:B10001)` | 0.6465902049489168 | 1 | momentRounding |
| en-US | stat-edges-held | PEARSON#1 | `PEARSON(Edge!A2:A10001,Edge!B2:B10001)` | 0.6465902049489168 | 1 | momentRounding |
| en-US | stat-edges-held | RSQ#1 | `RSQ(Edge!A2:A10001,Edge!B2:B10001)` | 0.4180788931358824 | 1 | momentRounding |
| en-US | stat-edges-held | SLOPE#2 | `SLOPE(Edge!B2:B10001,Edge!A2:A10001)` | 41.76710112889263 | 99.9024390243902 | momentRounding |
| en-US | stat-edges-held | STEYX#1 | `STEYX(Edge!A2:A10001,Edge!B2:B10001)` | 0.009643518110970913 | 0 | momentRounding |
| en-US | stat-edges-held | DVAR#1 | `DVAR(Edge!A1:A10001,1,Edge!Q1:Q2)` | 0.0001597950990014773 | 0.0000668069581190745 | momentRounding |
| en-US | stat-edges-held | DVARP#1 | `DVARP(Edge!A1:A10001,1,Edge!Q1:Q2)` | 0.00015977911949157714 | 0.0000668002774232626 | momentRounding |
| en-US | stat-edges-held | DSTDEV#1 | `DSTDEV(Edge!A1:A10001,1,Edge!Q1:Q2)` | 0.012641008622791034 | 0.00817355235617137 | momentRounding |
| en-US | stat-edges-held | DSTDEVP#1 | `DSTDEVP(Edge!A1:A10001,1,Edge!Q1:Q2)` | 0.012640376556557845 | 0.00817314366833611 | momentRounding |
| en-US | stat-edges-held | CORREL#3 | `CORREL(Edge!D2:D3,Edge!E2:E3)` | #DIV0 | 1 | momentRange |
| en-US | stat-edges-held | PEARSON#3 | `PEARSON(Edge!D2:D3,Edge!E2:E3)` | #DIV0 | 1 | momentRange |
| en-US | stat-edges-held | RSQ#3 | `RSQ(Edge!D2:D3,Edge!E2:E3)` | #DIV0 | 1 | momentRange |
| en-US | stat-edges-held | SLOPE#6 | `SLOPE(Edge!E2:E3,Edge!D2:D3)` | #DIV0 | 1e+200 | momentRange |
| en-US | stat-edges-held | AVERAGE#6 | `AVERAGE(Edge!G2:G4)` | #NUM | 1e+308 | momentRange |
| en-US | stat-edges-held | VAR.P#6 | `_xlfn.VAR.P(Edge!G2:G4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | VAR.S#5 | `_xlfn.VAR.S(Edge!G2:G4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | STDEV.P#6 | `_xlfn.STDEV.P(Edge!G2:G4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | STDEV.S#5 | `_xlfn.STDEV.S(Edge!G2:G4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | DEVSQ#6 | `DEVSQ(Edge!G2:G4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | DAVERAGE#3 | `DAVERAGE(Edge!G1:G4,1,Edge!T1:T2)` | #NUM | 1e+308 | momentRange |
| en-US | stat-edges-held | DVAR#3 | `DVAR(Edge!G1:G4,1,Edge!T1:T2)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | COVARIANCE.P#2 | `_xlfn.COVARIANCE.P(Edge!H2:H4,Edge!I2:I4)` | #NUM | 2.22222222222222e+107 | momentRange |
| en-US | stat-edges-held | COVARIANCE.S#4 | `_xlfn.COVARIANCE.S(Edge!H2:H4,Edge!I2:I4)` | #NUM | 3.33333333333333e+107 | momentRange |
| en-US | stat-edges-held | CORREL#4 | `CORREL(Edge!H2:H4,Edge!I2:I4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | PEARSON#4 | `PEARSON(Edge!H2:H4,Edge!I2:I4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | RSQ#4 | `RSQ(Edge!H2:H4,Edge!I2:I4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | SLOPE#7 | `SLOPE(Edge!I2:I4,Edge!H2:H4)` | #NUM | 0 | momentRange |
| en-US | stat-edges-held | INTERCEPT#2 | `INTERCEPT(Edge!I2:I4,Edge!H2:H4)` | #NUM | 3.33333333333333e+307 | momentRange |
| en-US | stat-edges-held | FORECAST.LINEAR#4 | `_xlfn.FORECAST.LINEAR(1,Edge!I2:I4,Edge!H2:H4)` | #NUM | 3.33333333333333e+307 | momentRange |
| en-US | stat-edges-held | STEYX#2 | `STEYX(Edge!K2:K4,Edge!J2:J4)` | 0 | 8.16496575965492e-9 | momentRounding |
| en-US | stat-edges-held | STEYX#3 | `STEYX(Edge!M2:M4,Edge!L2:L4)` | 0 | 8.16496575965492e-9 | momentRounding |
| en-US | stat | GAMMA.INV#5 | `_xlfn.GAMMA.INV(0.999,0.01,1)` | #NUM | 1.50908414769475 | gammaInvConvergence |
| en-US | stat | GAMMA.INV#16 | `_xlfn.GAMMA.INV(0.7,9.99E+9,1)` | #NUM | 9990052413.58303 | gammaInvConvergence |
| en-US | stat | GAMMA.INV#17 | `_xlfn.GAMMA.INV(0.7,1E+10,1)` | #NUM | 10000052439.8096 | gammaInvConvergence |
| en-US | stat | GAMMAINV#5 | `GAMMAINV(0.999,0.01,1)` | #NUM | 1.50908414769475 | gammaInvConvergence |
| en-US | stat | GAMMAINV#16 | `GAMMAINV(0.7,9.99E+9,1)` | #NUM | 9990052413.58303 | gammaInvConvergence |
| en-US | stat | GAMMAINV#17 | `GAMMAINV(0.7,1E+10,1)` | #NUM | 10000052439.8096 | gammaInvConvergence |
| en-US | stat | BINOM.INV#30 | `_xlfn.BINOM.INV(3,0.25,0.421875)` | 0 | 1 | binomInvTie |
| en-US | stat | GAMMA.INV#26 | `_xlfn.GAMMA.INV(0.99,0.01,1)` | #NUM | 0.265052550251588 | gammaInvConvergence |
| en-US | stat | GAMMA.INV#27 | `_xlfn.GAMMA.INV(0.995,0.02,1)` | #NUM | 0.934242187465308 | gammaInvConvergence |
| en-US | stat | GAMMA.INV#31 | `_xlfn.GAMMA.INV(0.95,5E+9,1)` | #NUM | 5000116309.28388 | gammaInvConvergence |
| en-US | stat | GAMMA.INV#33 | `_xlfn.GAMMA.INV(0.3,6E+9,1)` | #NUM | 5999959379.86929 | gammaInvConvergence |
| en-US | stat | CHISQ.INV#27 | `_xlfn.CHISQ.INV(0.95,1E+10)` | #NUM | 10000232618.5678 | gammaInvConvergence |
| en-US | stat | CHISQ.INV.RT#27 | `_xlfn.CHISQ.INV.RT(0.05,9.99E+9)` | #NUM | 9990232502.22995 | gammaInvConvergence |
| ru-RU | locale | format#244 | `TEXT(-0.5,"#,##0,00")` | "-,50,00" | "-,500,00" | secondDecimal |
