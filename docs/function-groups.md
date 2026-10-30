# Groups of functions

The package root calculates every function it has. A host that bundles the engine for a browser can take a lighter
entry, `xlcalc/core`: the same API without five groups of rarely used functions, which it loads only if its workbooks
need them. `compute` from `xlcalc/core` is 320.8 KB minified, from the root 381.7 KB.

```js
import { compute, registerFunctions } from 'xlcalc/core'
import { engineering, finance } from 'xlcalc/functions'

registerFunctions(engineering, finance)
```

`registerFunctions(...allFunctions)` loads all of them, and then `xlcalc/core` calculates what the root does. A bundle
keeps only the groups it imports. The root needs none of this: its values and hashes are the same as before the groups
existed.

## The groups

| Group | Functions | Minified, beyond `compute` |
|---|---|---|
| `engineering` | Bessel functions (`BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`), complex numbers (`COMPLEX` and the 25 `IM…`), `CONVERT`, number systems (`BIN2DEC` and the other 11), bits (`BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`), `ERF`, `ERFC` and their `.PRECISE` forms, `DELTA`, `GESTEP`: 54 | 16.7 KB |
| `finance` | securities (`PRICE`, `YIELD`, `DURATION`, `MDURATION`, `ACCRINT`, `ACCRINTM`), coupons (`COUPDAYBS` and the other 5), odd periods (`ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`), discounted securities (`DISC`, `INTRATE`, `PRICEDISC`, `PRICEMAT`, `RECEIVED`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `YIELDDISC`, `YIELDMAT`), depreciation (`DB`, `DDB`, `SLN`, `SYD`, `VDB`, `AMORDEGRC`, `AMORLINC`): 33 | 9.4 KB |
| `lambda` | `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`, `GROUPBY`, `PIVOTBY`: 8 | 8.5 KB |
| `statistics` | the distributions and their inverses (`NORM.DIST`, `NORM.S.INV`, `T.DIST`, `BINOM.DIST`, `GAMMA`, `PERMUT` and their kin), the tests (`T.TEST`, `CHISQ.TEST`, `F.TEST`, `Z.TEST`), `LINEST`, `LOGEST`, `TREND`, `GROWTH`, and the old names of all of them (`NORMDIST`, `TDIST`, `CRITBINOM`…): 77 | 19.0 KB |
| `xml` | `FILTERXML`, with its XML reader and XPath: 1 | 13.2 KB |
| all five | 173 | 65.4 KB |

The core keeps the functions most workbooks use, the rare ones of their categories aside: `AVERAGE`, `MEDIAN`, `RANK`,
`STDEV`, `CORREL`, `SLOPE`, `FORECAST` and the rest of the descriptive statistics; `PMT`, `NPV`, `IRR`, `XIRR`, `RATE`
and the other annuities; `LAMBDA` and `LET` themselves, which are syntax. Each group's functions are listed in its
`functions` field (`engineering.functions`).

## A function of a group not loaded

Excel has no such state, so the engine treats the call as Excel treats a function it does not have: the formula is
`#NAME?`, and its issue tells why, with the check `function_not_loaded` and the function in `fn`:

```js
const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'BESSELJ(2,1)' } } }] })
r.diagnostics[0].reason // 'BESSELJ is in the function group engineering, which is not loaded (registerFunctions)'
```

A defined name over such a function returns `#NAME?` where it is used, with the same reason among the name issues. With
`unknownFunctions: 'held'`, the formula keeps the value its file stores, with `fromFile`, as a formula over an add-in's
function does.

The checks of formula text know every function whatever is loaded: `validateFormula('BESSELJ(1)')` refuses the call for
its argument count, `parseEntry` takes `=besselj(A1,1)` as `BESSELJ(A1,1)`, `xlcalc/function-text` describes it, and
`FUNCTION_NAMES` lists it. A workbook's external functions never take the name of one of them.

## Loading a group later

A group can be loaded at any time, for example once a workbook's issues name it. The calculations from then on have its
functions. A session opened before builds its workbook again at its next call, and the formulas over the group are
calculated then: in a session, an inline calculator and a pool's threads alike, in both kinds of execution. This tells
no change and takes no revision, so a host that loads a group for an open session reads the cells it shows again
(`getViewport`, `getValues`). The history of edits stays as it was: undo and redo after the rebuild work as before.

## In a pool

A pool's threads load the groups the host registered: a thread imports each group's module itself, before its next task,
so a bundler that splits the thread's dynamic imports gives each group a chunk of its own. A group registered after the
pool was created is sent to its threads the same way. A pool's threads take the groups registered once the host has
called `xlcalc/core` (a calculation, a session or `registerFunctions`); a pool created before that, or by a host that
uses only the root, has every group.

## Checked against Excel

The functions of the groups are the ones the coverage run and the case sets check ([Functions](functions.md)); loading
them changes no value. The tests calculate a workbook over a function of each group by the root, by the light entry
without groups and with them, in a session opened before the groups were loaded, and in a pool's thread, in both kinds
of execution (`test/function-groups.test.ts`); the root's hashes of the synthetic workbook are those of the release
before the groups.
