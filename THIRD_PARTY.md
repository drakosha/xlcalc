# Third-party code in formula-engine

| package | version | license | used for |
|---|---|---|---|
| `@formulajs/formulajs` | 4.6.1 | MIT | `ROUND`, `EOMONTH`, `DATE`, `SUM` (through the adapter in `src/runtime.ts`) |
| `jstat` (a dependency of formulajs) | 1.9.6 | MIT | statistics inside formulajs |
| `bessel` (a dependency of formulajs) | 1.0.2 | Apache-2.0 | Bessel functions inside formulajs |

All three are permissive licenses: the library can be distributed on its own or inside other software as long as their license texts are kept. They are installed as dependencies, not copied into this package; a bundle that includes formula-engine includes them too.

Development only (not in the package): `exceljs` (MIT) writes workbooks for the LibreOffice oracle, `esbuild` (MIT) builds the browser bundle in the tests, plus `vitest`, `tsx`, `typescript`, `publint` and `@arethetypeswrong/cli`.

HyperFormula (GPL-3.0) is not used in any form; `test/package.test.ts` checks that it is in no dependency and not in the lock file.
