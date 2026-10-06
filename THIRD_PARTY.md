# Third-party code in formula-engine

formula-engine has no runtime dependencies. Part of its code is ported (copied, typed and modified) from two MIT-licensed projects; each file that holds ported code names its source, version and copyright in its header.

| project | version | license | ported into |
|---|---|---|---|
| [formula.js](https://github.com/formulajs/formulajs) (`@formulajs/formulajs`) | 4.6.1 | MIT | `src/functions/math.ts`, `statistical.ts`, `logical.ts`, `information.ts`, `text/strings.ts`, `text/numbers.ts`, `financial.ts`, `lookup/reference.ts`, `common.ts` (from `src/math-trig.js`, `src/statistical.js`, `src/compatibility.js`, `src/logical.js`, `src/information.js`, `src/text.js`, `src/financial.js`, `src/lookup-reference.js`) |
| [jStat](https://github.com/jstat/jstat) (`jstat`) | 1.9.6 | MIT | `src/functions/jstat.ts` (from `src/special.js`, `src/distribution.js`) |

The MIT license asks that the copyright notice and the license text be kept with copies of the code: they are below, and the package ships this file.

Development only (not in the package): `@formulajs/formulajs` 4.6.1 (MIT; with its dependencies `jstat` and `bessel`, Apache-2.0) is the measured object of the coverage measurement in `bench/coverage`; `exceljs` (MIT) writes workbooks for the LibreOffice oracle and the measurement, `esbuild` (MIT) builds the browser bundle in the tests, plus `vitest`, `tsx`, `typescript`, `publint` and `@arethetypeswrong/cli`. No code of `bessel` is in the package.

HyperFormula (GPL-3.0) is not used in any form; `test/package.test.ts` checks that it is in no dependency and not in the lock file.

## License texts

### @formulajs/formulajs 4.6.1

From `LICENSE` of the package (<https://github.com/formulajs/formulajs>):

```text
Copyright (c) 2014 Sutoiku, Inc.

The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit
persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the
Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Other libraries included:

BESSELI, BESSELJ, BESSELK, BESSELY functions:

Copyright (c) 2013 SheetJS

The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

jStat - JavaScript Statistical Library:

Copyright (c) 2013 jStat

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

### jStat 1.9.6

From `LICENSE` of the package (<https://github.com/jstat/jstat>):

```text
Copyright (c) 2013 jStat

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```
