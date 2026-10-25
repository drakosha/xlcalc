/**
 * Acceptance against Excel in each locale but en-US on the coverage measurement: every whitelisted function's cases
 * computed in a book of the locale and compared with Excel's answers there (`test/excel/coverage/excel-ru.json` for
 * ru-RU), as `excel.test.ts` does for en-US. Known differences are pinned in `KNOWN_BY_LOCALE` (`excel-locale-cases.ts`).
 * A locale Excel has not computed the measurement in is a skipped test that names the missing file.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { FUNCTION_NAMES } from '../src/index.js'
import { CASES } from './excel-cases.js'
import { describeOff, KNOWN_BY_LOCALE, LOCALE_EXCEL, LOCALE_SETS, LOCALE_TAGS, measurementPath, reasonText, runLocaleCases, type LocaleOutcome } from './excel-locale-cases.js'
import { pinnedCases } from './stored-cases.js'

for (const tag of LOCALE_TAGS) {
  const excel = LOCALE_EXCEL.get(tag)
  if (!excel) it.skip(`acceptance against Excel ${tag}: no answers of Excel ${tag} to the coverage measurement (${measurementPath(tag)})`)
  else {
    const outcomes = runLocaleCases(tag, CASES.filter((c) => FUNCTION_NAMES.includes(c.fn)))
    const byFunction = new Map<string, LocaleOutcome[]>()
    for (const o of outcomes) byFunction.set(o.fn!, [...(byFunction.get(o.fn!) ?? []), o])

    describe(`acceptance against Excel ${tag}: the coverage measurement`, () => {
      it(`the stored Excel ${tag} answers cover every case`, () => {
        expect(CASES.filter((c) => !excel[c.id]).map((c) => c.id)).toEqual([])
      })

      for (const [f, os] of byFunction)
        it(`${f}: all cases match Excel ${tag}`, () => {
          expect(os.filter((o) => !o.ok).map(describeOff)).toEqual([])
        })
    })
  }

  const known = KNOWN_BY_LOCALE[tag]
  if (known)
    describe(`known differences from Excel ${tag}`, () => {
      it(`every entry of KNOWN_BY_LOCALE['${tag}'] names an existing case`, () => expect(pinnedCases([{ name: 'measurement', cases: CASES }, ...LOCALE_SETS.get(tag)!], known)).toEqual([]))

      it(`docs/limitations.md's known differences have a line for every reason in KNOWN_BY_LOCALE['${tag}']`, () => {
        const readme = readFileSync(new URL('../docs/limitations.md', import.meta.url), 'utf8')
        const section = readme.slice(readme.indexOf('- **Known differences from Excel:**'), readme.indexOf('- **Syntax.**'))
        const used = [...new Set(Object.values(known).map((k) => k.reason))]
        expect(used.filter((r) => !section.includes(`  - ${reasonText(r)}\n`))).toEqual([])
      })
    })
}
