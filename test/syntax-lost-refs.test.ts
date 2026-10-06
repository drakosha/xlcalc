/** References Excel lost, as it writes them after deleting rows or a sheet: read as #REF!, printed back as written. */
import { describe, expect, it } from 'vitest'
import { compute, validateFormula } from '../src/index.js'

describe('lost references', () => {
  it('parse to #REF! and keep their sheet or reference in the canonical text', () => {
    const canon = (f: string) => {
      const r = validateFormula(f, { sheets: ['S', 'My Data'] })
      return r.ok ? r.canonical : r.code
    }
    expect(canon('s!#ref!')).toBe("'s'!#REF!")
    expect(canon("SUM('My Data'!#REF!,1)")).toBe("SUM('My Data'!#REF!,1)")
    expect(canon('#REF!A5+#REF!$A$3:$B$8')).toBe('#REF!A5+#REF!$A$3:$B$8')
    expect(canon('SUM(#REF!B:B,#REF!30:31)')).toBe('SUM(#REF!B:B,#REF!30:31)')
    expect(canon('#REF!#REF!')).toBe('#REF!#REF!')
    expect(canon("'Gone'!#REF!")).toBe('unknown_sheet')
    expect(canon('#REF!x')).toBe('syntax')
  })

  it('compute to #REF!', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: "s!#REF!" }, A2: { formula: 'SUM(#REF!A1:B2)' }, A3: { formula: '#REF!#REF!+1' } } }] })
    expect(Object.values(r.sheets[0].cells).map((c) => [c.value, c.formula])).toEqual([
      [{ error: 'REF' }, "'S'!#REF!"],
      [{ error: 'REF' }, 'SUM(#REF!A1:B2)'],
      [{ error: 'REF' }, '#REF!#REF!+1'],
    ])
  })

  it('a name whose reference is lost gives #REF! where it is used, as in Excel, and is no name issue', () => {
    const r = compute({
      sheets: [{ name: 'S', cells: { A1: { formula: 'gone' }, A2: { formula: 'SUM(gone,1)' }, A3: { formula: 'deleted*2' }, A4: { formula: 'IFERROR(both,7)' }, A5: { formula: 'SUMIF(loc,">0")' }, A6: { formula: 'ERROR.TYPE(gone)' } } }],
      names: [
        { name: 'gone', ref: "'S'!#REF!" },
        { name: 'deleted', ref: '#REF!$A$5' },
        { name: 'both', ref: '#REF!#REF!' },
        { name: 'loc', ref: "'S'!#REF!", scope: 'S' },
      ],
    })
    expect(r.nameIssues).toEqual([])
    expect(Object.values(r.sheets[0].cells).map((c) => c.value)).toEqual([{ error: 'REF' }, { error: 'REF' }, { error: 'REF' }, 7, { error: 'REF' }, 4])
  })
})
