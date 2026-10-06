/**
 * The measurement's columns against the Excel en-US reference: a summary by function (`summary.json`) and the
 * tables of the report `docs/coverage-formulajs.md`. The tables stand between the marks `<!-- auto:name -->`
 * and `<!-- /auto:name -->` and are rebuilt on every run; the text outside the marks is written by hand and
 * left alone.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import type { CaseSet } from './cases.js'
import { compare, show, type Kind, type Verdict } from './compare.js'
import type { FjResult, Val } from './formulajs.js'

interface Col<T = Val> {
  meta: Record<string, string | number>
  values: Record<string, T>
}

export type Status = 'all' | 'typical only' | 'typical differ' | 'hangs'

export interface FnRow {
  fn: string
  category: string
  cases: number
  typical: number
  adapter: { ok: number; okTypical: number; soft: number; status: Status; kinds: Partial<Record<Kind, number>>; tags: string[] }
  bare: { ok: number; okTypical: number; status: Status; kinds: Partial<Record<Kind, number>> }
  lo: number
  ru: number
}

function status(list: { kind: string }[], vs: Verdict[]): Status {
  if (vs.some((v) => v.kind === 'hangs')) return 'hangs'
  if (vs.every((v) => v.ok)) return 'all'
  return list.every((c, i) => c.kind !== 'typical' || vs[i].ok) ? 'typical only' : 'typical differ'
}

const count = <T extends string>(xs: (T | undefined)[]) => {
  const out: Partial<Record<T, number>> = {}
  for (const x of xs) if (x) out[x] = (out[x] ?? 0) + 1
  return out
}

/** `at` gives the path of a measurement file by name. */
export function analyse(at: (name: string) => string) {
  const load = <T>(n: string) => JSON.parse(readFileSync(at(n), 'utf8')) as T
  const set = load<CaseSet>('cases.json')
  const fj = load<Col<FjResult>>('formulajs.json')
  const enRaw = load<Col>('excel-en.json')
  const ru = load<Col>('excel-ru.json')
  const lo = load<Col>('libreoffice.json')
  const en = enRaw.values

  const byFn = new Map<string, typeof set.cases>()
  for (const c of set.cases) byFn.set(c.fn, [...(byFn.get(c.fn) ?? []), c])

  const verdicts: Record<string, { adapter: Verdict; bare: Verdict; lo: Verdict; ru: Verdict }> = {}
  const rows: FnRow[] = []
  for (const [fn, list] of byFn) {
    const va = list.map((c) => compare(en[c.id], fj.values[c.id].adapter, !!fj.values[c.id].adapterThrew))
    const vb = list.map((c) => compare(en[c.id], fj.values[c.id].bare))
    const vl = list.map((c) => compare(en[c.id], lo.values[c.id]))
    const vr = list.map((c) => compare(enRaw.values[c.id], ru.values[c.id]))
    list.forEach((c, i) => (verdicts[c.id] = { adapter: va[i], bare: vb[i], lo: vl[i], ru: vr[i] }))
    const okT = (vs: Verdict[]) => list.filter((c, i) => c.kind === 'typical' && vs[i].ok).length
    rows.push({
      fn,
      category: list[0].category,
      cases: list.length,
      typical: list.filter((c) => c.kind === 'typical').length,
      adapter: {
        ok: va.filter((v) => v.ok).length,
        okTypical: okT(va),
        soft: va.filter((v) => v.soft).length,
        status: status(list, va),
        kinds: count(va.map((v) => v.kind)),
        tags: [...new Set(list.filter((_, i) => !va[i].ok).map((c) => c.tag))],
      },
      bare: { ok: vb.filter((v) => v.ok).length, okTypical: okT(vb), status: status(list, vb), kinds: count(vb.map((v) => v.kind)) },
      lo: vl.filter((v) => !v.ok).length,
      ru: vr.filter((v) => !v.ok).length,
    })
  }
  return { set, fj, en, enRaw, ru, lo, rows, verdicts }
}

// ---------- markdown ----------

/** Classes of edge cases for the proposal: what a shared adapter rule fixes and what it does not. */
const TAG_CLASS: Record<string, string> = Object.fromEntries(
  Object.entries({
    coercion: ['boolean', 'numeric text', 'number text with a point', 'numeric text in a cell', 'reference to numeric text', 'reference to a boolean', 'boolean argument', 'numeric text argument', 'text argument', 'text TRUE', 'boolean in a cell', 'number 0', 'number 2', 'text', 'number', 'FALSE'],
    blank: ['blank cell', 'blank range', 'empty string', 'blank cell argument', 'empty string in a cell'],
    errors: ['error input', '#N/A', '#DIV/0!', 'range with an error', 'cash flow with an error'],
    ranges: ['range: mixed', 'different size', 'holidays: mixed', 'holidays: blank range', 'holidays with an error', 'cash flow: mixed'],
    'date text': ['ISO date text', 'm/d/yyyy date text', 'time text'],
    domain: ['negative', 'zero', 'fractional', 'large', 'above 1', 'negative and fractional', 'date in a cell', 'over a day', 'large (100)', 'large (1000)'],
    '1900 dates': ['1900-02-28', 'false leap day 1900-02-29', '1900-03-01', 'leap day 2024-02-29', 'date with a time'],
    text: ['Cyrillic', 'case', 'spaces', 'Cyrillic in a cell'],
  }).flatMap(([cls, tags]) => tags.map((t) => [t, cls])),
)

const esc = (s: string) => s.replace(/\|/g, '\\|').replace(/\n/g, '\\n')
const code = (s: string) => '`' + s.replace(/`/g, "'") + '`'

function kindsText(k: Partial<Record<Kind, number>>): string {
  return Object.entries(k)
    .sort((a, b) => b[1] - a[1])
    .map(([n, c]) => `${n} ${c}`)
    .join(', ')
}

function blocks(at: (name: string) => string): Record<string, string> {
  const a = analyse(at)
  const { rows, set } = a
  const st = (sel: (r: FnRow) => Status) => {
    const c = count(rows.map(sel))
    return { all: c.all ?? 0, typical: c['typical only'] ?? 0, differ: c['typical differ'] ?? 0, hangs: c.hangs ?? 0 }
  }
  const sa = st((r) => r.adapter.status)
  const sb = st((r) => r.bare.status)
  const casesOk = (k: 'adapter' | 'bare') => Object.values(a.verdicts).filter((v) => v[k].ok).length
  const softFns = rows.filter((r) => r.adapter.status !== 'all' && r.adapter.ok + r.adapter.soft === r.cases)
  const fixedByAdapter = rows.filter((r) => r.adapter.status === 'all' && r.bare.status !== 'all')
  const brokenByAdapter = rows.filter((r) => r.adapter.ok < r.bare.ok)

  const out: Record<string, string> = {}
  out.summary = [
    '| | through the adapter | bare formulajs |',
    '|---|---|---|',
    `| functions in the set | ${rows.length} | ${rows.length} |`,
    `| all cases match Excel en-US | ${sa.all} | ${sb.all} |`,
    `| only the typical cases match | ${sa.typical} | ${sb.typical} |`,
    `| typical cases differ already | ${sa.differ} | ${sb.differ} |`,
    `| hangs (on one of the cases) | ${sa.hangs} | ${sb.hangs} |`,
    `| cases that match | ${casesOk('adapter')} of ${set.cases.length} | ${casesOk('bare')} of ${set.cases.length} |`,
    '',
    `Within the soft tolerance 1e-6, ${softFns.length} more match fully through the adapter: ${softFns.map((r) => r.fn).join(', ') || 'none'}.`,
    `The adapter brought ${fixedByAdapter.length} functions to a full match where bare formulajs differs: ${fixedByAdapter.map((r) => r.fn).join(', ')}.`,
    `Worse through the adapter than bare: ${brokenByAdapter.length}: ${brokenByAdapter.map((r) => `${r.fn} (${r.adapter.ok} against ${r.bare.ok})`).join(', ') || 'none'}.`,
  ].join('\n')

  // kinds of difference: cases and functions
  const kindRows = (k: 'adapter' | 'bare') => {
    const byKind = new Map<string, { cases: number; fns: Set<string> }>()
    for (const c of set.cases) {
      const v = a.verdicts[c.id][k]
      if (v.ok || !v.kind) continue
      const e = byKind.get(v.kind) ?? { cases: 0, fns: new Set() }
      e.cases++
      e.fns.add(c.fn)
      byKind.set(v.kind, e)
    }
    return byKind
  }
  const ka = kindRows('adapter')
  const kb = kindRows('bare')
  const allKinds = [...new Set([...ka.keys(), ...kb.keys()])].sort((x, y) => (ka.get(y)?.cases ?? 0) - (ka.get(x)?.cases ?? 0))
  out.kinds = [
    '| kind of difference | cases through the adapter | functions through the adapter | cases bare | functions bare |',
    '|---|---|---|---|---|',
    ...allKinds.map((k) => `| ${k} | ${ka.get(k)?.cases ?? 0} | ${ka.get(k)?.fns.size ?? 0} | ${kb.get(k)?.cases ?? 0} | ${kb.get(k)?.fns.size ?? 0} |`),
  ].join('\n')

  // the edge tags where the adapter differs, for functions whose typical cases match
  const tagFns = new Map<string, Set<string>>()
  for (const c of set.cases) {
    const r = rows.find((x) => x.fn === c.fn)!
    if (r.adapter.status !== 'typical only' || a.verdicts[c.id].adapter.ok) continue
    tagFns.set(c.tag, (tagFns.get(c.tag) ?? new Set()).add(c.fn))
  }
  out.tags = [
    '| edge case | functions | which |',
    '|---|---|---|',
    ...[...tagFns]
      .sort((x, y) => y[1].size - x[1].size)
      .map(([t, fs]) => `| ${t} | ${fs.size} | ${[...fs].sort().join(', ')} |`),
  ].join('\n')

  const sorted = [...rows].sort((x, y) => x.category.localeCompare(y.category) || x.fn.localeCompare(y.fn))
  out.functions = [
    '| function | category | cases | match through the adapter | match bare | result through the adapter | kind of difference through the adapter | on which cases |',
    '|---|---|---|---|---|---|---|---|',
    ...sorted.map(
      (r) =>
        `| ${r.fn} | ${r.category} | ${r.cases} | ${r.adapter.ok} | ${r.bare.ok} | ${r.adapter.status}${r.adapter.soft ? `, ${r.adapter.soft} more soft` : ''} | ${kindsText(r.adapter.kinds)} | ${esc(r.adapter.tags.join('; '))} |`,
    ),
  ].join('\n')

  // formulajs hangs
  const hangs = set.cases.filter((c) => a.fj.values[c.id].adapter.t === 'other' && ['hang', 'crash'].includes(String(a.fj.values[c.id].adapter.v)))
  out.hangs = hangs.length
    ? ['| case | formula | what | Excel en-US |', '|---|---|---|---|', ...hangs.map((c) => `| ${c.id} | ${code(c.formula)} | ${a.fj.values[c.id].adapter.v === 'hang' ? 'hangs over 5 s' : 'crashes (memory)'} | ${show(a.en[c.id])} |`)].join('\n')
    : 'No hangs.'

  // LibreOffice against Excel
  const loRows = rows.filter((r) => r.lo > 0).sort((x, y) => y.lo - x.lo || x.fn.localeCompare(y.fn))
  const loCases = Object.values(a.verdicts).filter((v) => !v.lo.ok).length
  const loExample = (fn: string) => {
    const c = set.cases.find((x) => x.fn === fn && !a.verdicts[x.id].lo.ok)!
    return `${code(c.formula)}: Excel ${show(a.en[c.id])}, LibreOffice ${show(a.lo.values[c.id])}`
  }
  const loKinds = count(Object.values(a.verdicts).map((v) => (v.lo.ok ? undefined : v.lo.kind)))
  const loTypical = set.cases.filter((c) => c.kind === 'typical' && !a.verdicts[c.id].lo.ok)
  out.libreoffice = [
    `LibreOffice differs from Excel en-US in ${loCases} cases of ${set.cases.length}, for ${loRows.length} functions of ${rows.length}. By kind: ${kindsText(loKinds)}.`,
    '',
    `On typical cases it differs ${loTypical.length} times, for ${new Set(loTypical.map((c) => c.fn)).size} functions:`,
    '',
    '| case | formula | Excel en-US | LibreOffice |',
    '|---|---|---|---|',
    ...loTypical.map((c) => `| ${c.id} | ${code(c.formula)} | ${esc(show(a.en[c.id]))} | ${esc(show(a.lo.values[c.id]))} |`),
    '',
    'Every function where LibreOffice differs in at least one case:',
    '',
    '| function | cases that differ | example |',
    '|---|---|---|',
    ...loRows.map((r) => `| ${r.fn} | ${r.lo} of ${r.cases} | ${esc(loExample(r.fn))} |`),
  ].join('\n')

  // ru-RU against en-US: by edge tag, then the full list
  const ruCases = set.cases.filter((c) => !a.verdicts[c.id].ru.ok)
  const ruByTag = new Map<string, typeof ruCases>()
  for (const c of ruCases) ruByTag.set(c.tag, [...(ruByTag.get(c.tag) ?? []), c])
  const ruLine = (c: (typeof ruCases)[number]) => `${code(c.formula)}: ${esc(show(a.enRaw.values[c.id]))} and ${esc(show(a.ru.values[c.id]))}`
  out.ru = [
    `Excel ru-RU (Russian interface and locale) differs from en-US (English ones) in ${ruCases.length} cases, for ${new Set(ruCases.map((c) => c.fn)).size} functions.`,
    '',
    '| case | differ | functions | example (en-US and ru-RU) |',
    '|---|---|---|---|',
    ...[...ruByTag].sort((x, y) => y[1].length - x[1].length).map(([t, cs]) => `| ${t} | ${cs.length} | ${[...new Set(cs.map((c) => c.fn))].join(', ')} | ${ruLine(cs[0])} |`),
    '',
    '<details><summary>All cases</summary>',
    '',
    '| case | formula | en-US | ru-RU |',
    '|---|---|---|---|',
    ...ruCases.map((c) => `| ${c.id} | ${code(c.formula)} | ${esc(show(a.enRaw.values[c.id]))} | ${esc(show(a.ru.values[c.id]))} |`),
    '',
    '</details>',
  ].join('\n')

  // functions formulajs lacks
  const miss = new Map<string, string[]>()
  for (const m of set.missing) miss.set(m.category, [...(miss.get(m.category) ?? []), m.fn + (m.note ? '*' : '')])
  out.missing = [
    `${set.missing.length} of the reference's ${set.missing.length + set.inFormulajs.length} functions. An asterisk: formulajs has the name, but as an object of subfunctions, not a function (it cannot be called).`,
    '',
    '| category | count | functions |',
    '|---|---|---|',
    ...[...miss].sort((x, y) => y[1].length - x[1].length).map(([cat, fs]) => `| ${cat} | ${fs.length} | ${fs.join(', ')} |`),
  ].join('\n')

  out.notinexcel = [
    '| formulajs name | same implementation as the Excel function |',
    '|---|---|',
    ...set.notInExcel.map((x) => `| ${x.name} | ${x.same.join(', ') || 'none'} |`),
  ].join('\n')

  // the proposal: by the classes of difference through the adapter (answer format, edge class)
  const FIXABLE = new Set(['coercion', 'blank', 'errors', 'ranges', 'date text', 'format'])
  const DOMAIN = new Set(['domain', '1900 dates', 'error code'])
  const plan = new Map<string, string[]>()
  for (const r of sorted) {
    const fails = set.cases.filter((c) => c.fn === r.fn && !a.verdicts[c.id].adapter.ok)
    const cls = fails.map((c) => {
      const k = a.verdicts[c.id].adapter.kind
      if (k === 'string: case' || k === 'string: number format') return 'format'
      if (c.kind === 'typical') return 'typical'
      if (k === 'other error code') return 'error code'
      return TAG_CLASS[c.tag] ?? 'semantics'
    })
    const key =
      r.adapter.status === 'all'
        ? 'as is'
        : r.adapter.status === 'hangs'
          ? 'own or a guard (hangs)'
          : cls.every((c) => FIXABLE.has(c))
            ? 'adapter'
            : cls.every((c) => FIXABLE.has(c) || DOMAIN.has(c))
              ? 'adapter and a domain check'
              : cls.includes('typical')
                ? 'own (typical cases differ)'
                : 'own (edges differ in meaning)'
    plan.set(key, [...(plan.get(key) ?? []), r.fn])
  }
  const order = ['as is', 'adapter', 'adapter and a domain check', 'own (edges differ in meaning)', 'own (typical cases differ)', 'own or a guard (hangs)']
  out.plan = ['| proposal | functions | which |', '|---|---|---|', ...order.filter((k) => plan.has(k)).map((k) => `| ${k} | ${plan.get(k)!.length} | ${plan.get(k)!.join(', ')} |`)].join('\n')

  out.meta = [
    `- formulajs ${a.fj.meta.formulajs}, Node ${a.fj.meta.node}, time zone ${a.fj.meta.tz}, ${a.fj.meta.seconds} s to compute;`,
    `- Excel ${a.enRaw.meta.version} for Mac: en-US ${a.enRaw.meta.seconds} s, ru-RU ${a.ru.meta.seconds} s (${a.enRaw.meta.parts} parts, Excel's restarts not counted);`,
    `- ${String(a.lo.meta.version).split(' ').slice(0, 2).join(' ')}, en-US profile, ${a.lo.meta.seconds} s.`,
  ].join('\n')

  writeFileSync(at('summary.json'), JSON.stringify({ rows, verdicts: a.verdicts }, null, 1) + '\n')
  return out
}

/** Rebuilds the report's blocks; without a report it writes a frame with every block. */
export function writeReport(at: (name: string) => string, md: string): void {
  const b = blocks(at)
  let text = existsSync(md) ? readFileSync(md, 'utf8') : Object.keys(b).map((k) => `<!-- auto:${k} -->\n<!-- /auto:${k} -->\n`).join('\n')
  for (const [k, v] of Object.entries(b)) {
    const re = new RegExp(`<!-- auto:${k} -->[\\s\\S]*?<!-- /auto:${k} -->`)
    if (!re.test(text)) text += `\n<!-- auto:${k} -->\n<!-- /auto:${k} -->\n`
    text = text.replace(re, () => `<!-- auto:${k} -->\n${v}\n<!-- /auto:${k} -->`)
  }
  writeFileSync(md, text)
  console.log(`report: ${md}`)
}
