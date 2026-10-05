/**
 * Сравнение колонок замера с эталоном Excel en-US: сводка по функциям (`summary.json`) и таблицы
 * отчёта `docs/coverage-formulajs.md`. Таблицы стоят между метками `<!-- auto:имя -->` и
 * `<!-- /auto:имя -->` и пересобираются при каждом запуске; текст вне меток пишется руками и не
 * трогается.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { CaseSet } from './cases.js'
import { compare, show, type Kind, type Verdict } from './compare.js'
import type { FjResult, Val } from './formulajs.js'

interface Col<T = Val> {
  meta: Record<string, string | number>
  values: Record<string, T>
}

/** Интерфейс Excel русский: логическое в тексте выходит ИСТИНА/ЛОЖЬ и при локали en-US. */
const normBool = (v: Val): Val => (v.t === 's' && /ИСТИНА|ЛОЖЬ/.test(v.v) ? { t: 's', v: v.v.replace(/ИСТИНА/g, 'TRUE').replace(/ЛОЖЬ/g, 'FALSE') } : v)

export type Status = 'все' | 'обычные' | 'расходятся обычные' | 'зависает'

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

const CAT_RU: Record<string, string> = {
  'Math and trigonometry': 'математические',
  Statistical: 'статистические',
  Compatibility: 'совместимость',
  'Date and time': 'дата и время',
  Engineering: 'инженерные',
  Financial: 'финансовые',
  Information: 'информационные',
  Logical: 'логические',
  'Lookup and reference': 'поиск и ссылки',
  Text: 'текстовые',
  Database: 'база данных',
  Cube: 'кубы',
  Web: 'веб',
  'Add-in and Automation': 'надстройки',
}
export const catRu = (c: string) => CAT_RU[c] ?? c

function status(list: { kind: string }[], vs: Verdict[]): Status {
  if (vs.some((v) => v.kind === 'зависает')) return 'зависает'
  if (vs.every((v) => v.ok)) return 'все'
  return list.every((c, i) => c.kind !== 'typical' || vs[i].ok) ? 'обычные' : 'расходятся обычные'
}

const count = <T extends string>(xs: (T | undefined)[]) => {
  const out: Partial<Record<T, number>> = {}
  for (const x of xs) if (x) out[x] = (out[x] ?? 0) + 1
  return out
}

export function analyse(dir: string) {
  const load = <T>(n: string) => JSON.parse(readFileSync(join(dir, n), 'utf8')) as T
  const set = load<CaseSet>('cases.json')
  const fj = load<Col<FjResult>>('formulajs.json')
  const enRaw = load<Col>('excel-en.json')
  const ru = load<Col>('excel-ru.json')
  const lo = load<Col>('libreoffice.json')
  const en: Record<string, Val> = Object.fromEntries(Object.entries(enRaw.values).map(([k, v]) => [k, normBool(v)]))
  const boolText = set.cases.filter((c) => enRaw.values[c.id] !== en[c.id])

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
  return { set, fj, en, enRaw, ru, lo, rows, verdicts, boolText }
}

// ---------- markdown ----------

/** Классы краевых случаев для предложения: что чинится общим правилом адаптера, что нет. */
const TAG_CLASS: Record<string, string> = Object.fromEntries(
  Object.entries({
    приведение: ['логическое', 'текст-число', 'число текстом с точкой', 'текст-число в ячейке', 'ссылка на текст-число', 'ссылка на логическое', 'логическое аргументом', 'текст-число аргументом', 'текст аргументом', 'текст TRUE', 'логическое в ячейке', 'число 0', 'число 2', 'текст', 'число', 'ЛОЖЬ'],
    пустое: ['пустая ячейка', 'пустой диапазон', 'пустая строка', 'пустая ячейка аргументом', 'пустая строка в ячейке'],
    ошибки: ['ошибка на входе', '#N/A', '#DIV/0!', 'диапазон с ошибкой', 'поток с ошибкой'],
    диапазоны: ['диапазон: смесь', 'разный размер', 'праздники: смесь', 'праздники: пустой диапазон', 'праздники с ошибкой', 'поток: смесь'],
    'текст-дата': ['дата текстом ISO', 'дата текстом m/d/yyyy', 'время текстом'],
    область: ['отрицательное', 'ноль', 'дробное', 'большое', 'больше 1', 'отрицательные и дробные', 'дата в ячейке', 'больше суток', `большое (${100})`, 'большое (1000)'],
    'даты 1900': ['28.02.1900', 'ложный 29.02.1900', '01.03.1900', 'високосный 29.02.2024', 'дата со временем'],
    текст: ['кириллица', 'регистр', 'пробелы', 'кириллица в ячейке'],
  }).flatMap(([cls, tags]) => tags.map((t) => [t, cls])),
)

const esc = (s: string) => s.replace(/\|/g, '\\|').replace(/\n/g, '\\n')
const code = (s: string) => '`' + s.replace(/`/g, "'") + '`'
const statusRu: Record<Status, string> = { все: 'все', обычные: 'только обычные', 'расходятся обычные': 'расходятся обычные', зависает: 'зависает' }

function kindsText(k: Partial<Record<Kind, number>>): string {
  return Object.entries(k)
    .sort((a, b) => b[1] - a[1])
    .map(([n, c]) => `${n} ${c}`)
    .join(', ')
}

function blocks(dir: string): Record<string, string> {
  const a = analyse(dir)
  const { rows, set } = a
  const st = (sel: (r: FnRow) => Status) => {
    const c = count(rows.map(sel))
    return { все: c['все'] ?? 0, обычные: c['обычные'] ?? 0, 'расходятся обычные': c['расходятся обычные'] ?? 0, зависает: c['зависает'] ?? 0 }
  }
  const sa = st((r) => r.adapter.status)
  const sb = st((r) => r.bare.status)
  const casesOk = (k: 'adapter' | 'bare') => Object.values(a.verdicts).filter((v) => v[k].ok).length
  const softFns = rows.filter((r) => r.adapter.status !== 'все' && r.adapter.ok + r.adapter.soft === r.cases)
  const fixedByAdapter = rows.filter((r) => r.adapter.status === 'все' && r.bare.status !== 'все')
  const brokenByAdapter = rows.filter((r) => r.adapter.ok < r.bare.ok)

  const out: Record<string, string> = {}
  out.summary = [
    '| | через адаптер | голая formulajs |',
    '|---|---|---|',
    `| функций в наборе | ${rows.length} | ${rows.length} |`,
    `| все случаи совпали с Excel en-US | ${sa['все']} | ${sb['все']} |`,
    `| совпали только обычные случаи | ${sa['обычные']} | ${sb['обычные']} |`,
    `| расходятся уже обычные случаи | ${sa['расходятся обычные']} | ${sb['расходятся обычные']} |`,
    `| зависает (на одном из случаев) | ${sa['зависает']} | ${sb['зависает']} |`,
    `| случаев совпало | ${casesOk('adapter')} из ${set.cases.length} | ${casesOk('bare')} из ${set.cases.length} |`,
    '',
    `С мягким допуском 1e-6 через адаптер полностью совпадают ещё ${softFns.length}: ${softFns.map((r) => r.fn).join(', ') || 'нет'}.`,
    `Адаптер довёл до полного совпадения ${fixedByAdapter.length} функций, где голая formulajs расходится: ${fixedByAdapter.map((r) => r.fn).join(', ')}.`,
    `Хуже голой через адаптер ${brokenByAdapter.length}: ${brokenByAdapter.map((r) => `${r.fn} (${r.adapter.ok} против ${r.bare.ok})`).join(', ') || 'нет'}.`,
  ].join('\n')

  // виды расхождений: случаи и функции
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
    '| вид расхождения | случаев через адаптер | функций через адаптер | случаев голой | функций голой |',
    '|---|---|---|---|---|',
    ...allKinds.map((k) => `| ${k} | ${ka.get(k)?.cases ?? 0} | ${ka.get(k)?.fns.size ?? 0} | ${kb.get(k)?.cases ?? 0} | ${kb.get(k)?.fns.size ?? 0} |`),
  ].join('\n')

  // краевые метки, на которых адаптер расходится у функций со совпавшими обычными
  const tagFns = new Map<string, Set<string>>()
  for (const c of set.cases) {
    const r = rows.find((x) => x.fn === c.fn)!
    if (r.adapter.status !== 'обычные' || a.verdicts[c.id].adapter.ok) continue
    tagFns.set(c.tag, (tagFns.get(c.tag) ?? new Set()).add(c.fn))
  }
  out.tags = [
    '| краевой случай | функций | функции |',
    '|---|---|---|',
    ...[...tagFns]
      .sort((x, y) => y[1].size - x[1].size)
      .map(([t, fs]) => `| ${t} | ${fs.size} | ${[...fs].sort().join(', ')} |`),
  ].join('\n')

  const sorted = [...rows].sort((x, y) => x.category.localeCompare(y.category) || x.fn.localeCompare(y.fn))
  out.functions = [
    '| функция | категория | случаев | совпало через адаптер | совпало голой | итог через адаптер | вид расхождения через адаптер | на каких случаях |',
    '|---|---|---|---|---|---|---|---|',
    ...sorted.map(
      (r) =>
        `| ${r.fn} | ${catRu(r.category)} | ${r.cases} | ${r.adapter.ok} | ${r.bare.ok} | ${statusRu[r.adapter.status]}${r.adapter.soft ? `, мягко ещё ${r.adapter.soft}` : ''} | ${kindsText(r.adapter.kinds)} | ${esc(r.adapter.tags.join('; '))} |`,
    ),
  ].join('\n')

  // зависания formulajs
  const hangs = set.cases.filter((c) => a.fj.values[c.id].adapter.t === 'other' && ['hang', 'crash'].includes(String(a.fj.values[c.id].adapter.v)))
  out.hangs = hangs.length
    ? ['| случай | формула | что | Excel en-US |', '|---|---|---|---|', ...hangs.map((c) => `| ${c.id} | ${code(c.formula)} | ${a.fj.values[c.id].adapter.v === 'hang' ? 'висит дольше 5 с' : 'падает (память)'} | ${show(a.en[c.id])} |`)].join('\n')
    : 'Зависаний нет.'

  // LibreOffice против Excel
  const loRows = rows.filter((r) => r.lo > 0).sort((x, y) => y.lo - x.lo || x.fn.localeCompare(y.fn))
  const loCases = Object.values(a.verdicts).filter((v) => !v.lo.ok).length
  const loExample = (fn: string) => {
    const c = set.cases.find((x) => x.fn === fn && !a.verdicts[x.id].lo.ok)!
    return `${code(c.formula)}: Excel ${show(a.en[c.id])}, LibreOffice ${show(a.lo.values[c.id])}`
  }
  const loKinds = count(Object.values(a.verdicts).map((v) => (v.lo.ok ? undefined : v.lo.kind)))
  const loTypical = set.cases.filter((c) => c.kind === 'typical' && !a.verdicts[c.id].lo.ok)
  out.libreoffice = [
    `LibreOffice расходится с Excel en-US в ${loCases} случаях из ${set.cases.length}, у ${loRows.length} функций из ${rows.length}. По видам: ${kindsText(loKinds)}.`,
    '',
    `На обычных случаях расходится ${loTypical.length} раз у ${new Set(loTypical.map((c) => c.fn)).size} функций:`,
    '',
    '| случай | формула | Excel en-US | LibreOffice |',
    '|---|---|---|---|',
    ...loTypical.map((c) => `| ${c.id} | ${code(c.formula)} | ${esc(show(a.en[c.id]))} | ${esc(show(a.lo.values[c.id]))} |`),
    '',
    'Все функции, где LibreOffice расходится хотя бы в одном случае:',
    '',
    '| функция | случаев расходится | пример |',
    '|---|---|---|',
    ...loRows.map((r) => `| ${r.fn} | ${r.lo} из ${r.cases} | ${esc(loExample(r.fn))} |`),
  ].join('\n')

  // ru против en: по краевому случаю, затем полный список
  const ruCases = set.cases.filter((c) => !a.verdicts[c.id].ru.ok)
  const ruByTag = new Map<string, typeof ruCases>()
  for (const c of ruCases) ruByTag.set(c.tag, [...(ruByTag.get(c.tag) ?? []), c])
  const ruLine = (c: (typeof ruCases)[number]) => `${code(c.formula)}: ${esc(show(a.enRaw.values[c.id]))} и ${esc(show(a.ru.values[c.id]))}`
  out.ru = [
    `Excel ru-RU расходится с en-US в ${ruCases.length} случаях у ${new Set(ruCases.map((c) => c.fn)).size} функций (сравниваются ответы Excel как есть, без поправки на ИСТИНА/ЛОЖЬ).`,
    '',
    '| случай | расходится | функции | пример (en-US и ru-RU) |',
    '|---|---|---|---|',
    ...[...ruByTag].sort((x, y) => y[1].length - x[1].length).map(([t, cs]) => `| ${t} | ${cs.length} | ${[...new Set(cs.map((c) => c.fn))].join(', ')} | ${ruLine(cs[0])} |`),
    '',
    '<details><summary>Все случаи</summary>',
    '',
    '| случай | формула | en-US | ru-RU |',
    '|---|---|---|---|',
    ...ruCases.map((c) => `| ${c.id} | ${code(c.formula)} | ${esc(show(a.enRaw.values[c.id]))} | ${esc(show(a.ru.values[c.id]))} |`),
    '',
    '</details>',
  ].join('\n')

  out.booltext = [
    `Таких случаев ${a.boolText.length}:`,
    '',
    ...a.boolText.map((c) => `- ${code(c.formula)}: Excel ${esc(show(a.enRaw.values[c.id]))}`),
  ].join('\n')

  // нет в formulajs
  const miss = new Map<string, string[]>()
  for (const m of set.missing) miss.set(m.category, [...(miss.get(m.category) ?? []), m.fn + (m.note ? '*' : '')])
  out.missing = [
    `Всего ${set.missing.length} из ${set.missing.length + set.inFormulajs.length} функций справочника. Звёздочка: в formulajs это имя есть, но это объект с подфункциями, а не функция (вызвать нельзя).`,
    '',
    '| категория | число | функции |',
    '|---|---|---|',
    ...[...miss].sort((x, y) => y[1].length - x[1].length).map(([cat, fs]) => `| ${catRu(cat)} | ${fs.length} | ${fs.join(', ')} |`),
  ].join('\n')

  out.notinexcel = [
    '| имя formulajs | та же реализация, что у функции Excel |',
    '|---|---|',
    ...set.notInExcel.map((x) => `| ${x.name} | ${x.same.join(', ') || 'нет'} |`),
  ].join('\n')

  // предложение: по классам расхождений через адаптер (формат ответа, класс краевого случая)
  const FIXABLE = new Set(['приведение', 'пустое', 'ошибки', 'диапазоны', 'текст-дата', 'формат'])
  const DOMAIN = new Set(['область', 'даты 1900', 'код ошибки'])
  const plan = new Map<string, string[]>()
  for (const r of sorted) {
    const fails = set.cases.filter((c) => c.fn === r.fn && !a.verdicts[c.id].adapter.ok)
    const cls = fails.map((c) => {
      const k = a.verdicts[c.id].adapter.kind
      if (k === 'строка: регистр' || k === 'строка: формат числа') return 'формат'
      if (c.kind === 'typical') return 'обычный'
      if (k === 'другой код ошибки') return 'код ошибки'
      return TAG_CLASS[c.tag] ?? 'семантика'
    })
    const key =
      r.adapter.status === 'все'
        ? 'как есть'
        : r.adapter.status === 'зависает'
          ? 'своя или страж (зависает)'
          : cls.every((c) => FIXABLE.has(c))
            ? 'адаптер'
            : cls.every((c) => FIXABLE.has(c) || DOMAIN.has(c))
              ? 'адаптер и проверка области'
              : cls.includes('обычный')
                ? 'своя (расходятся обычные)'
                : 'своя (расходятся по смыслу на краях)'
    plan.set(key, [...(plan.get(key) ?? []), r.fn])
  }
  const order = ['как есть', 'адаптер', 'адаптер и проверка области', 'своя (расходятся по смыслу на краях)', 'своя (расходятся обычные)', 'своя или страж (зависает)']
  out.plan = ['| предложение | функций | функции |', '|---|---|---|', ...order.filter((k) => plan.has(k)).map((k) => `| ${k} | ${plan.get(k)!.length} | ${plan.get(k)!.join(', ')} |`)].join('\n')

  out.meta = [
    `- formulajs ${a.fj.meta.formulajs}, Node ${a.fj.meta.node}, часовой пояс ${a.fj.meta.tz}, расчёт ${a.fj.meta.seconds} с;`,
    `- Excel ${a.enRaw.meta.version} на Маке: en-US ${a.enRaw.meta.seconds} с, ru-RU ${a.ru.meta.seconds} с (${a.enRaw.meta.parts} частей, без учёта перезапуска Excel);`,
    `- ${String(a.lo.meta.version).split(' ').slice(0, 2).join(' ')}, профиль en-US, ${a.lo.meta.seconds} с.`,
  ].join('\n')

  writeFileSync(join(dir, 'summary.json'), JSON.stringify({ rows, verdicts: a.verdicts }, null, 1) + '\n')
  return out
}

/** Пересобирает блоки отчёта; без отчёта создаёт заготовку со всеми блоками. */
export function writeReport(dir: string, pilot: boolean): void {
  const b = blocks(dir)
  const md = pilot ? join(dir, 'report.md') : join(dir, '../coverage-formulajs.md')
  let text = existsSync(md) ? readFileSync(md, 'utf8') : Object.keys(b).map((k) => `<!-- auto:${k} -->\n<!-- /auto:${k} -->\n`).join('\n')
  for (const [k, v] of Object.entries(b)) {
    const re = new RegExp(`<!-- auto:${k} -->[\\s\\S]*?<!-- /auto:${k} -->`)
    if (!re.test(text)) text += `\n<!-- auto:${k} -->\n<!-- /auto:${k} -->\n`
    text = text.replace(re, () => `<!-- auto:${k} -->\n${v}\n<!-- /auto:${k} -->`)
  }
  writeFileSync(md, text)
  console.log(`отчёт: ${md}`)
}
