/**
 * Замер покрытия formulajs против Excel. Шаги (каждый пишет свой JSON в `docs/coverage-formulajs/`):
 *
 *   npx tsx bench/coverage/run.ts cases      случаи и книга (bench/coverage/out/book.xlsx)
 *   npx tsx bench/coverage/run.ts fj         formulajs через адаптер и голая
 *   npx tsx bench/coverage/run.ts lo         LibreOffice en-US
 *   npx tsx bench/coverage/run.ts excel en   Excel en-US (перезапускает Excel с английским интерфейсом и локалью)
 *   npx tsx bench/coverage/run.ts excel ru   Excel ru-RU (перезапускает Excel без аргументов)
 *   npx tsx bench/coverage/run.ts report     сравнение, сводка и таблицы отчёта
 *   npx tsx bench/coverage/run.ts all        всё подряд, в конце Excel остаётся запущенным без аргументов
 *   npx tsx bench/coverage/run.ts probes     пробы Excel en-US вне замера (probes.ts) в docs/excel-probes.json
 *   npx tsx bench/coverage/run.ts probes X   набор sets/X.ts в docs/excel-cases/X.json (тест test/excel-sets.test.ts)
 *   npx tsx bench/coverage/run.ts struct X   операции над структурой struct/X.ts в docs/excel-struct/X.json (тест test/structure-excel.test.ts)
 *
 * FCOV_ONLY=ABS,SUM ограничивает набор функций (пробный прогон): тогда всё пишется в out/.
 * Excel считается частями до 400 случаев; после каждой части проверяется, нет ли модального окна.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeBook } from './book.js'
import { buildCases, type Case, type CaseSet } from './cases.js'
import { ensureExcel, excelDialogs, excelVersion, runExcel, runExcelWatched, declineRepair, withExcelLock, type Locale } from './excel.js'
import type { FjResult, Val } from './formulajs.js'
import { libreOfficeVersion, runLibreOffice } from './libreoffice.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from './probes.js'
import { writeReport } from './report.js'
import { stepStruct } from './struct.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const ONLY = process.env.FCOV_ONLY?.split(',').filter(Boolean)
const OUT = join(HERE, 'out')
export const RESULTS = ONLY ? OUT : join(ROOT, 'docs/coverage-formulajs')
const BOOK = join(OUT, 'book.xlsx')

export interface Column<T = Val> {
  meta: Record<string, string | number>
  values: Record<string, T>
}

const save = (name: string, data: unknown) => {
  mkdirSync(RESULTS, { recursive: true })
  writeFileSync(join(RESULTS, name), JSON.stringify(data, null, 1) + '\n')
}
export const load = <T>(name: string): T => JSON.parse(readFileSync(join(RESULTS, name), 'utf8')) as T

function caseSet(): CaseSet {
  const set = buildCases()
  return ONLY ? { ...set, cases: set.cases.filter((c) => ONLY.includes(c.fn)) } : set
}

async function stepCases() {
  const set = caseSet()
  save('cases.json', set)
  mkdirSync(OUT, { recursive: true })
  await writeBook(set.cases, BOOK)
  console.log(`случаев ${set.cases.length}, функций ${new Set(set.cases.map((c) => c.fn)).size}, книга ${BOOK}`)
}

const cases = (): Case[] => load<CaseSet>('cases.json').cases

/** formulajs в дочернем процессе: зависший случай (дольше 5 с) убивается и помечается. */
async function stepFj() {
  const cs = cases()
  const file = join(RESULTS, 'cases.json')
  const out: Record<string, FjResult> = {}
  const t0 = Date.now()
  let start = 0
  while (start < cs.length) {
    start = await new Promise<number>((resolve) => {
      const child = spawn(process.execPath, ['--max-old-space-size=2048', '--import', 'tsx', join(HERE, 'fj-worker.ts'), file, String(start)], { stdio: ['ignore', 'pipe', 'ignore'] })
      let next = start
      let buf = ''
      let timer: NodeJS.Timeout
      let killed = false
      const arm = () => {
        clearTimeout(timer)
        timer = setTimeout(() => {
          killed = true
          child.kill('SIGKILL')
        }, 5_000)
      }
      arm()
      child.stdout.on('data', (d: Buffer) => {
        buf += d.toString('utf8')
        let nl: number
        while ((nl = buf.indexOf('\n')) >= 0) {
          const { i, r } = JSON.parse(buf.slice(0, nl)) as { i: number; r: FjResult }
          buf = buf.slice(nl + 1)
          out[cs[i].id] = r
          next = i + 1
          arm()
        }
      })
      child.on('exit', () => {
        clearTimeout(timer)
        if (next < cs.length) {
          const fail: Val = { t: 'other', v: killed ? 'hang' : 'crash' }
          out[cs[next].id] = { adapter: fail, bare: fail }
          console.log(`${killed ? 'завис' : 'упал'}: ${cs[next].id} ${cs[next].formula}`)
          next++
        }
        resolve(next)
      })
    })
  }
  const pkg = JSON.parse(readFileSync(join(ROOT, 'node_modules/@formulajs/formulajs/package.json'), 'utf8')) as { version: string }
  save('formulajs.json', { meta: { formulajs: pkg.version, node: process.version, tz: Intl.DateTimeFormat().resolvedOptions().timeZone, seconds: (Date.now() - t0) / 1000 }, values: out })
  console.log(`formulajs: ${cs.length} случаев за ${((Date.now() - t0) / 1000).toFixed(1)} с`)
}

const byId = (cs: Case[], vals: Val[]): Record<string, Val> => Object.fromEntries(cs.map((c, i) => [c.id, vals[i]]))

async function stepLo() {
  const cs = cases()
  if (!existsSync(BOOK)) throw new Error('нет книги: сначала шаг cases')
  const t0 = Date.now()
  const vals = runLibreOffice(BOOK, cs.length)
  save('libreoffice.json', { meta: { version: libreOfficeVersion(), locale: 'en-US', seconds: (Date.now() - t0) / 1000 }, values: byId(cs, vals) })
  console.log(`LibreOffice: ${((Date.now() - t0) / 1000).toFixed(1)} с`)
}

/** Части прогона Excel: подряд по функциям, не больше PART случаев в книге. */
const PART = 400

function parts(cs: Case[]): Case[][] {
  const out: Case[][] = []
  let cur: Case[] = []
  for (let i = 0; i < cs.length; ) {
    let j = i
    while (j < cs.length && cs[j].fn === cs[i].fn) j++
    if (cur.length && cur.length + (j - i) > PART) {
      out.push(cur)
      cur = []
    }
    cur.push(...cs.slice(i, j))
    i = j
  }
  if (cur.length) out.push(cur)
  return out
}

interface ExcelColumn extends Column {
  meta: Column['meta'] & { parts: number; done: number; stopped?: string }
}

/**
 * Excel частями: перед частью и после неё проверяется, нет ли у Excel окна, которое не окно книги
 * (модальный диалог). Есть: остановка с текстом диалога и номером части, посчитанные части уже в
 * файле. Повторный запуск продолжает с первой непосчитанной части (FCOV_FRESH=1 заново).
 */
async function stepExcel(locale: Locale) {
  const cs = cases()
  const name = `excel-${locale}.json`
  const file = join(RESULTS, name)
  await ensureExcel(locale)
  const ps = parts(cs)
  const col: ExcelColumn =
    existsSync(file) && !process.env.FCOV_FRESH
      ? load<ExcelColumn>(name)
      : { meta: { version: excelVersion(), locale: locale === 'en' ? 'en-US (-AppleLanguages (en) -AppleLocale en_US)' : 'ru-RU (системные язык и локаль)', seconds: 0, parts: ps.length, done: 0 }, values: {} }
  delete col.meta.stopped
  const stop = (k: number, why: string) => {
    col.meta.stopped = `часть ${k + 1} из ${ps.length} (${ps[k][0].id} … ${ps[k][ps[k].length - 1].id}): ${why}`
    save(name, col)
    throw new Error(`Excel ${locale}: ${col.meta.stopped}`)
  }
  mkdirSync(OUT, { recursive: true })
  for (let k = 0; k < ps.length; k++) {
    const part = ps[k]
    if (part.every((c) => c.id in col.values)) continue
    const before = excelDialogs()
    if (before.length) stop(k, `до части у Excel окно не книги: ${before.map((d) => `${d.subrole} «${d.text}»`).join('; ')}`)
    const book = join(OUT, `part-${locale}-${k + 1}.xlsx`)
    await writeBook(part, book)
    const t0 = Date.now()
    let vals: Val[] | null = null
    let err = ''
    try {
      vals = runExcel(book, part.length)
    } catch (e) {
      err = e instanceof Error ? e.message.split('\n')[0] : String(e)
    }
    const after = excelDialogs()
    if (after.length) stop(k, `модальное окно Excel: ${after.map((d) => `${d.subrole} «${d.text}»`).join('; ')}${err ? `; ошибка чтения: ${err}` : ''}`)
    if (!vals) stop(k, `ошибка чтения: ${err}`)
    Object.assign(col.values, byId(part, vals!))
    col.meta.seconds = Math.round((Number(col.meta.seconds) + (Date.now() - t0) / 1000) * 10) / 10
    col.meta.done = ps.filter((p) => p.every((c) => c.id in col.values)).length
    save(name, col)
    console.log(`Excel ${locale}: часть ${k + 1}/${ps.length}, ${part.length} случаев, ${((Date.now() - t0) / 1000).toFixed(1)} с`)
  }
}

/**
 * Пробы в Excel en-US частями по PART, с той же проверкой модальных окон до и после части. Без имени
 * набора это `probes.ts` и `docs/excel-probes.json`; с именем `<set>` генератор `sets/<set>.ts`
 * (экспорт `buildProbes()` и, если нужны листы данных, `SHEETS`) и ответы в `docs/excel-cases/<set>.json`.
 * Наборы нужны, чтобы параллельные ветки добавляли свои случаи, не трогая общие файлы.
 */
async function stepProbes(set?: string) {
  if (set !== undefined && !/^[a-z0-9-]+$/.test(set)) throw new Error(`имя набора: [a-z0-9-], а не ${set}`)
  const mod: { buildProbes: () => Probe[]; SHEETS?: SheetCells } = set ? await import(`./sets/${set}.js`) : { buildProbes, SHEETS: PROBE_SHEETS }
  const ps = mod.buildProbes()
  const sheets = mod.SHEETS ?? {}
  const tag = set ?? 'probes'
  const target = set ? join(ROOT, `docs/excel-cases/${set}.json`) : join(ROOT, 'docs/excel-probes.json')
  // продолжение: ответы, чья формула не изменилась, берутся из файла (FCOV_FRESH=1 заново)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as { cases: Probe[]; values: Record<string, Val> }) : null
  const was = new Map(prev?.cases.map((c) => [c.id, c.formula]) ?? [])
  const values: Record<string, Val> = {}
  for (const p of ps) if (prev?.values[p.id] && was.get(p.id) === p.formula) values[p.id] = prev.values[p.id]
  const rejected: Probe[] = []
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: (Date.now() - t0) / 1000 }
    writeFileSync(target, JSON.stringify({ meta, cases: ps, values, ...(rejected.length ? { rejected: rejected.map((p) => p.id) } : {}) }, null, 1) + '\n')
  }
  // книга, которую Excel не читает, делится пополам до формулы-виновника
  const run = async (cases: Probe[], label: string): Promise<void> => {
    if (declineRepair()) console.log('закрыто оставшееся окно восстановления книги')
    const before = excelDialogs().map((d) => `${d.subrole} «${d.text}»`).join('; ')
    if (before) throw new Error(`пробы ${label}: у Excel окно не книги: ${before}`)
    const book = join(OUT, `${tag}-${label}.xlsx`)
    // скрипт чтения ждёт колонку не короче двух ячеек: одиночный случай дополняется пустышкой
    const sent = cases.length === 1 ? [...cases, { id: 'pad', group: 'pad', formula: '0' }] : cases
    await writeBook(sent, book, sheets)
    const vals = await runExcelWatched(book, sent.length)
    if (vals !== 'rejected') return void cases.forEach((p, i) => (values[p.id] = vals[i]))
    if (cases.length === 1) {
      rejected.push(cases[0])
      return void console.log(`Excel не принимает формулу: ${cases[0].id} ${cases[0].formula}`)
    }
    const mid = cases.length >> 1
    await run(cases.slice(0, mid), `${label}a`)
    await run(cases.slice(mid), `${label}b`)
  }
  await ensureExcel('en')
  mkdirSync(OUT, { recursive: true })
  mkdirSync(dirname(target), { recursive: true })
  for (let k = 0; k * PART < ps.length; k++) {
    const part = ps.slice(k * PART, (k + 1) * PART).filter((p) => !values[p.id])
    if (!part.length) continue
    await run(part, String(k + 1))
    save()
    console.log(`пробы ${tag}: часть ${k + 1}, ${part.length} случаев, ${((Date.now() - t0) / 1000).toFixed(1)} с`)
  }
  save()
  if (rejected.length) throw new Error(`Excel не принимает при загрузке ${rejected.length} формул, убрать из генератора: ${rejected.map((p) => `${p.id} ${p.formula}`).join('; ')}`)
}

const step = process.argv[2]
const t0 = Date.now()
if (step === 'cases') await stepCases()
else if (step === 'fj') await stepFj()
else if (step === 'lo') await stepLo()
else if (step === 'excel') await withExcelLock(() => stepExcel(process.argv[3] as Locale))
else if (step === 'report') writeReport(RESULTS, ONLY !== undefined)
else if (step === 'probes') await withExcelLock(() => stepProbes(process.argv[3]))
else if (step === 'struct') await withExcelLock(() => stepStruct(process.argv[3]))
else if (step === 'all') {
  await stepCases()
  await stepFj()
  await stepLo()
  await withExcelLock(async () => {
    await stepExcel('en')
    await stepExcel('ru')
  })
  writeReport(RESULTS, ONLY !== undefined)
} else {
  console.error('шаг: cases | fj | lo | excel en|ru | report | all | probes [набор] | struct набор')
  process.exit(2)
}
console.log(`готово за ${((Date.now() - t0) / 1000).toFixed(1)} с`)
