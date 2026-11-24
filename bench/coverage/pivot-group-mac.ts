/** Build an ungrouped PivotTable for each case, then have Excel for Mac do the grouping itself. */
import { fileURLToPath } from 'node:url'
import type { CellValue, PivotFieldGroup } from '../../src/contract.js'
import { dateToSerial, serialToUtcMs } from '../../src/model.js'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { localeMeta } from './locales.js'
import type { Period, PivotCase, PivotOp, PivotSet } from './pivot/cases.js'
import { drive } from './pivot/drive.js'
import { shown } from './pivot/phases.js'
import { fileAnswers, runFileCases, secondsSince, setModule } from './stored.js'
import { editParts, partsHash } from './xlsx/sent.js'

const PERIODS: Period[] = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']
const LOCALE = 'ru-RU'
const dateInput = new Intl.DateTimeFormat(LOCALE, { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' })
const numberInput = new Intl.NumberFormat(LOCALE, { useGrouping: false, maximumFractionDigits: 15 })

type Group = Extract<PivotOp, { group: string }>
type Range = NonNullable<PivotFieldGroup['range']>

export interface GroupMacAnswer {
  area: string
  values: CellValue[][]
}

/** Excel's dialog reads date bounds as local dates, and numeric bounds and intervals with a decimal comma. */
function formatBound(bound: Group['start']): string {
  if (bound === undefined) return ''
  if (typeof bound === 'string') throw new Error('Grouping bounds must be numbers or date serials')
  return typeof bound === 'object' ? dateInput.format(serialToUtcMs(bound.d)) : numberInput.format(bound)
}

function checkGrouping(id: string, group: Group, ranges: Range[]): void {
  const expected = (group.periods ?? ['numbers']).slice().sort().join(',')
  const actual = ranges.map((range) => range.by).sort().join(',')
  if (actual !== expected) throw new Error(`${id}: Excel grouped by ${actual}, expected ${expected}`)

  for (const range of ranges) {
    for (const bound of ['start', 'end'] as const) {
      const requested = group[bound]
      if (requested === undefined) continue
      const saved = range[bound]
      const serial = typeof saved === 'object' ? dateToSerial(saved.date) : saved
      const expectedSerial = typeof requested === 'object' ? requested.d : requested
      if (serial !== expectedSerial) throw new Error(`${id}: Excel saved the wrong ${bound} of the ${range.by} grouping`)
    }
    const interval = range.interval ?? 1
    if (group.by !== undefined && interval !== group.by) {
      throw new Error(`${id}: Excel saved interval ${interval}, expected ${group.by}`)
    }
  }
}

export async function stepPivotGroupMac(set: string): Promise<void> {
  const { buildPivotCases } = await setModule<PivotSet>('pivot', set)
  const cases = await buildPivotCases()
  const files = new Map<string, { bytes: Uint8Array; args: string[]; group: Group }>()
  for (const c of cases) {
    const group = c.ops.at(-1)
    if (!group || !('group' in group)) throw new Error(`${c.id}: the last step must be a grouping`)
    const [{ book, errors }] = drive({ ...c, ops: c.ops.slice(0, -1) }, undefined, LOCALE)
    if (errors.length) throw new Error(`${c.id}: ${errors.join('; ')}`)
    const periods = group.periods
    const flags = periods ? PERIODS.map((period) => periods.includes(period) ? '1' : '0').join('') : ''
    const args = [flags, formatBound(group.start), formatBound(group.end), numberInput.format(group.by ?? 1)]
    const { data: bytes } = await writeXlsx({ book }, { values: 'omit' })
    files.set(c.id, { bytes, args, group })
  }
  const sent = Object.fromEntries(await Promise.all([...files].map(async ([id, { bytes }]) => [id, await partsHash(bytes)])))
  const path = fileURLToPath(new URL(`../../test/excel/pivot/${LOCALE}/${set}.json`, import.meta.url))
  const store = fileAnswers<PivotCase, GroupMacAnswer>(path, cases, sent)
  const script = fileURLToPath(new URL('./excel-pivot-group.applescript', import.meta.url))
  const t0 = Date.now()
  await ensureExcel(LOCALE)
  const save = () => store.save({ version: excelVersion(), locale: localeMeta(LOCALE), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, { cases })
  await runFileCases(`pivot-group ${set} (Mac)`, store, store.todo, 'FPIVOT_ONLY', async (c) => {
    const { bytes, args, group } = files.get(c.id)!
    const result = await runSavingScript(`pivot-group-${c.id}`, script, bytes, args, undefined, { fatal: false })
    if ('rejected' in result) return result
    const area = result.out.trim().replaceAll('$', '')
    // Excel puts the machine's absolute directory in the workbook's application metadata.
    const saved = await editParts(result.saved, (name, xml) => name === 'xl/workbook.xml'
      ? xml.replace(/<mc:AlternateContent [^>]*><mc:Choice Requires="x15"><x15ac:absPath [^>]*\/><\/mc:Choice><\/mc:AlternateContent>/, '')
      : undefined)
    const { book } = await readXlsx(saved)
    const ranges = book.pivotCaches?.[0].fields.flatMap((field) => field.group?.range ? [field.group.range] : []) ?? []
    checkGrouping(c.id, group, ranges)
    return { answer: { area, values: shown(book, 'P', area) }, saved }
  }, save, t0)
}
