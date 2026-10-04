/**
 * LibreOffice как оракул в тестах (не в рантайме): книга пишется в xlsx через ExcelJS без кэша
 * значений, `soffice --headless` пересчитывает и сохраняет её в xlsx с кэшем результатов, который
 * читается обратно (без форматов и локали).
 */
import ExcelJS from 'exceljs'
import { execFileSync } from 'node:child_process'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { BookInput, CellValue } from '../src/contract'
import { dateToSerial, serialToUtcMs } from '../src/model'

const PROFILE = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>en-US</value></prop></item>
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooLocale" oor:op="fuse"><value>en-US</value></prop></item>
</oor:items>`

/** Функции, которые в OOXML пишутся с префиксом `_xlfn.`. */
const XLFN = ['XLOOKUP']

export function hasSoffice(): boolean {
  try {
    execFileSync('soffice', ['--version'], { stdio: 'pipe' })
    return true
  } catch {
    return false
  }
}

const ERRORS: Record<string, CellValue> = {
  '#DIV/0!': { error: 'DIV0' },
  '#N/A': { error: 'NA' },
  '#NAME?': { error: 'NAME' },
  '#NUM!': { error: 'NUM' },
  '#REF!': { error: 'REF' },
  '#VALUE!': { error: 'VALUE' },
}

const unxml = (t: string) => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')

/**
 * Кэш результатов листа из xlsx LibreOffice. XML читается напрямую: ExcelJS теряет результат формулы,
 * равный 0 или FALSE.
 */
function sheetResults(xml: string): Map<string, CellValue> {
  const out = new Map<string, CellValue>()
  for (const m of xml.matchAll(/<c r="([A-Z]+[0-9]+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const t = /t="(\w+)"/.exec(m[2])?.[1] ?? 'n'
    const v = /<v>([\s\S]*?)<\/v>/.exec(m[3] ?? '')?.[1]
    if (v === undefined) continue
    const text = unxml(v)
    out.set(m[1], t === 'b' ? text === '1' : t === 'e' ? (ERRORS[text] ?? { error: 'ERROR' }) : t === 'str' || t === 'inlineStr' ? text : Number(text))
  }
  return out
}

/** Результаты формул листа `sheet` по адресу; null если soffice нет. */
export async function oracleValues(book: BookInput, sheet: string): Promise<Map<string, CellValue> | null> {
  if (!hasSoffice()) return null
  const dir = mkdtempSync(join(tmpdir(), 'wb1-oracle-'))
  try {
    const wb = new ExcelJS.Workbook()
    for (const s of book.sheets) {
      const ws = wb.addWorksheet(s.name)
      for (const [addr, c] of Object.entries(s.cells)) {
        const cell = ws.getCell(addr)
        if (c.formula !== undefined) {
          let f = c.formula
          for (const fn of XLFN) f = f.replace(new RegExp(`\\b${fn}\\(`, 'g'), `_xlfn.${fn}(`)
          cell.value = { formula: f } as ExcelJS.CellFormulaValue
        } else if (c.value !== null && typeof c.value === 'object') {
          cell.value = new Date(serialToUtcMs(dateToSerial(c.value.date)))
          cell.numFmt = 'yyyy-mm-dd'
        } else if (c.value !== undefined) cell.value = c.value
      }
    }
    const src = join(dir, 'src')
    mkdirSync(src)
    const file = join(src, 'book.xlsx')
    await wb.xlsx.writeFile(file)
    // LibreOffice пересчитывает книгу без кэша значений и сохраняет её с результатами
    // отдельный профиль с локалью en-US: иначе TEXT берёт десятичную запятую системы
    const profile = join(dir, 'profile')
    mkdirSync(join(profile, 'user'), { recursive: true })
    writeFileSync(join(profile, 'user', 'registrymodifications.xcu'), PROFILE)
    execFileSync('soffice', [`-env:UserInstallation=file://${profile}`, '--headless', '--convert-to', 'xlsx', '--outdir', dir, file], { stdio: 'pipe' })
    const idx = book.sheets.findIndex((x) => x.name === sheet) + 1
    const xml = execFileSync('unzip', ['-p', join(dir, 'book.xlsx'), `xl/worksheets/sheet${idx}.xml`], { encoding: 'utf8', maxBuffer: 1 << 28 })
    return sheetResults(xml)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}
