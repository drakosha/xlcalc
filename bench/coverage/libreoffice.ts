/**
 * LibreOffice как пятая колонка: тот же xlsx пересчитывается `soffice --headless` с отдельным
 * профилем en-US (как оракул тестов, `test/oracle.ts`), результаты колонки A листа случаев читаются
 * из XML сохранённой книги.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ErrCode } from './data.js'
import type { Val } from './formulajs.js'

const PROFILE = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>en-US</value></prop></item>
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooLocale" oor:op="fuse"><value>en-US</value></prop></item>
</oor:items>`

export const ERROR_TEXT: Record<string, ErrCode> = {
  '#NULL!': 'NULL',
  '#DIV/0!': 'DIV0',
  '#VALUE!': 'VALUE',
  '#REF!': 'REF',
  '#NAME?': 'NAME',
  '#NUM!': 'NUM',
  '#N/A': 'NA',
  '#SPILL!': 'SPILL',
  '#CALC!': 'CALC',
}

const unxml = (t: string) =>
  t
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&')

export function libreOfficeVersion(): string {
  return execFileSync('soffice', ['--version'], { encoding: 'utf8' }).trim()
}

/** Значения колонки A листа случаев по строкам (индекс с 0). */
export function runLibreOffice(book: string, n: number): Val[] {
  const dir = mkdtempSync(join(tmpdir(), 'fcov-lo-'))
  try {
    const src = join(dir, 'src')
    mkdirSync(src)
    const file = join(src, 'book.xlsx')
    copyFileSync(book, file)
    const profile = join(dir, 'profile')
    mkdirSync(join(profile, 'user'), { recursive: true })
    writeFileSync(join(profile, 'user', 'registrymodifications.xcu'), PROFILE)
    execFileSync('soffice', [`-env:UserInstallation=file://${profile}`, '--headless', '--convert-to', 'xlsx', '--outdir', dir, file], { stdio: 'pipe', timeout: 600_000 })
    const xml = execFileSync('unzip', ['-p', join(dir, 'book.xlsx'), 'xl/worksheets/sheet2.xml'], { encoding: 'utf8', maxBuffer: 1 << 28 })
    const shared = sharedStrings(execFileSync('unzip', ['-p', join(dir, 'book.xlsx'), 'xl/sharedStrings.xml'], { encoding: 'utf8', maxBuffer: 1 << 28 }))
    const out: Val[] = Array.from({ length: n }, () => ({ t: 'other', v: 'missing' }) as Val)
    for (const m of xml.matchAll(/<c r="A(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const row = Number(m[1]) - 1
      const t = /t="(\w+)"/.exec(m[2])?.[1] ?? 'n'
      const v = /<v>([\s\S]*?)<\/v>/.exec(m[3] ?? '')?.[1]
      if (v === undefined || row >= n) continue
      const text = unxml(v)
      out[row] =
        t === 'b' ? { t: 'b', v: text === '1' } : t === 'e' ? { t: 'e', v: ERROR_TEXT[text] ?? 'ERROR' } : t === 'str' || t === 'inlineStr' ? { t: 's', v: text } : t === 's' ? { t: 's', v: shared[Number(text)] } : { t: 'n', v: Number(text) }
    }
    return out
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

function sharedStrings(xml: string): string[] {
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => unxml(t[1])).join(''))
}
