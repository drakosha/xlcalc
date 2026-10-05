/**
 * Excel на Маке как оракул: книга кладётся в песочницу Excel (там он открывает её без запроса
 * доступа), `excel-read.applescript` открывает, пересчитывает, читает три колонки без ошибок и
 * закрывает без сохранения. Язык интерфейса и локаль расчёта задаются аргументами запуска Excel
 * (`-AppleLanguages (en) -AppleLocale en_US` для en-US, без аргументов системные ru для ru-RU);
 * настройки Excel не меняются.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdirSync, rmSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { ErrCode } from './data.js'
import type { Val } from './formulajs.js'

export type Locale = 'en' | 'ru'

const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const SCRIPT = fileURLToPath(new URL('./excel-read.applescript', import.meta.url))

/** ERROR.TYPE → код ошибки. */
const ERROR_TYPE: Record<number, ErrCode> = { 1: 'NULL', 2: 'DIV0', 3: 'VALUE', 4: 'REF', 5: 'NAME', 6: 'NUM', 7: 'NA', 8: 'GETTING_DATA', 9: 'SPILL', 14: 'CALC' }

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** Процесс Excel и его аргументы; null если не запущен. */
export function excelProcess(): { pid: number; args: string } | null {
  const ps = execFileSync('ps', ['-axo', 'pid=,args='], { encoding: 'utf8' })
  for (const line of ps.split('\n')) {
    const m = /^\s*(\d+)\s+(\/Applications\/Microsoft Excel\.app\/Contents\/MacOS\/Microsoft Excel.*)$/.exec(line)
    if (m) return { pid: Number(m[1]), args: m[2] }
  }
  return null
}

const osa = (script: string, timeout = 60_000) => execFileSync('osascript', ['-e', script], { encoding: 'utf8', timeout }).trim()

/** AppleEvent только запущенному Excel: иначе `tell` запустит его без нужных аргументов. */
function countWorkbooks(): number {
  if (!excelProcess()) throw new Error('Excel не запущен')
  return Number(osa('tell application "Microsoft Excel" to count workbooks'))
}

/** Режим запущенного Excel: en (английский интерфейс и локаль), ru (системные), иначе прочий. */
const localeOf = (args: string): Locale | 'other' => {
  const lang = /-AppleLanguages \(en\)/.test(args)
  const loc = /-AppleLocale en_US/.test(args)
  return lang && loc ? 'en' : !lang && !loc && !/-Apple(Languages|Locale)/.test(args) ? 'ru' : 'other'
}

/**
 * Аргументы запуска: для en-US английский интерфейс и локаль. Язык интерфейса влияет на расчёт, не
 * только на показ: с русским интерфейсом текст "TRUE" не логическое (`IF("TRUE",1,2)` это #VALUE!),
 * а логическое в тексте выходит ИСТИНА; с английским `IF("TRUE",1,2)` это 1.
 */
const LAUNCH: Record<Locale, string[]> = { en: ['--args', '-AppleLanguages', '(en)', '-AppleLocale', 'en_US'], ru: [] }

export function excelVersion(): string {
  return osa('tell application "Microsoft Excel" to get version')
}

/**
 * Excel с нужной локалью: если запущен с другой, при нуле открытых книг выходит и запускается
 * заново (`open -a`, для en-US с `--args -AppleLanguages (en) -AppleLocale en_US`). Открытые книги: отказ.
 */
export async function ensureExcel(locale: Locale): Promise<void> {
  const p = excelProcess()
  if (p && localeOf(p.args) === locale) return closeStartupBook()
  if (p) {
    const n = countWorkbooks()
    if (n !== 0) throw new Error(`в Excel открыто книг: ${n}; перезапуск для локали ${locale} отменён`)
    osa('tell application "Microsoft Excel" to quit')
    for (let i = 0; excelProcess(); i++) {
      if (i > 120) throw new Error('Excel не вышел за 60 с')
      await sleep(500)
    }
  }
  execFileSync('open', ['-a', 'Microsoft Excel', ...LAUNCH[locale]])
  for (let i = 0; ; i++) {
    if (i > 120) throw new Error('Excel не запустился за 60 с')
    await sleep(500)
    const q = excelProcess()
    if (!q) continue
    try {
      countWorkbooks()
      break
    } catch {
      // ещё не отвечает на AppleEvent
    }
  }
  const q = excelProcess()!
  if (localeOf(q.args) !== locale) throw new Error(`Excel запустился не с той локалью: ${q.args}`)
  closeStartupBook()
}

/**
 * При запуске без документа Excel открывает пустую новую книгу («Книга1»): новую книгу без файла и
 * без правок закрываем без сохранения. Любая другая открытая книга: отказ.
 */
function closeStartupBook() {
  if (countWorkbooks() === 1) {
    const [name, saved, path] = osa('tell application "Microsoft Excel" to get {name, saved, path} of workbook 1').split(/,\s*/)
    if (/^(Книга|Book)\d+$/.test(name) && saved === 'true' && !path) osa(`tell application "Microsoft Excel" to close workbook "${name}" saving no`)
  }
  const n = countWorkbooks()
  if (n !== 0) throw new Error(`в Excel открыто книг: ${n} (не пустая новая книга запуска)`)
}

/**
 * Окна Excel, которые не окна книг (модальный диалог и т. п.): подроль и текст. Только чтение
 * через System Events, ничего не нажимается.
 */
export function excelDialogs(): { subrole: string; text: string }[] {
  const script = `tell application "System Events" to tell process "Microsoft Excel"
  set out to ""
  repeat with w in windows
    set sr to ""
    try
      set sr to subrole of w
    end try
    if sr is not "AXStandardWindow" then
      set tx to ""
      try
        repeat with t in (value of every static text of w)
          set tx to tx & (t as text) & " "
        end repeat
      end try
      set out to out & sr & "\t" & tx & linefeed
    end if
  end repeat
  return out
end tell`
  const out = execFileSync('osascript', ['-e', script], { encoding: 'utf8', timeout: 30_000 })
  return out
    .split('\n')
    .filter((l) => l.trim())
    .map((l) => {
      const [subrole, text = ''] = l.split('\t')
      return { subrole, text: text.trim() }
    })
}

/**
 * Значения формул книги (одна часть замера) по строкам; книга копируется в песочницу и удаляется
 * после. Часть небольшая, поэтому ожидание ответа Excel короткое (3 минуты).
 */
export function runExcel(book: string, n: number): Val[] {
  if (countWorkbooks() !== 0) throw new Error('в Excel есть открытые книги')
  const dir = join(SANDBOX, `fcov-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    copyFileSync(book, file)
    const out = execFileSync('osascript', [SCRIPT, file, String(n)], { encoding: 'utf8', timeout: 180_000, maxBuffer: 1 << 28 })
    const [types, nums, strs] = JSON.parse(out) as [number[][], number[][], string[][]]
    if (types.length !== n || nums.length !== n || strs.length !== n) throw new Error(`Excel отдал ${types.length}/${nums.length}/${strs.length} строк из ${n}`)
    return types.map(([t], i): Val => {
      if (t === 0) return { t: 'n', v: nums[i][0] }
      if (t === -1) return { t: 'b', v: true }
      if (t === -2) return { t: 'b', v: false }
      if (t === -3) return { t: 's', v: strs[i][0] }
      if (t > 0) return { t: 'e', v: ERROR_TYPE[t] ?? 'ERROR' }
      return { t: 'other', v: `type ${t}` }
    })
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}
