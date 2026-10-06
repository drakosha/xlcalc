/**
 * Excel на Маке как оракул: книга кладётся в песочницу Excel (там он открывает её без запроса
 * доступа), `excel-read.applescript` открывает, пересчитывает, читает три колонки без ошибок и
 * закрывает без сохранения. Язык интерфейса и локаль расчёта задаются аргументами запуска Excel
 * (`-AppleLanguages (en) -AppleLocale en_US` для en-US, без аргументов системные ru для ru-RU);
 * настройки Excel не меняются.
 */
import { execFileSync, spawn } from 'node:child_process'
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
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

const LOCK = join(homedir(), '.cache/formula-engine-excel.lock')

const alive = (pid: number) => {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}

/**
 * Excel один на машину, а прогоны из разных worktree идут параллельно: перезапуск под другую
 * локаль посреди чужого прогона ломает оба. Блокировка каталогом (mkdir атомарен), в нём pid
 * владельца; брошенную блокировку мёртвого процесса забираем. Ждём до 60 минут.
 */
export async function withExcelLock<T>(fn: () => Promise<T>): Promise<T> {
  mkdirSync(join(homedir(), '.cache'), { recursive: true })
  const until = Date.now() + 60 * 60_000
  let said = false
  for (;;) {
    try {
      mkdirSync(LOCK)
      writeFileSync(join(LOCK, 'pid'), String(process.pid))
      break
    } catch {
      let owner = NaN
      try {
        owner = Number(readFileSync(join(LOCK, 'pid'), 'utf8'))
      } catch {}
      if (Number.isFinite(owner) && !alive(owner)) {
        rmSync(LOCK, { recursive: true, force: true })
        continue
      }
      if (Date.now() > until) throw new Error(`Excel занят процессом ${owner} дольше 60 минут (${LOCK})`)
      if (!said) console.log(`Excel занят процессом ${owner}, жду очереди`)
      said = true
      await sleep(2000)
    }
  }
  const release = () => rmSync(LOCK, { recursive: true, force: true })
  process.once('exit', release)
  try {
    return await fn()
  } finally {
    release()
    process.removeListener('exit', release)
  }
}

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
 * System Events fails with -1719 when a book's window closes while it lists them: then it asks again (three times at most).
 */
export function excelDialogs(): { subrole: string; text: string }[] {
  for (let i = 0; ; i++)
    try {
      return listDialogs()
    } catch (e) {
      if (i >= 2 || !String((e as { stderr?: string }).stderr).includes('(-1719)')) throw e
    }
}

/** The windows that are not books in one line (`subrole «text»; …`), empty when there are none. */
export const dialogText = (): string => excelDialogs().map((d) => `${d.subrole} «${d.text}»`).join('; ')

function listDialogs(): { subrole: string; text: string }[] {
  const script = `tell application "System Events" to tell process "Microsoft Excel"
  set out to ""
  repeat with w in windows
    set sr to ""
    try
      set sr to subrole of w
    end try
    if sr is not "AXStandardWindow" then
      set tx to ""
      repeat with t in (static texts of w)
        try
          set v to value of t
          if v is not missing value then set tx to tx & (v as text) & " "
        end try
      end repeat
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
function parseRead(out: string, n: number): Val[] {
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
}

export function runExcel(book: string, n: number): Val[] {
  if (countWorkbooks() !== 0) throw new Error('в Excel есть открытые книги')
  const dir = join(SANDBOX, `fcov-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    copyFileSync(book, file)
    return parseRead(execFileSync('osascript', [SCRIPT, file, String(n)], { encoding: 'utf8', timeout: 180_000, maxBuffer: 1 << 28 }), n)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** Окно «We found a problem with some content»: Excel не читает книгу целиком, обычно из-за формулы, которую он не принимает при загрузке. */
const REPAIR = /problem with some content|проблем\S* с содержим|ошибка в части содержимого/i

/** Закрывает окно восстановления книги кнопкой No; другие окна не трогает. */
export function declineRepair(): boolean {
  if (!excelDialogs().some((d) => REPAIR.test(d.text))) return false
  osa(`tell application "System Events" to tell process "Microsoft Excel"
  repeat with w in windows
    set tx to ""
    repeat with t in (static texts of w)
      try
        set v to value of t
        if v is not missing value then set tx to tx & (v as text)
      end try
    end repeat
    if tx contains "problem with some content" or tx contains "содержим" then
      repeat with b in (buttons of w)
        if name of b is in {"No", "Нет"} then
          click b
          -- the window is gone: the list of windows is stale
          return
        end if
      end repeat
    end if
  end repeat
end tell`)
  return true
}

/** Как runExcel, но следит за окном восстановления: закрывает его и отдаёт 'rejected', а не ждёт таймаута AppleEvent. */
export async function runExcelWatched(book: string, n: number): Promise<Val[] | 'rejected'> {
  const out = await runScriptWatched(SCRIPT, book, n)
  return out === 'rejected' ? out : parseRead(out, n)
}

/** AppleScript `script` над копией книги в песочнице (аргументы: путь, `n`): его вывод или 'rejected', как в runExcelWatched. */
export async function runScriptWatched(script: string, book: string, n: number): Promise<string | 'rejected'> {
  if (countWorkbooks() !== 0) throw new Error('в Excel есть открытые книги')
  const dir = join(SANDBOX, `fcov-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    copyFileSync(book, file)
    const child = spawn('osascript', [script, file, String(n)])
    let out = ''
    let err = ''
    child.stdout.on('data', (b: Buffer) => (out += b.toString('utf8')))
    child.stderr.on('data', (b: Buffer) => (err += b.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 180_000
    for (;;) {
      const code = await Promise.race([exited, sleep(2000).then(() => undefined)])
      if (code !== undefined) {
        if (code !== 0) throw new Error(`osascript ${code}: ${err.trim().split('\n')[0]}`)
        return out
      }
      if (declineRepair()) {
        child.kill()
        await exited
        await sleep(1500)
        declineRepair()
        if (countWorkbooks() !== 0) osa('tell application "Microsoft Excel" to close every workbook saving no')
        return 'rejected'
      }
      if (Date.now() > until) {
        child.kill()
        await exited
        throw new Error('Excel не ответил за 180 с')
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}
