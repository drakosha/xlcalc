/**
 * Excel for Windows as a second oracle, on the machine `windows-host` over SSH. Books are written here, as the Mac harness
 * writes them; a job copies them to `C:\xlcalc\work\`, where `xl.ps1` opens each in one Excel through COM, takes its
 * steps and writes its answer; the answers come back by `scp`. `host.ps1` sets the interface language, the regional
 * format and the system locale, holds the lock (one run at a time) and stops Excel left by a job cut short.
 */
import { execFileSync, spawn } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir, hostname } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

export type Tag = 'en-US' | 'ru-RU'

/** `user@host` of the Windows machine. */
export const HOST = process.env.XLWIN_HOST ?? 'user@windows-host'
const BASE = 'C:\\xlcalc'
const HERE = fileURLToPath(new URL('.', import.meta.url))
export const OUT = join(HERE, 'out')
/** One SSH connection shared by the calls of a run: a logon on the machine costs seconds. */
const OPTS = ['-o', 'BatchMode=yes', '-o', 'ConnectTimeout=15', '-o', 'ControlMaster=auto', '-o', `ControlPath=${join(homedir(), '.cache/xlwin-%C')}`, '-o', 'ControlPersist=120']
const OWNER = `${hostname()}:${process.pid}`

/** What the language of a tag means on Windows: Excel's interface language id and the system code page. */
const WANT: Record<Tag, { lcid: number; codepage: number }> = { 'en-US': { lcid: 1033, codepage: 1252 }, 'ru-RU': { lcid: 1049, codepage: 1251 } }

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

const exec = (cmd: string, args: string[], timeout: number) => execFileSync(cmd, args, { encoding: 'utf8', timeout, maxBuffer: 1 << 26 }).trim()

/** A PowerShell command line on the machine (its SSH shell is PowerShell). */
export const ssh = (command: string, timeout = 120_000) => exec('ssh', [...OPTS, HOST, command], timeout)

const psLine = (script: string, args: string) => `powershell -NoProfile -ExecutionPolicy Bypass -File ${BASE}\\bin\\${script} ${args}`

/** `host.ps1 -do <what>`. */
const host = (what: string, arg = '', timeout?: number) => ssh(psLine('host.ps1', `-do ${what}${arg ? ` -arg ${arg}` : ''}`), timeout)

/** The settings a run depends on, as `host.ps1 -do env` reports them with Excel started. */
export interface WinEnv {
  ui: string
  culture: string
  system: string
  codepage: number
  dpi: number
  excelRunning: number
  excelUI: number
  version: string
  build: number
  decimal: string
  /** The thousands separator's code (160, a no-break space, in ru-RU). */
  thousands: number
  list: string
  books: number
}

export const env = (): WinEnv => JSON.parse(host('env', '', 180_000)) as WinEnv

/** The scripts the machine runs, copied from here at the start of a run. */
function syncScripts() {
  ssh(`New-Item -ItemType Directory -Force ${BASE}\\bin, ${BASE}\\work | Out-Null`)
  exec('scp', ['-q', ...OPTS, join(HERE, 'host.ps1'), join(HERE, 'xl.ps1'), `${HOST}:C:/xlcalc/bin/`], 120_000)
}

/** Restarts the machine and waits for SSH (a new system locale needs it). */
async function restart() {
  console.log('windows-host: restarting for the system locale')
  try {
    ssh('Restart-Computer -Force', 30_000)
  } catch {
    // the connection may drop as the machine goes down
  }
  await sleep(30_000)
  for (let i = 0; ; i++) {
    if (i > 60) throw new Error('windows-host did not come back in 10 minutes')
    try {
      ssh('echo up', 20_000)
      return
    } catch {
      await sleep(10_000)
    }
  }
}

/** The machine in the locale wanted (interface, regional format, code page), restarting it when the system locale changes; Excel checked as it starts. */
async function ensureLocale(tag: Tag): Promise<WinEnv> {
  let e = env()
  if (e.ui !== tag || e.culture !== tag || e.system !== tag || e.codepage !== WANT[tag].codepage) {
    const r = host('locale', tag)
    if (r.endsWith('restart')) await restart()
    e = env()
  }
  const bad = [e.ui !== tag && `language ${e.ui}`, e.culture !== tag && `format ${e.culture}`, e.system !== tag && `system locale ${e.system}`, e.codepage !== WANT[tag].codepage && `code page ${e.codepage}`, e.excelUI !== WANT[tag].lcid && `Excel interface ${e.excelUI}`, e.books !== 0 && `${e.books} books open`].filter(Boolean)
  if (bad.length) throw new Error(`windows-host is not in ${tag}: ${bad.join(', ')}`)
  return e
}

/**
 * The lock on the machine (one run at a time, whichever Mac or worktree starts it): taken by `host.ps1 -do lock`;
 * a lock left by a dead process of this Mac is taken over; another owner is waited for up to 60 minutes.
 */
async function lock() {
  const until = Date.now() + 60 * 60_000
  for (let said = false; ; said = true) {
    const r = host('lock', OWNER)
    if (r === 'taken') return
    const owner = /^busy (\S+):(\d+) /.exec(r)
    if (owner && owner[1] === hostname() && !alive(Number(owner[2]))) {
      ssh(`Remove-Item -Recurse -Force ${BASE}\\lock`)
      continue
    }
    if (Date.now() > until) throw new Error(`windows-host has been busy for over 60 minutes: ${r}`)
    if (!said) console.log(`windows-host: ${r}, waiting`)
    await sleep(10_000)
  }
}

const alive = (pid: number) => {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}

/** The metadata of a run's answers: Excel's version and the settings it computed with. */
export const metaOf = (e: WinEnv, tag: Tag) => ({
  version: e.version,
  platform: 'Windows',
  locale: tag,
  ui: `${e.ui} (Excel ${e.excelUI})`,
  format: `${e.culture} (decimal ${JSON.stringify(e.decimal)}, thousands U+${e.thousands.toString(16).toUpperCase().padStart(4, '0')}, list ${JSON.stringify(e.list)})`,
  system: `${e.system} (code page ${e.codepage})`,
  dpi: e.dpi,
})

/** Runs `fn` holding the machine's lock, with the machine in `tag`; Excel left by a job cut short is stopped before and after. */
export async function withWindowsExcel<T>(tag: Tag, fn: (e: WinEnv) => Promise<T>): Promise<T> {
  syncScripts()
  await lock()
  try {
    host('clean')
    return await fn(await ensureLocale(tag))
  } finally {
    host('clean')
    host('unlock', OWNER)
  }
}

/** A step of `xl.ps1` (see there). */
export type Step = { do: string } & Record<string, unknown>

/** A book (a local xlsx file) and its steps. */
export interface Book {
  file: string
  steps: Step[]
}

/** A book's answer: one value per step (and the file a `saveAs` step saved), or why Excel did not open it; undefined when the job stopped before it. */
export type Answer = { answers: unknown[]; seconds: number; saved?: Uint8Array } | { rejected: string; seconds: number } | undefined

let jobs = 0

/**
 * The books in one Excel on the machine. A job that runs past `timeout` (Excel hangs on a window nobody sees) is
 * stopped, Excel with it: the books done keep their answers, the rest are undefined.
 */
export async function runJob(label: string, books: Book[], timeout = 120_000 + 20_000 * books.length): Promise<Answer[]> {
  const name = `${label}-${process.pid}-${++jobs}`
  const local = join(OUT, name)
  const remote = `${BASE}\\work\\${name}`
  rmSync(local, { recursive: true, force: true })
  mkdirSync(local, { recursive: true })
  books.forEach((b, i) => copyFileSync(b.file, join(local, `${i}.xlsx`)))
  writeFileSync(join(local, 'job.json'), JSON.stringify({ books: books.map((b, i) => ({ file: `${i}.xlsx`, steps: b.steps })) }))
  try {
    exec('scp', ['-q', '-r', ...OPTS, local, `${HOST}:C:/xlcalc/work/`], 300_000)
    const timedOut = await new Promise<boolean>((resolve, reject) => {
      const child = spawn('ssh', [...OPTS, HOST, psLine('xl.ps1', remote)])
      let err = ''
      child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
      child.stdout.on('data', () => {})
      const timer = setTimeout(() => {
        child.kill()
        resolve(true)
      }, timeout)
      child.on('close', (code) => {
        clearTimeout(timer)
        if (code === 0) resolve(false)
        else if (code !== null) reject(new Error(`xl.ps1 ${label}: exit ${code}: ${err.trim().split('\n').slice(0, 12).join(' | ')}`))
      })
    })
    if (timedOut) console.log(`${label}: the job ran past ${timeout / 1000} s, Excel stopped: ${host('clean')}`)
    exec('scp', ['-q', '-r', ...OPTS, `${HOST}:C:/xlcalc/work/${name}/out`, local], 300_000)
    const done = join(local, 'out')
    if (!timedOut && JSON.parse(readFileSync(join(done, 'done.json'), 'utf8')).booksLeft !== 0) throw new Error(`${label}: Excel had books open at the end`)
    return books.map((_, i) => {
      if (!existsSync(join(done, `${i}.json`))) return undefined
      const a = JSON.parse(readFileSync(join(done, `${i}.json`), 'utf8')) as Answer
      const saved = join(done, `${i}.saved.xlsx`)
      return a && 'answers' in a && existsSync(saved) ? { ...a, saved: new Uint8Array(readFileSync(saved)) } : a
    })
  } finally {
    // the machine's copy goes with the next `clean`
    if (!process.env.XLWIN_KEEP) rmSync(local, { recursive: true, force: true })
  }
}

/**
 * Items (a part of cases, a case) as books in jobs of up to `per`: `take` stores a book's answer and gives back what is
 * to run again (a part halved, the books after the one a cut-short job hung on); `save` follows every job.
 */
export async function inJobs<T>(label: string, items: T[], build: (t: T, file: string) => Promise<Book>, take: (t: T, a: Answer, cutShort: boolean) => T[], save: () => void, per = 40): Promise<void> {
  const queue = [...items]
  const t0 = Date.now()
  mkdirSync(OUT, { recursive: true })
  for (let job = 1; queue.length; job++) {
    const batch = queue.splice(0, per)
    const books = await Promise.all(batch.map((t, k) => build(t, join(OUT, `${label}-${k}.xlsx`))))
    const answers = await runJob(label, books)
    const stop = answers.findIndex((a) => !a)
    answers.forEach((a, k) => queue.push(...take(batch[k], a, stop >= 0 && k > stop)))
    save()
    console.log(`${label} job ${job}: ${batch.length} books, ${queue.length} left, ${((Date.now() - t0) / 1000).toFixed(0)} s`)
  }
}
