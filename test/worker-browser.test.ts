/**
 * The built package in real browsers (headless Chromium, Firefox and WebKit through `playwright-core`, each skipped
 * when its engine is not found: `CHROME_PATH` or Playwright's own for Chromium, Playwright's own for the others): the
 * pool on Web Workers computes the synthetic book and a session's edits with the hashes of Node's synchronous engine,
 * a hard deadline terminates a worker and the session comes back on a new one, `maxHeapMb` is refused; ru-RU, loaded
 * after the pool started, reaches its workers, which import its module for its messages; the same-thread calculator
 * lets timers run while it computes; in a cross-origin isolated page an abort reaches a formula in progress in a worker.
 */
import { existsSync, mkdirSync, readFileSync } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { tmpdir } from 'node:os'
import { join, normalize } from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, firefox, webkit, type BrowserType } from 'playwright-core'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { compute, createSession } from '../src/index.js'
import { classAEdits, synthBook } from '../src/synth.js'

const root = fileURLToPath(new URL('..', import.meta.url))
const found = (...paths: (string | undefined)[]) => paths.find((p) => p && existsSync(p))
const engines: [string, BrowserType, string | undefined][] = [
  ['Chromium', chromium, found(process.env.CHROME_PATH, chromium.executablePath())],
  ['Firefox', firefox, found(firefox.executablePath())],
  ['WebKit', webkit, found(webkit.executablePath())],
]
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

const page = `<!doctype html><meta charset="utf-8"><script type="module">
const { compute, createInlineCalc, registerLocale } = await import('/dist/index.js')
const { createCalcPool } = await import('/dist/worker/index.js')
const { ruRU } = await import('/dist/locales/ru-RU.js')
const { synthBook, classAEdits } = await import('/dist/synth.js')
const clock = ${JSON.stringify(clock)}
const synth = synthBook({ seed: 1, rows: 340 })
const out = { worker: typeof Worker }
try {
  const pool = createCalcPool({ threads: 2, hardTimeoutMs: 2000 })
  out.pool = (await pool.compute(synth.book, { clock })).hash
  out.eval = (await pool.compute(synth.book, { clock, execution: 'eval' })).hash
  const s = await pool.createSession(synth.book, { clock })
  for (const e of classAEdits(synth, 3, 10)) await s.setCells([e])
  out.edited = (await s.getResult()).hash
  registerLocale(ruRU)
  const ru = await pool.createSession({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { formula: 'ROUND(1.25*2,1)' } } }] }, { clock })
  out.ru = (await ru.explain('S', 'A1')).text
  const slow = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'IF(A1>1,SUM(MMULT(SEQUENCE(1500,1500),SEQUENCE(1500,1500))),A1)' } } }] }
  const t = await pool.createSession(slow, { clock })
  out.killed = await t.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }]).then(() => 'answered', (e) => e.code)
  await t.setCells([{ sheet: 'S', cell: 'A1', input: { value: 0.5 } }])
  out.restored = (await t.getResult()).hash
  try { createCalcPool({ maxHeapMb: 64 }); out.heap = 'accepted' } catch (e) { out.heap = e.message }
  await pool.close()
  let ticks = 0, done = false
  const tick = () => { ticks++; if (!done) setTimeout(tick, 0) }
  setTimeout(tick, 0)
  out.inline = (await createInlineCalc({ sliceMs: 2 }).compute(synthBook({ seed: 1, rows: 1700 }).book, { clock })).hash
  done = true
  out.ticks = ticks
} catch (e) { out.error = String(e && e.stack || e) }
window.__out = out
</script>`

/** A cross-origin isolated page: its pool shares memory with its workers, and an abort reaches a formula in progress. */
const isolatedPage = `<!doctype html><meta charset="utf-8"><script type="module">
const { createCalcPool } = await import('/dist/worker/index.js')
const out = { isolated: globalThis.crossOriginIsolated }
try {
  const pool = createCalcPool({ threads: 1 })
  const heavy = { sheets: [{ name: 'S', cells: { B1: { formula: 'SUMPRODUCT(SEQUENCE(3000)*TRANSPOSE(SEQUENCE(3000)))' } } }] }
  let t = performance.now()
  await pool.compute(heavy)
  out.full = performance.now() - t
  const ac = new AbortController()
  setTimeout(() => ac.abort('stop'), 50)
  t = performance.now()
  out.aborted = await pool.compute(heavy, { signal: ac.signal }).then(() => 'answered', (e) => e.code)
  out.ms = performance.now() - t
  out.after = (await pool.compute({ sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] })).sheets[0].cells.A1.value
  await pool.close()
} catch (e) { out.error = String(e && e.stack || e) }
window.__out = out
</script>`

/** A server of the page and of the built package, with the headers of cross-origin isolation or without them. */
function serve(html: string, isolated: boolean): Server {
  const headers = isolated ? { 'cross-origin-opener-policy': 'same-origin', 'cross-origin-embedder-policy': 'require-corp' } : {}
  return createServer((req, res) => {
    const path = new URL(req.url ?? '/', 'http://x').pathname
    if (path === '/') return res.writeHead(200, { ...headers, 'content-type': 'text/html' }).end(html)
    const file = normalize(join(root, path))
    if (!file.startsWith(join(root, 'dist')) || !existsSync(file)) return res.writeHead(404).end()
    res.writeHead(200, { ...headers, 'content-type': 'text/javascript' }).end(readFileSync(file))
  })
}

const servers: Server[] = []
let base = ''
let isolatedBase = ''

beforeAll(async () => {
  const listen = async (server: Server) => {
    servers.push(server)
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
    return `http://127.0.0.1:${(server.address() as { port: number }).port}/`
  }
  base = await listen(serve(page, false))
  isolatedBase = await listen(serve(isolatedPage, true))
})
afterAll(() => Promise.all(servers.map((server) => new Promise<void>((resolve) => server.close(() => resolve())))))

/** What the page at `url` leaves in `__out`, in a headless browser of `engine`. */
async function pageOut(engine: BrowserType, executablePath: string, url: string): Promise<Record<string, unknown>> {
  // macOS 27 denies Firefox from a terminal its app-data folder (microsoft/playwright#42768): a home of its own, kept (its Library cannot be deleted)
  const home = engine === firefox ? join(tmpdir(), 'xlcalc-firefox-home') : undefined
  if (home) mkdirSync(home, { recursive: true })
  const browser = await engine.launch({ executablePath, headless: true, ...(home ? { env: { ...process.env, CFFIXED_USER_HOME: home } } : {}) })
  try {
    const tab = await browser.newPage()
    await tab.goto(url)
    await tab.waitForFunction(() => (globalThis as unknown as { __out?: unknown }).__out, undefined, { timeout: 90_000 })
    return (await tab.evaluate(() => (globalThis as unknown as { __out: Record<string, unknown> }).__out)) as Record<string, unknown>
  } finally {
    await browser.close()
  }
}

describe.each(engines)('in %s', (_, engine, executablePath) => {
  it.skipIf(!executablePath)('the pool on Web Workers and the same-thread calculator give the hashes of Node', async () => {
    expect(existsSync(join(root, 'dist/worker/thread.js')), 'run npm run build first').toBe(true)
    const out = await pageOut(engine, executablePath!, base)
    const synth = synthBook({ seed: 1, rows: 340 })
    const hash = compute(synth.book, { clock }).hash
    const session = createSession(synth.book, { clock })
    for (const e of classAEdits(synth, 3, 10)) session.setCells([e])
    const slow = compute({ sheets: [{ name: 'S', cells: { A1: { value: 0.5 }, B1: { formula: 'IF(A1>1,SUM(MMULT(SEQUENCE(1500,1500),SEQUENCE(1500,1500))),A1)' } } }] }, { clock }).hash
    expect(out.error).toBeUndefined()
    expect(out).toMatchObject({ worker: 'function', pool: hash, eval: hash, edited: session.getResult().hash, killed: 'hard_timeout', restored: slow, inline: compute(synthBook({ seed: 1, rows: 1700 }).book, { clock }).hash })
    expect(out.ru).toBe('S!A1 = 2,5  ← ОКРУГЛ(1,25*2;1)')
    expect(out.heap).toMatch(/^opts\.maxHeapMb: a thread's memory limit needs the pool's own Node worker_threads/)
    expect(out.ticks).toBeGreaterThan(5)
  }, 120_000)

  it.skipIf(!executablePath)('a cross-origin isolated page: abort() stops one long formula in a worker within milliseconds, and the worker goes on', async () => {
    const out = await pageOut(engine, executablePath!, isolatedBase)
    expect(out.error).toBeUndefined()
    expect(out).toMatchObject({ isolated: true, aborted: 'aborted', after: 2 })
    expect(out.ms).toBeLessThan((out.full as number) / 2)
  }, 120_000)
})
