// Computes book.json (beside this script) with the xlcalc whose root module is argv[2], writes the result to argv[3].
// The clock is fixed and `ms` left out, so the same version writes the same file.
import { readFileSync, writeFileSync } from 'node:fs'
import { pathToFileURL } from 'node:url'

const [lib, out] = process.argv.slice(2)
const x = await import(pathToFileURL(lib).href)
// ru-RU is built in up to 0.16; later versions load it as a module
if (x.registerLocale) x.registerLocale((await import(pathToFileURL(lib.replace(/index\.(js|ts)$/, 'locales/ru-RU.$1')).href)).ruRU)
const book = JSON.parse(readFileSync(new URL('book.json', import.meta.url), 'utf8'))
const { ms: _ms, ...result } = x.compute(book, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
writeFileSync(out, JSON.stringify(result, null, 1) + '\n')
console.log(out, Object.keys(result).join(' '))
