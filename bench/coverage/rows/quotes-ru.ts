/**
 * The set `quotes` again, held out for Excel for Windows in ru-RU: the en-US pastes of both Excels read a field begun by
 * a quote to its closing quote, and a paste of copied text in ru-RU on Windows kept its quotes (`docs/entry.md`), so the
 * ru-RU paste is measured on its own. The prediction is the engine's, which reads quotes alike in every locale.
 */
export { buildRowCases } from './quotes.js'
