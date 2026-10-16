/**
 * Typed entry without the engine (the `./entry` subpath, the same functions in the root): what text typed into a cell
 * is as Excel reads it in a locale (`parseEntry`), whether a text needs an apostrophe to stay text (`needsApostrophe`),
 * and the text a value opens with for editing (`entryText`), for a grid that writes what it types and shows what it
 * edits without loading the computation. The functions' arguments come by their declared kinds (`function-arity.ts`),
 * not their implementations; a locale other than en-US is given as an object (`ruRU` of `xlcalc/locales/ru-RU`).
 */
import { aritiesOf, useArities } from './arity.js'
import * as entry from './entry/index.js'
import * as entryTexts from './entry/text.js'
import { FUNCTION_ARITY } from './function-arity.js'
import { strict } from './strict-args.js'

// without the engine, the check of a formula's calls counts by the functions' declared arguments
useArities(aritiesOf(FUNCTION_ARITY))

export const parseEntry = /* @__PURE__ */ strict(entry.parseEntry)
export const entryText = /* @__PURE__ */ strict(entryTexts.entryText)
export const needsApostrophe = /* @__PURE__ */ strict(entryTexts.needsApostrophe)
