/** Text functions, assembled from `./text/`. */
import type { FnDef } from '../args.js'
import { functions as astext } from './text/astext.js'
import { functions as baht } from './text/baht.js'
import { functions as codepage } from './text/codepage.js'
import { functions as numbers } from './text/numbers.js'
import { functions as split } from './text/split.js'
import { functions as strings } from './text/strings.js'
import { functions as textsplit } from './text/textsplit.js'
import { functions as value } from './text/value.js'

export const functions: Record<string, FnDef> = { ...strings, ...numbers, ...value, ...split, ...astext, ...baht, ...textsplit, ...codepage }
