/** Text functions, assembled from `./text/`. */
import type { FnDef } from '../args.js'
import { functions as astext } from './text/astext.js'
import { functions as baht } from './text/baht.js'
import { functions as codepage } from './text/codepage.js'
import { functions as eastAsian } from './text/east-asian.js'
import { functions as filterxml } from './text/filterxml.js'
import { functions as numbers } from './text/numbers.js'
import { functions as regex } from './text/regex.js'
import { functions as split } from './text/split.js'
import { functions as strings } from './text/strings.js'
import { functions as textsplit } from './text/textsplit.js'
import { functions as value } from './text/value.js'
import { functions as web } from './text/web.js'

export const functions: Record<string, FnDef> = { ...strings, ...numbers, ...value, ...split, ...astext, ...baht, ...textsplit, ...codepage, ...eastAsian, ...web, ...regex, ...filterxml }
