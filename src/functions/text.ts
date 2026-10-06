/** Text functions, assembled from `./text/`. */
import type { FnDef } from '../args.js'
import { functions as astext } from './text/astext.js'
import { functions as numbers } from './text/numbers.js'
import { functions as split } from './text/split.js'
import { functions as strings } from './text/strings.js'

export const functions: Record<string, FnDef> = { ...strings, ...numbers, ...split, ...astext }
