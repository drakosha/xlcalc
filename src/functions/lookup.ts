/** Lookup and reference functions, assembled from `./lookup/`. */
import type { FnDef } from '../args.js'
import { functions as order } from './lookup/order.js'
import { functions as reference } from './lookup/reference.js'
import { functions as reshape } from './lookup/reshape.js'
import { functions as tables } from './lookup/tables.js'
import { functions as xlookup } from './lookup/xlookup.js'

export const functions: Record<string, FnDef> = { ...reference, ...tables, ...xlookup, ...reshape, ...order }
