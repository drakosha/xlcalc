/** ENCODEURL: text as a URL's query part writes it, UTF-8 bytes in percent signs but for letters, digits and `-_.`. */
import type { FnDef } from '../../args.js'
import { E, type V } from '../../values.js'
import { capped } from './common.js'

// a lone surrogate has no UTF-8 bytes
const LONE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/

function encodeUrl(s: string): V {
  if (LONE.test(s)) return E.VALUE
  return capped(encodeURIComponent(s).replace(/[!'()*~]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`))
}

export const functions: Record<string, FnDef> = {
  ENCODEURL: { args: 's', fn: encodeUrl },
}
