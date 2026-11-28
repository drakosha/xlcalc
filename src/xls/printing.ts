import type { Payload } from '../xlsb/records.js'

export interface PrintMarkup {
  properties: string
  settings: string
}


/** The leading fields of BIFF's Setup record, also present in its shorter historical layouts. */
export function setupFields(payload: Payload) {
  const paper = payload.u16()
  const scale = payload.u16()
  const first = payload.i16()
  const width = payload.u16()
  const height = payload.u16()
  const flags = payload.u16()
  return { paper, scale, first, width, height, flags }
}
