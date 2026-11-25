import type { Payload } from '../xlsb/records.js'

export function segmentedString(payload: Payload, capture: boolean): string {
  let remaining = payload.u32()
  const chunks: string[] = []
  while (remaining) {
    const count = payload.u16()
    if (!count || count > remaining || count > 512) payload.fail('invalid connection string segment')
    const wide = payload.u8() & 1
    if (capture) chunks.push(wide ? payload.wide(count) : payload.narrow(count))
    else payload.skip(count * (wide ? 2 : 1))
    remaining -= count
  }
  return chunks.join('')
}

export interface ConnectedCache {
  id: number
  olap: boolean
}

/** The SXStreamID and cube command flag of a DConn attached to a database PivotTable. */
export function connectedCache(payload: Payload): ConnectedCache | undefined {
  try {
    payload.skip(4)
    const provider = payload.u16()
    const flags = payload.u16()
    const parameterCount = payload.u16()
    if ((flags & 0x18) || parameterCount || provider === 4 || provider === 6) return undefined
    payload.skip(4)
    const commandType = payload.u16() & 7
    payload.skip(13)
    // Source files, name, description and SSO identifier, followed by provider-specific properties.
    for (let i = 0; i < 5; i++) segmentedString(payload, false)
    if (provider === 1) segmentedString(payload, false)
    else if (provider === 5) {
      payload.skip(4)
      const connectionStrings = payload.u16()
      if (connectionStrings) return undefined
      payload.skip(10)
    }
    // Current SQL, original SQL and the web editor address are three string sequences.
    for (let i = 0; i < 3; i++) {
      const count = payload.u16()
      for (let j = 0; j < count; j++) segmentedString(payload, false)
    }
    if (payload.remaining !== 3 || payload.u8() !== 2) return undefined
    return { id: payload.u16(), olap: provider === 5 && commandType === 1 }
  } catch {
    // An unreadable connection description can be left out without losing the saved cache.
    return undefined
  }
}
