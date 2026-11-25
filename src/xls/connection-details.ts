/** Saved BIFF8 connection definitions. Database properties can span workbook and worksheet records. */
import type { Payload } from '../xlsb/records.js'
import { segmentedString } from './connections.js'
import { readHierarchy, readOlapField, type XlsOlapPivot } from './olap.js'
import { pivotHeader, queryTag } from './pivot-table.js'
import { BOF, EOF, records, unicode, type RecordInfo } from './records.js'
import { sheetRecords } from './worksheet.js'
import type { Workbook } from './workbook.js'

/** A connection and database command as stored in the file. */
export interface XlsDatabaseConnection {
  connectionString?: string
  command?: string
  originalCommand?: string
  /** The OLE DB command type. A cube name is type 1. */
  commandType?: number
  localCubeConnectionString?: string
  drillthroughRows?: number
}

/** A stored parameter value or prompt and its ODBC SQL type. */
export type XlsConnectionParameter = {
  name: string
  sqlType: number
  defaultName: boolean
} & ({ prompt: string } | { value: string | number | boolean })

/** The source and import settings of a saved text query. */
export interface XlsTextConnection {
  sourceFile: string
  delimited: boolean
  prompt: boolean
  startRow: number
  tab: boolean
  space: boolean
  comma: boolean
  semicolon: boolean
  consecutive: boolean
  delimiter?: string
  /** 0 for a double quote, 1 for an apostrophe, 2 for no qualifier. */
  qualifier: number
  decimalSeparator: string
  thousandsSeparator: string
  /** The field types and zero-based character offsets as stored in the text query. */
  fields: { type: number; start: number }[]
}

/** An XLS connection. Its id is the one-based position of its DConn record. */
export interface XlsConnection {
  id: number
  name: string
  /** BIFF DataSourceType, including 1 for ODBC, 4 for Web, 5 for OLE DB, and 6 for text. */
  type: number
  description?: string
  sourceFile?: string
  connectionFile?: string
  pivotCache?: number
  queryTable?: string
  deleted?: boolean
  savePassword?: boolean
  saveData?: boolean
  keepAlive?: boolean
  background?: boolean
  refreshOnLoad?: boolean
  refreshInterval?: number
  useConnectionFile?: boolean
  olap?: boolean
  parameters?: XlsConnectionParameter[]
  database?: XlsDatabaseConnection
  text?: XlsTextConnection
  web?: { url: string; post?: string; editUrl?: string; tables?: string }
}

/** A connection whose record could not be fully read. */
export interface XlsConnectionIssue {
  id: number
  reason: 'invalid'
}

function value(payload: Payload): string {
  return segmentedString(payload, true)
}

function sequence(payload: Payload): string {
  const count = payload.u16()
  if (count * 4 > payload.remaining) payload.fail('connection string sequence is truncated')
  const chunks: string[] = []
  for (let i = 0; i < count; i++) chunks.push(value(payload))
  return chunks.join('')
}

function textQuery(payload: Payload): XlsTextConnection {
  if (payload.u16() !== 0x0805) payload.fail('invalid TxtQry record')
  payload.skip(2)
  const flags = payload.u16()
  payload.skip(2)
  const startRow = payload.i32()
  const delimiters = payload.u8()
  const custom = payload.u16()
  payload.skip(1)
  const count = payload.i32()
  const decimalSeparator = String.fromCharCode(payload.u8())
  const thousandsSeparator = String.fromCharCode(payload.u8())
  if (count < 1 || count > 256) payload.fail('invalid text field count')
  const fields: XlsTextConnection['fields'] = []
  for (let i = 0; i < count; i++) fields.push({ type: payload.u32(), start: payload.u32() })
  return {
    sourceFile: unicode(payload), delimited: Boolean(flags & 2), prompt: Boolean(flags & 16), startRow,
    tab: Boolean(delimiters & 1), space: Boolean(delimiters & 2), comma: Boolean(delimiters & 4),
    semicolon: Boolean(delimiters & 8), consecutive: Boolean(delimiters & 32),
    ...(delimiters & 16 ? { delimiter: String.fromCharCode(custom) } : {}),
    qualifier: delimiters >>> 6, decimalSeparator, thousandsSeparator, fields,
  }
}

function oleDatabase(payload: Payload): XlsDatabaseConnection {
  const drillthroughRows = payload.u32()
  const count = payload.u16()
  if (count > 4) payload.fail('invalid OLE DB connection string count')
  const kinds: number[] = []
  for (let i = 0; i < count; i++) kinds.push(payload.u16())
  payload.skip((4 - count) * 2 + 2)
  const strings = new Map(kinds.map((kind) => [kind, value(payload)]))
  const connectionString = strings.get(0) ?? strings.get(1)
  const localCubeConnectionString = strings.get(2) ?? strings.get(3)
  return {
    ...(connectionString ? { connectionString } : {}),
    ...(localCubeConnectionString ? { localCubeConnectionString } : {}),
    ...(drillthroughRows ? { drillthroughRows } : {}),
  }
}

function parameterValue(payload: Payload): string | number | boolean {
  switch (payload.u16()) {
    case 1: return payload.f64()
    case 2: payload.skip(8); return value(payload)
    case 4: {
      const value = payload.u8()
      payload.skip(7)
      return Boolean(value)
    }
    case 0x0800: {
      const value = payload.u32()
      payload.skip(4)
      return value
    }
    default: return payload.fail('invalid connection parameter type')
  }
}

function readParameter(payload: Payload): XlsConnectionParameter {
  const name = value(payload)
  const binding = payload.u16() & 7
  const sqlType = payload.i16()
  const defaultName = Boolean(payload.u16() & 1)
  if (binding === 0) return { name, sqlType, defaultName, prompt: value(payload) }
  if (binding === 1) return { name, sqlType, defaultName, value: parameterValue(payload) }
  return payload.fail('invalid connection parameter binding')
}

function definition(payload: Payload, connection: XlsConnection): void {
  payload.skip(4)
  const type = payload.u16()
  connection.type = type
  const flags = payload.u16()
  const parameters = payload.u16()
  payload.skip(2)
  const detail = payload.u16()
  const query = payload.u16()
  payload.skip(3)
  const interval = payload.u16()
  payload.skip(8)
  const sourceFile = value(payload)
  const connectionFile = value(payload)
  connection.name = value(payload)
  const description = value(payload)
  value(payload) // SSO identifier
  if (type === 5 && sourceFile) connection.sourceFile = sourceFile
  if (connectionFile) connection.connectionFile = connectionFile
  if (description) connection.description = description
  if (flags & 8) connection.deleted = true
  if (flags & 64) connection.background = true
  if (flags & 128) connection.refreshOnLoad = true
  if (flags & 256) connection.saveData = true
  if (detail & 1) connection.keepAlive = true
  if (interval) connection.refreshInterval = interval
  if (type === 1 || type === 5) {
    if (flags & 1) connection.savePassword = true
    if (flags & 32) connection.useConnectionFile = true
  }
  const tables = type === 4 && flags & 4 ? value(payload) : ''
  if (parameters) {
    const bindings: XlsConnectionParameter[] = []
    for (let i = 0; i < parameters; i++) bindings.push(readParameter(payload))
    connection.parameters = bindings
  }

  if (type === 1) {
    const connectionString = value(payload)
    connection.database = connectionString ? { connectionString } : {}
  } else if (type === 5) {
    connection.database = oleDatabase(payload)
    const commandType = query & 7
    if (commandType) connection.database.commandType = commandType
    if (commandType === 1) connection.olap = true
  } else if (type === 6) connection.text = textQuery(payload)
  else if (type === 4) {
    const url = sequence(payload)
    const post = sequence(payload)
    connection.web = { url, ...(post ? { post } : {}), ...(tables ? { tables } : {}) }
  }
  const command = sequence(payload)
  const originalCommand = sequence(payload)
  const editUrl = sequence(payload)
  if (connection.database && command) connection.database.command = command
  if (type === 1 && connection.database && originalCommand) connection.database.originalCommand = originalCommand
  if (connection.web && editUrl) connection.web.editUrl = editUrl
  switch (payload.u8()) {
    case 0: break
    case 1: connection.queryTable = value(payload); break
    case 2: connection.pivotCache = payload.u16(); break
    default: payload.fail('invalid connection target')
  }
  payload.done()
}

/** Only workbook-level records belong to the DConn catalogue. */
function* globalRecords(workbook: Workbook): Generator<RecordInfo> {
  if (workbook.encoding) return
  for (const record of records(workbook.bytes, 0, workbook.bytes.length)) {
    if (record.type === EOF) return
    if (record.type !== BOF) yield record
  }
}

type DatabaseStrings = Map<number, { command: string; connectionString: string; originalCommand: string }>

function legacyQueries(workbook: Workbook): DatabaseStrings {
  const queries: DatabaseStrings = new Map()
  let cache: number | undefined
  let external = false
  let pending: { cache: number; command: number; connection: number; original: number; strings: string[] } | undefined
  for (const { type, payload } of globalRecords(workbook)) {
    try {
      if (type === 0x00cd && pending) {
        pending.strings.push(unicode(payload))
        const { command, connection, original, strings } = pending
        if (strings.length === command + connection + original) {
          queries.set(pending.cache, {
            command: strings.slice(0, command).join(''),
            connectionString: strings.slice(command, command + connection).join(''),
            originalCommand: strings.slice(command + connection).join(''),
          })
          pending = undefined
        }
        continue
      }
      pending = undefined
      if (type === 0x00d5) {
        external = false
        cache = payload.u16()
      } else if (type === 0x00e3) external = payload.u16() === 2
      else if (type === 0x00dc && external && cache !== undefined) {
        external = false
        const flags = payload.u16()
        const parameters = payload.u16()
        const commandCount = payload.u16()
        const webPost = payload.u16()
        const originalCount = payload.u16()
        const connectionCount = payload.u16()
        const command = flags & 16 ? commandCount : 0
        const original = flags & 32 ? originalCount : 0
        const connection = flags & 8 ? connectionCount : 0
        if (!parameters && !webPost && (flags & 7) !== 4 && command + connection + original > 0) {
          pending = { cache, command, connection, original, strings: [] }
        }
      }
    } catch {
      pending = undefined
    }
  }
  return queries
}

/** The legacy OLE DB strings and cube properties are attached to worksheet views by name. */
function completeViewConnections(workbook: Workbook, connections: XlsConnection[]): XlsOlapPivot[] {
  const byCache = new Map(connections.filter((connection) => connection.pivotCache !== undefined).map((connection) => [connection.pivotCache!, connection]))
  const byQuery = new Map(connections.filter((connection) => connection.queryTable !== undefined).map((connection) => [connection.queryTable!, connection]))
  const olapPivots: XlsOlapPivot[] = []
  for (const sheet of workbook.sheets) {
    if (sheet.kind !== 'worksheet') continue
    const pivots = new Map<string, ReturnType<typeof pivotHeader>>()
    let connection: XlsConnection | undefined
    let olap: XlsOlapPivot | undefined
    let pending: { left: number; local: boolean; chunks: string[] } | undefined
    for (const { type, payload } of sheetRecords(workbook, sheet)) {
      try {
        if (type === 0x00b0) {
          const header = pivotHeader(payload)
          pivots.set(header.name, header)
        } else if (type === 0x0802) {
          connection = undefined
          olap = undefined
          const { pivot, flags, name } = queryTag(payload)
          const view = pivot ? pivots.get(name) : undefined
          const cache = view === undefined ? undefined : workbook.caches[view.cacheIndex]
          connection = pivot ? cache === undefined ? undefined : byCache.get(cache.id) : byQuery.get(name)
          if (pivot && flags & 4 && view && cache) {
            olap = { sheet: sheet.name, name, range: view.range, cache: cache.id, hierarchies: [], fields: [] }
            olapPivots.push(olap)
            if (connection) connection.olap = true
          }
        } else if (type === 0x080d && olap) {
          // The field extensions use positions in this list, including malformed entries.
          const hierarchy = olap.hierarchies.length
          olap.hierarchies.push(null)
          olap.hierarchies[hierarchy] = readHierarchy(payload)
        } else if (type === 0x080f && olap) {
          // Keep the field index even if an optional extension record is truncated.
          const field = olap.fields.length
          olap.fields.push({})
          olap.fields[field] = readOlapField(payload)
        }
        if (pending && type === 0x0804 && connection?.database) {
          payload.skip(4)
          pending.chunks.push(unicode(payload))
          if (--pending.left === 0) {
            const value = pending.chunks.join('')
            if (pending.local) connection.database.localCubeConnectionString ??= value
            else connection.database.connectionString ??= value
            pending = undefined
          }
          continue
        }
        pending = undefined
        if (type === 0x080a && connection?.type === 5) {
          payload.skip(4)
          const flags = payload.u16()
          const left = payload.u16()
          if (left > 0) pending = { left, local: Boolean(flags & 2), chunks: [] }
        }
      } catch {
        pending = undefined
      }
    }
  }
  return olapPivots
}

export function readConnections(workbook: Workbook): { connections: XlsConnection[]; connectionIssues: XlsConnectionIssue[]; olapPivots: XlsOlapPivot[] } {
  const connections: XlsConnection[] = []
  const connectionIssues: XlsConnectionIssue[] = []
  for (const { type, payload } of globalRecords(workbook)) {
    if (type !== 0x0876) continue
    const connection: XlsConnection = { id: connections.length + 1, name: '', type: 0 }
    connections.push(connection)
    try {
      definition(payload, connection)
    } catch {
      connectionIssues.push({ id: connection.id, reason: 'invalid' })
    }
  }
  if (!connections.length && !workbook.caches.some((cache) => cache.olap)) return { connections, connectionIssues, olapPivots: [] }
  const legacy = legacyQueries(workbook)
  for (const connection of connections) {
    const database = connection.database
    const query = connection.pivotCache === undefined ? undefined : legacy.get(connection.pivotCache)
    if (!query || !database) continue
    if (query.command) database.command ??= query.command
    if (query.connectionString) database.connectionString ??= query.connectionString
    if (query.originalCommand) database.originalCommand ??= query.originalCommand
  }
  const olapPivots = completeViewConnections(workbook, connections)
  return { connections, connectionIssues, olapPivots }
}
