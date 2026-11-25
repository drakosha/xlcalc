import { strict } from './strict-args.js'
import type { ReadXlsxOptions, XlsxRead } from './xlsx/types.js'
import { readConnections, type XlsConnection, type XlsConnectionIssue } from './xls/connection-details.js'
import { readXls } from './xls/read.js'
import type { XlsOlapPivot } from './xls/olap.js'

export type { XlsConnection, XlsConnectionIssue, XlsConnectionParameter, XlsDatabaseConnection, XlsTextConnection } from './xls/connection-details.js'

export type { XlsOlapHierarchy, XlsOlapPivot } from './xls/olap.js'

/** The workbook and the definitions of the external connections stored in its BIFF records. */
export interface XlsConnectionsRead extends XlsxRead {
  connections: XlsConnection[]
  connectionIssues: XlsConnectionIssue[]
  olapPivots: XlsOlapPivot[]
}

/** Reads an XLS workbook and its saved connection definitions for inspection. */
async function readXlsWithConnections(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsConnectionsRead> {
  const { read, workbook } = await readXls(data, opts)
  return { ...read, ...readConnections(workbook) }
}

const read = /* @__PURE__ */ strict(readXlsWithConnections)
export { read as readXlsWithConnections }
