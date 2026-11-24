/** Optional ECMA-376 preset data and workbook table styles. */
import { strict } from './strict-args.js'
import { checkTableStyleData } from './input/table-style-data.js'
import { setTableStyleFiles, setTableStyleWindow } from './table-style-hook.js'
import { readTableStyles } from './xlsx/read/table-styles.js'
import { writeTableStyles } from './xlsx/write/table-styles.js'
import { tableStyleWindow } from './table-styles/render.js'

export { PRESET_DIFFERENTIAL_FORMATS, PRESET_TABLE_STYLES, type PresetStyleElement } from './table-styles/presets.js'

/** Enables styles of tables and PivotTables in viewports and workbook style definitions in XLSX. */
export const registerTableStyles = /* @__PURE__ */ strict(function registerTableStyles(): void {
  setTableStyleFiles({ read: readTableStyles, check: checkTableStyleData, write: writeTableStyles })
  setTableStyleWindow(tableStyleWindow)
})
