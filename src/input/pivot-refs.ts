/** PivotTables named by an operation (`PivotRef`): each a sheet's name and a PivotTable's, checked for shape. */
import type { PivotRef } from '../contract.js'
import { array, checkString, object, onlyKeys } from './basic.js'

export function pivotRefs(v: unknown, path: string): PivotRef[] {
  return array(v, path, 'an array of PivotTables as { sheet, name }').map((x, i) => {
    const at = `${path}[${i}]`
    const o = object(x, at, 'a PivotTable as { sheet, name }')
    onlyKeys(o, at, ['sheet', 'name'])
    return { sheet: checkString(o.sheet, `${at}.sheet`), name: checkString(o.name, `${at}.name`) }
  })
}
