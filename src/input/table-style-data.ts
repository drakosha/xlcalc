import { tableKey } from '../tables/model.js'
import { InputError } from '../input-error.js'
import { array, checkBoolean, checkNumber, checkString, given, object, onlyKeys } from './basic.js'
import { checkColor, checkFormat, checkLine, checkStyle } from './styles.js'

/** The palette and custom definitions are workbook metadata, independent of which tables use them. */
export function checkTableStyleData(v: unknown, path: string): void {
  const data = object(v, path, 'table style data')
  onlyKeys(data, path, ['theme', 'custom'])
  given(data, path, {
    theme: (value, at) => {
      const colors = array(value, at, 'the 12 theme colors')
      if (colors.length !== 12) throw new InputError(at, 'expected the 12 theme colors')
      for (let i = 0; i < 12; i++) checkColor(colors[i], `${at}[${i}]`)
    },
    custom: (value, at) => {
      const definitions = object(value, at, 'table styles by name')
      const names = new Set<string>()
      for (const [name, definition] of Object.entries(definitions)) {
        const scope = `${at}[${JSON.stringify(name)}]`
        const key = tableKey(name)
        if (!name || names.has(key)) throw new InputError(scope, 'expected a distinct style name')
        names.add(key)
        const style = object(definition, scope, 'a table style definition')
        onlyKeys(style, scope, ['table', 'pivot', 'elements'])
        given(style, scope, { table: checkBoolean, pivot: checkBoolean })
        const elements = array(style.elements, `${scope}.elements`, 'an array of table style elements')
        for (const [index, value] of elements.entries()) {
          const where = `${scope}.elements[${index}]`
          const element = object(value, where, 'a table style element')
          onlyKeys(element, where, ['type', 'size', 'style', 'inside', 'numberFormat'])
          if (!checkString(element.type, `${where}.type`)) throw new InputError(`${where}.type`, 'expected a non-empty element type')
          if (element.size !== undefined) checkNumber(element.size, `${where}.size`, 1, 0xffffffff, true)
          checkStyle(element.style, `${where}.style`)
          if (element.numberFormat !== undefined) checkFormat(element.numberFormat, `${where}.numberFormat`)
          if (element.inside !== undefined) {
            const inside = object(element.inside, `${where}.inside`, 'inside borders such as { horizontal: { style: "thin" } }')
            onlyKeys(inside, `${where}.inside`, ['horizontal', 'vertical'])
            for (const side of ['horizontal', 'vertical']) if (inside[side] !== undefined) checkLine(inside[side], `${where}.inside.${side}`)
          }
        }
      }
    },
  })
}
