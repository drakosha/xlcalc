import { InputError } from '../input-error.js'
import { array, checkBoolean, checkNumber, checkString, given, object, onlyKeys } from './basic.js'
import { checkColor, checkStyle } from './styles.js'

/** The palette and custom definitions are workbook metadata, independent of which tables use them. */
export function checkTableStyleData(v: unknown, path: string): void {
  const data = object(v, path, 'table style data')
  onlyKeys(data, path, ['theme', 'custom'])
  given(data, path, {
    theme: (value, at) => {
      const colors = array(value, at, 'the 12 theme colors')
      if (colors.length !== 12) throw new InputError(at, 'expected the 12 theme colors')
      colors.forEach((color, index) => checkColor(color, `${at}[${index}]`))
    },
    custom: (value, at) => {
      const definitions = object(value, at, 'table styles by name')
      for (const [name, definition] of Object.entries(definitions)) {
        const scope = `${at}[${JSON.stringify(name)}]`
        if (!name) throw new InputError(scope, 'expected a style name')
        const style = object(definition, scope, 'a table style definition')
        onlyKeys(style, scope, ['table', 'pivot', 'elements'])
        given(style, scope, { table: checkBoolean, pivot: checkBoolean })
        array(style.elements, `${scope}.elements`, 'an array of table style elements').forEach((value, index) => {
          const where = `${scope}.elements[${index}]`
          const element = object(value, where, 'a table style element')
          onlyKeys(element, where, ['type', 'size', 'style'])
          if (!checkString(element.type, `${where}.type`)) throw new InputError(`${where}.type`, 'expected a non-empty element type')
          if (element.size !== undefined) checkNumber(element.size, `${where}.size`, 1, 0xffffffff, true)
          checkStyle(element.style, `${where}.style`)
        })
      }
    },
  })
}
