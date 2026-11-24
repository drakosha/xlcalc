import { expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { pool } from './fixtures/pool.js'
import { libUrl } from './fixtures/target.js'

const { registerTableStyles } = await import(libUrl('table-styles.js')) as typeof import('../src/table-styles.js')

it('loads optional table styles in an existing pool and edits them for the gallery', async () => {
  const book: BookInput = {
    sheets: [{
      name: 'S',
      cells: { A1: { value: 'Item' }, A2: { value: 1 } },
      tables: [{ name: 'Items', range: 'A1:A3', columns: [{ name: 'Item' }] }],
    }],
  }
  const calc = await pool({ threads: 1 })
  try {
    const session = await calc.createSession(book)
    try {
      const original = await session.getViewport('S', 'A1:A3')
      expect(original.resolvedStyles).toBeUndefined()
      registerTableStyles()
      const styled = await session.getViewport('S', 'A1:A3')
      expect(styled.resolvedStyles?.[styled.cells.A1.resolvedStyle!]?.fill?.color).toBe('#156082')
      const changed = await session.editStructure({ op: 'setTableStyle', table: 'items', style: {} })
      expect(changed.ok).toBe(true)
      const plain = await session.getViewport('S', 'A1:A3')
      expect(plain.resolvedStyles).toBeUndefined()
      expect((await session.undo())?.call).toBe('editStructure')
      const restored = await session.getViewport('S', 'A1:A3')
      expect(restored.resolvedStyles).toEqual(styled.resolvedStyles)
    } finally {
      await session.close()
    }
  } finally {
    await calc.close()
  }
})
