import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'
import { libUrl, TARGET } from './fixtures/target.js'

describe('optional table style presets', () => {
  it('has all 144 ECMA presets with a differential format for each region', async () => {
    const { PRESET_DIFFERENTIAL_FORMATS: formats, PRESET_TABLE_STYLES: styles } =
      await import(libUrl('table-styles.js')) as typeof import('../src/table-styles.js')
    expect(Object.keys(styles)).toHaveLength(144)
    expect(styles.TableStyleLight14).toBeDefined()
    expect(styles.TableStyleDark9).toBeDefined()
    expect(Object.values(styles).every((regions) => regions.every(([, size, index]) => size >= 1 && !!formats[index]))).toBe(true)
  })

  it('keeps the supplied theme and custom definitions in a session layout', () => {
    const tableStyles = {
      theme: Array.from({ length: 12 }, () => '#112233'),
      custom: { Accent: { table: true, elements: [{ type: 'headerRow', style: { font: { bold: true } } }] } },
    }
    const session = createSession({ sheets: [{ name: 'S', cells: {} }], tableStyles })
    try {
      expect(session.getLayout().tableStyles).toEqual(tableStyles)
    } finally {
      session.close()
    }
    const malformed = { ...tableStyles, theme: ['#112233'] }
    expect(() => createSession({ sheets: [{ name: 'S', cells: {} }], tableStyles: malformed })).toThrow('12 theme colors')
  })

  it('the root and core imports do not include the preset module', async () => {
    for (const entry of ['index.js', 'core.js']) {
      const file = TARGET === 'src' ? entry.replace(/\.js$/, '.ts') : entry
      const bundle = await build({
        stdin: { contents: `export * from ${JSON.stringify(fileURLToPath(libUrl(file)))}`, resolveDir: fileURLToPath(new URL('..', import.meta.url)) },
        bundle: true, logLevel: 'silent', platform: 'browser', write: false, metafile: true,
      })
      expect(Object.keys(bundle.metafile.inputs).some((file) => /table-styles\.(?:ts|js)$|table-styles\/presets/.test(file))).toBe(false)
    }
  })
})
