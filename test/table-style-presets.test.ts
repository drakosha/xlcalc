import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
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
