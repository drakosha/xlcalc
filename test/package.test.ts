/**
 * The package as a standalone library: no runtime dependencies, the sources import only their own files
 * (no `node:*`, so the code runs outside Node); formula.js is a development dependency of the coverage
 * measurement only; HyperFormula (GPL) is in no dependency. Which functions are whitelisted is checked
 * against Excel in `test/excel.test.ts`.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { CONFIG_VERSION, FUNCTION_NAMES } from '../src/index.js'
import { PACKAGE_NAME } from '../src/config.js'

const root = fileURLToPath(new URL('..', import.meta.url))
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as { name: string } & Record<'dependencies' | 'devDependencies' | 'peerDependencies' | 'optionalDependencies', Record<string, string> | undefined>

/** The functions whose semantics are checked against Excel documentation. */
const SEMANTICS = ['SUM', 'SUMIFS', 'COUNTIFS', 'IF', 'IFERROR', 'XLOOKUP', 'ROUND', 'EOMONTH', 'DATE', 'TEXT']

describe('package', () => {
  it('function whitelist: the functions with semantics tests of their own are in it', () => {
    for (const f of [...SEMANTICS, 'SUMIF']) expect(FUNCTION_NAMES).toContain(f)
  })

  it('the package name in the code matches package.json', () => {
    expect(PACKAGE_NAME).toBe(pkg.name)
    expect(CONFIG_VERSION.startsWith(`${pkg.name}/`)).toBe(true)
  })

  it('formula.js is only the measured object of the coverage measurement: a pinned development dependency', () => {
    expect(pkg.devDependencies?.['@formulajs/formulajs']).toBe('4.6.1')
    expect(CONFIG_VERSION).not.toMatch(/formulajs/)
  })

  it('dependencies: none, HyperFormula nowhere', () => {
    expect(pkg.dependencies ?? {}).toEqual({})
    expect(pkg.peerDependencies ?? {}).toEqual({})
    expect(pkg.optionalDependencies ?? {}).toEqual({})
    expect(Object.keys(pkg.devDependencies ?? {})).not.toContain('hyperformula')
    expect(readFileSync(join(root, 'package-lock.json'), 'utf8')).not.toMatch(/hyperformula/i)
  })

  it('sources import only their own files', () => {
    const src = join(root, 'src')
    const bad: string[] = []
    for (const file of readdirSync(src, { recursive: true, encoding: 'utf8' }).filter((f) => f.endsWith('.ts'))) {
      const text = readFileSync(join(src, file), 'utf8')
      for (const m of text.matchAll(/(?:^|\n)\s*(?:import|export)[^'"\n]*?from\s+'([^']+)'|import\(\s*'([^']+)'\s*\)/g)) {
        const spec = m[1] ?? m[2]
        if (!/^\.\.?\//.test(spec)) bad.push(`${file}: ${spec}`)
      }
    }
    expect(bad).toEqual([])
  })
})
