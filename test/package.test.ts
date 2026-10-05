/**
 * The package as a standalone library: the function whitelist is pinned, the sources import only their own
 * files and `@formulajs/formulajs` (no `node:*`, so the code runs outside Node); HyperFormula (GPL) is in no
 * dependency.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { CONFIG_VERSION, FUNCTION_NAMES } from '../src/index.js'
import { FORMULAJS_VERSION, PACKAGE_NAME } from '../src/config.js'

const root = fileURLToPath(new URL('..', import.meta.url))
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as { name: string } & Record<'dependencies' | 'devDependencies' | 'peerDependencies' | 'optionalDependencies', Record<string, string> | undefined>

/** The functions whose semantics are checked against Excel documentation. */
const SEMANTICS = ['SUM', 'SUMIFS', 'COUNTIFS', 'IF', 'IFERROR', 'XLOOKUP', 'ROUND', 'EOMONTH', 'DATE', 'TEXT']

describe('package', () => {
  it('function whitelist', () => {
    expect(FUNCTION_NAMES).toEqual(['COUNTIFS', 'DATE', 'EOMONTH', 'FALSE', 'IF', 'IFERROR', 'NOW', 'RAND', 'RANDBETWEEN', 'ROUND', 'SUM', 'SUMIF', 'SUMIFS', 'TEXT', 'TODAY', 'TRUE', 'XLOOKUP'])
    for (const f of [...SEMANTICS, 'SUMIF']) expect(FUNCTION_NAMES).toContain(f)
  })

  it('the package name in the code matches package.json', () => {
    expect(PACKAGE_NAME).toBe(pkg.name)
    expect(CONFIG_VERSION.startsWith(`${pkg.name}/`)).toBe(true)
  })

  it('the formulajs version in the configuration matches the dependency', () => {
    expect(pkg.dependencies?.['@formulajs/formulajs']).toBe(FORMULAJS_VERSION)
    expect(CONFIG_VERSION).toContain(`formulajs@${FORMULAJS_VERSION}`)
  })

  it('dependencies: only formulajs, HyperFormula nowhere', () => {
    expect(Object.keys(pkg.dependencies ?? {})).toEqual(['@formulajs/formulajs'])
    expect(pkg.peerDependencies ?? {}).toEqual({})
    expect(pkg.optionalDependencies ?? {}).toEqual({})
    expect(Object.keys(pkg.devDependencies ?? {})).not.toContain('hyperformula')
    expect(readFileSync(join(root, 'package-lock.json'), 'utf8')).not.toMatch(/hyperformula/i)
  })

  it('sources import only their own files and formulajs', () => {
    const src = join(root, 'src')
    const allowed = new Set(Object.keys(pkg.dependencies ?? {}))
    const bad: string[] = []
    for (const file of readdirSync(src, { recursive: true, encoding: 'utf8' }).filter((f) => f.endsWith('.ts'))) {
      const text = readFileSync(join(src, file), 'utf8')
      for (const m of text.matchAll(/(?:^|\n)\s*(?:import|export)[^'"\n]*?from\s+'([^']+)'|import\(\s*'([^']+)'\s*\)/g)) {
        const spec = m[1] ?? m[2]
        if (spec.startsWith('./')) continue
        const name = spec.startsWith('@') ? spec.split('/').slice(0, 2).join('/') : spec.split('/')[0]
        if (allowed.has(name)) continue
        bad.push(`${file}: ${spec}`)
      }
    }
    expect(bad).toEqual([])
  })
})
