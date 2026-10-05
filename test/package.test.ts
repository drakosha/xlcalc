/**
 * Пакет как отдельная библиотека: белый список функций закреплён, исходники импортируют только свои
 * файлы, `node:*` и `@formulajs/formulajs`; HyperFormula (GPL) нет ни в одной зависимости.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { CONFIG_VERSION, FUNCTION_NAMES } from '../src/index.js'
import { FORMULAJS_VERSION } from '../src/config.js'

const root = fileURLToPath(new URL('..', import.meta.url))
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as Record<'dependencies' | 'devDependencies' | 'peerDependencies' | 'optionalDependencies', Record<string, string> | undefined>

/** Список семантики шага 1 (the app, `docs/design-workbook.md`). */
const SEMANTICS = ['SUM', 'SUMIFS', 'COUNTIFS', 'IF', 'IFERROR', 'XLOOKUP', 'ROUND', 'EOMONTH', 'DATE', 'TEXT']

describe('пакет', () => {
  it('белый список функций', () => {
    expect(FUNCTION_NAMES).toEqual(['COUNTIFS', 'DATE', 'EOMONTH', 'FALSE', 'IF', 'IFERROR', 'NOW', 'RAND', 'RANDBETWEEN', 'ROUND', 'SUM', 'SUMIF', 'SUMIFS', 'TEXT', 'TODAY', 'TRUE', 'XLOOKUP'])
    for (const f of [...SEMANTICS, 'SUMIF']) expect(FUNCTION_NAMES).toContain(f)
  })

  it('версия formulajs в конфигурации совпадает с зависимостью', () => {
    expect(pkg.dependencies?.['@formulajs/formulajs']).toBe(FORMULAJS_VERSION)
    expect(CONFIG_VERSION).toContain(`formulajs@${FORMULAJS_VERSION}`)
  })

  it('зависимости: только formulajs, HyperFormula нигде', () => {
    expect(Object.keys(pkg.dependencies ?? {})).toEqual(['@formulajs/formulajs'])
    expect(pkg.peerDependencies ?? {}).toEqual({})
    expect(pkg.optionalDependencies ?? {}).toEqual({})
    expect(Object.keys(pkg.devDependencies ?? {})).not.toContain('hyperformula')
    expect(readFileSync(join(root, 'package-lock.json'), 'utf8')).not.toMatch(/hyperformula/i)
  })

  it('исходники не импортируют ничего вне пакета', () => {
    const src = join(root, 'src')
    const allowed = new Set(Object.keys(pkg.dependencies ?? {}))
    const bad: string[] = []
    for (const file of readdirSync(src, { recursive: true, encoding: 'utf8' }).filter((f) => f.endsWith('.ts'))) {
      const text = readFileSync(join(src, file), 'utf8')
      for (const m of text.matchAll(/(?:^|\n)\s*(?:import|export)[^'"\n]*?from\s+'([^']+)'|import\(\s*'([^']+)'\s*\)/g)) {
        const spec = m[1] ?? m[2]
        if (spec.startsWith('./')) continue
        if (spec.startsWith('node:')) continue
        const name = spec.startsWith('@') ? spec.split('/').slice(0, 2).join('/') : spec.split('/')[0]
        if (allowed.has(name)) continue
        bad.push(`${file}: ${spec}`)
      }
    }
    expect(bad).toEqual([])
  })
})
