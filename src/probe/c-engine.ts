/** Вариант C за общим интерфейсом замера: хэши и сессия поверх `c-core`. */
import { inputHash, valuesHash } from './canon'
import { compileBook, ownBook, readValues, setCells, type CState } from './c-core'
import type { ProbeEngine, ProbeResult, ProbeSession } from './engine'
import { VOLATILE } from './functions'
import { walk } from './parser'

const FJ_VERSION = '4.6.1'

export function cEngine(ifs: 'fj' | 'idx'): ProbeEngine {
  const configVersion = `compiled-js;formulajs@${FJ_VERSION};ifs=${ifs};canon15`
  const volatileCells = (st: CState) => {
    const vol = new Set<string>(VOLATILE)
    return st.prep.formulas
      .filter((f) => {
        let hit = false
        walk(f.ast, (n) => {
          if (n.k === 'fn' && vol.has(n.name)) hit = true
        })
        return hit
      })
      .map((f) => `${f.sheet}!${f.addr}`)
  }
  return {
    name: ifs === 'idx' ? 'C' : 'C-fj',
    configVersion,
    compute(book, { clock }): ProbeResult {
      const t0 = performance.now()
      const st = compileBook(book, { clock, ifs })
      const values = readValues(st)
      const total = performance.now() - t0
      return {
        values,
        invalid: st.invalid,
        hash: valuesHash(values, configVersion),
        inputHash: inputHash(book),
        volatileCells: volatileCells(st),
        ms: { ...st.ms, total },
      }
    },
    open(book, { clock }): ProbeSession {
      let st = compileBook(ownBook(book), { clock, ifs })
      return {
        setCells(edits) {
          const r = setCells(st, edits)
          st = r.state
          return r.changed
        },
        values: () => readValues(st),
        hash: () => valuesHash(readValues(st), configVersion),
        close() {},
      }
    },
  }
}

/** Сведения о прототипе для отчёта: неподдержанные функции, размер сгенерированного кода. */
export function cDiagnostics(book: Parameters<typeof compileBook>[0], clock: Parameters<typeof compileBook>[1]['clock']) {
  const st = compileBook(book, { clock, ifs: 'idx' })
  return { unsupported: st.unsupported, generatedChars: st.generatedChars, formulas: st.formulas.length, cycles: st.topo.filter((p) => p < 0).length }
}
