/**
 * Compilation of formulas: one walk over each formula's AST (`exec/walk.ts`) gives its references and
 * flags and its code in the execution mode. `closures` (`exec/closures.ts`) builds a tree of closures
 * over the runtime; `eval` (`exec/js.ts`) generates JS text, and a batch of formulas is compiled by one
 * `new Function`. Both modes give the same values.
 */
import type { Arg } from './args.js'
import type { Execution } from './contract.js'
import { closureBackend } from './exec/closures.js'
import { compileSources, jsBackend } from './exec/js.js'
import { walk, type CodegenContext, type FormulaInfo } from './exec/walk.js'
import type { Runtime } from './runtime.js'
import type { Ast } from './syntax.js'

export { PRELUDE } from './exec/js.js'
export type { CodegenContext } from './exec/walk.js'

/** A compiled formula: what it reads, its flags and its function. */
export interface FormulaCode extends FormulaInfo {
  /** Temporary generated JS text (`eval`), released after `link`. Diagnostics regenerate it on demand. */
  source?: string
  fn: () => Arg
}

export interface Compiler {
  /** Walks one formula at sheet `s`, row `r`, column `c`; its function is ready after `link`. */
  formula(ctx: CodegenContext, ast: Ast, s: number, r: number, c: number): FormulaCode
  /** Gives the formulas of a batch their functions (`eval`: one `new Function` for the batch). */
  link(batch: FormulaCode[]): void
}

const unlinked = (): Arg => {
  throw new Error('a formula computed before link')
}

const stackOverflow = (e: unknown) => e instanceof RangeError || (e instanceof Error && e.name === 'InternalError' && /too much recursion/i.test(e.message))

export function makeCompiler(rt: Runtime, execution: Execution): Compiler {
  if (execution === 'closures') {
    const b = closureBackend(rt)
    return {
      formula(ctx, ast, s, r, c) {
        const { code, ...info } = walk(ctx, ast, s, r, c, b)
        return { ...info, fn: code }
      },
      link() {},
    }
  }
  // Retain the AST only until link, for JS engines whose parser cannot compile a deep expression.
  const fallbacks = new WeakMap<() => Arg, () => () => Arg>()
  return {
    formula(ctx, ast, s, r, c) {
      const b = jsBackend()
      const { code, ...info } = walk(ctx, ast, s, r, c, b)
      const f = { ...info, source: b.source(code), fn: () => unlinked() }
      fallbacks.set(f.fn, () => walk(ctx, ast, s, r, c, closureBackend(rt)).code)
      return f
    },
    link(batch) {
      if (!batch.length) return
      let fns: (() => Arg)[]
      try {
        fns = compileSources(rt, batch.map((f) => f.source!))
      } catch (e) {
        if (!stackOverflow(e)) throw e
        fns = batch.map((f) => {
          try {
            return compileSources(rt, [f.source!])[0]
          } catch (e) {
            if (!stackOverflow(e)) throw e
            return fallbacks.get(f.fn)?.() ?? f.fn
          }
        })
      }
      batch.forEach((f, i) => {
        fallbacks.delete(f.fn)
        f.fn = fns[i]
        f.source = undefined
      })
    },
  }
}
