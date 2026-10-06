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
  /** The generated JS text (`eval`), for `link` and for reading. */
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
  return {
    formula(ctx, ast, s, r, c) {
      const b = jsBackend()
      const { code, ...info } = walk(ctx, ast, s, r, c, b)
      return { ...info, source: b.source(code), fn: unlinked }
    },
    link(batch) {
      if (!batch.length) return
      const fns = compileSources(rt, batch.map((f) => f.source!))
      batch.forEach((f, i) => (f.fn = fns[i]))
    },
  }
}
