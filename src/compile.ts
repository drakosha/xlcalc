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
  readonly bodies: number
  /** Walks one formula at sheet `s`, row `r`, column `c`; its function is ready after `link`. */
  formula(ctx: CodegenContext, ast: Ast, s: number, r: number, c: number): FormulaCode
  /** Gives the formulas of a batch their functions (`eval`: one `new Function` for the batch). */
  link(batch: FormulaCode[]): void
}

const unlinked = (): Arg => {
  throw new Error('a formula computed before link')
}

const stackOverflow = (e: unknown) => e instanceof RangeError || (e instanceof Error && e.name === 'InternalError' && /too much recursion/i.test(e.message))

type Factory = (parameters: readonly Arg[]) => () => Arg
interface Template { source: string; factory?: Factory; fallback?: () => Factory }

/** Share code, never mutable formula storage; cap the retained template catalogue at 128 bodies and 64 KiB of keys. */
export function makeCompiler(rt: Runtime, execution: Execution): Compiler {
  const templates = new Map<string, Template>()
  const pending = new WeakMap<() => Arg, { template: Template; parameters: Arg[] }>()
  let bytes = 0, bodies = 0
  const closureFactory = (ctx: CodegenContext, ast: Ast, s: number, r: number, c: number): Factory => {
    const code = walk(ctx, ast, s, r, c, closureBackend(rt, [])).code
    return (parameters) => () => code(parameters)
  }
  return {
    get bodies() { return bodies },
    formula(ctx, ast, s, r, c) {
      const parameters: Arg[] = [], b = jsBackend(parameters)
      const { code, ...info } = walk(ctx, ast, s, r, c, b)
      const source = b.source(code)
      // LET owns mutable bindings; closure LETs retain a separate tree for each formula.
      if (execution === 'closures' && source.includes('=M(')) {
        bodies++
        return { ...info, fn: walk(ctx, ast, s, r, c, closureBackend(rt)).code }
      }
      let template = templates.get(source)
      if (!template) {
        bodies++
        template = { source }
        if (execution === 'closures') template.factory = closureFactory(ctx, ast, s, r, c)
        else template.fallback = () => closureFactory(ctx, ast, s, r, c)
        if (source.length * 2 <= 65536) {
          while (templates.size >= 128 || bytes + source.length * 2 > 65536) {
            const oldest = templates.keys().next().value!
            templates.delete(oldest); bytes -= oldest.length * 2
          }
          templates.set(source, template); bytes += source.length * 2
        }
      }
      if (template.factory) return { ...info, fn: template.factory(parameters) }
      const f: FormulaCode = { ...info, fn: () => unlinked() }
      pending.set(f.fn, { template, parameters })
      return f
    },
    link(batch) {
      const missing = [...new Set(batch.flatMap((f) => {
        const p = pending.get(f.fn)
        return p && !p.template.factory ? [p.template] : []
      }))]
      if (missing.length) {
        try {
          const factories = compileSources<Factory>(rt, missing.map((t) => t.source))
          missing.forEach((t, i) => { t.factory = factories[i] })
        } catch (e) {
          if (!stackOverflow(e)) throw e
          for (const t of missing) {
            try { t.factory = compileSources<Factory>(rt, [t.source])[0] }
            catch (error) { if (!stackOverflow(error)) throw error; t.factory = t.fallback!() }
          }
        }
        for (const t of missing) t.fallback = undefined
      }
      for (const f of batch) {
        const p = pending.get(f.fn)
        if (p) { pending.delete(f.fn); f.fn = p.template.factory!(p.parameters) }
        f.source = undefined
      }
    },
  }
}
