/**
 * Formulas of the old kind (no dynamic array metadata) that Excel 365 computes with implicit intersection: where
 * such a formula puts several cells, an array or a function that can answer one in a place that takes one value
 * (the cell itself, an operator, a one-value argument), Excel takes the value in the formula's row or column
 * (and shows `@` there), while the engine computes a dynamic array. Arguments a function takes whole (SUM's
 * ranges) do not intersect; an argument passed on to the answer (IF's values) is taken as the call is.
 */
import type { NameInput } from '../../contract.js'
import { sortAddresses } from '../../model.js'
import type { Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { MODES } from '../formulas/modes.js'
import type { Issues } from '../issues.js'

type Place = 'one' | 'whole'

/** The mode of argument `i` of a function by its `MODES` entry; `w` for a function the table does not have. */
function modeAt(spec: string | undefined, i: number): string {
  if (!spec) return 'w'
  const [kinds, rest = ''] = spec.split('/')
  const rep = parseInt(rest, 10) || 0
  const fixed = kinds.length - rep
  return i < kinds.length ? kinds[i] : rep ? kinds[fixed + ((i - fixed) % rep)] : 'w'
}

/** Names that reference several cells, by upper-case name. */
function areaNames(names: NameInput[]): Set<string> {
  const out = new Set<string>()
  for (const n of names) {
    try {
      const ast = parse(n.ref)
      if (ast.k === 'ref' && (ast.ref.kind !== 'cell' || ast.ref.spill)) out.add(n.name.toUpperCase())
    } catch {
      // a name that is not a reference is not counted
    }
  }
  return out
}

function intersects(ast: Ast, areas: Set<string>): boolean {
  const go = (n: Ast, at: Place): boolean => {
    switch (n.k) {
      case 'ref':
        return at === 'one' && (n.ref.kind !== 'cell' || n.ref.spill === true)
      case 'name':
        return at === 'one' && areas.has(n.name.toUpperCase())
      case 'arr':
        return at === 'one' && n.rows.length * n.rows[0].length > 1
      case 'fn': {
        const spec = MODES[n.name]
        if (at === 'one' && spec?.endsWith('!') && n.args.length > 0) return true
        return n.args.some((a, i) => {
          const m = modeAt(spec, i)
          return go(a, m === 'w' ? 'whole' : m === 'p' ? at : 'one')
        })
      }
      case 'un':
      case 'pct':
        return go(n.a, 'one')
      case 'bin':
        return go(n.a, 'one') || go(n.b, 'one')
      case 'paren':
        return go(n.a, at)
      case 'let':
        return n.binds.some((b) => go(b.value, 'whole')) || (n.body !== undefined && go(n.body, at))
      default:
        return false
    }
  }
  return go(ast, 'one')
}

/** Reports the formulas of the old kind on a sheet that Excel 365 computes with implicit intersection. */
export function checkIntersections(sheet: string, old: Map<string, Ast>, names: NameInput[], issues: Issues): void {
  if (!old.size) return
  const areas = areaNames(names)
  for (const cell of sortAddresses(old.keys()))
    if (intersects(old.get(cell)!, areas)) issues.add('implicitIntersection', 'a formula of the old kind: Excel 365 computes it with implicit intersection (@) where the engine computes a dynamic array', { sheet, cell })
}
