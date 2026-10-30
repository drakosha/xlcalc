/** A group of `xlcalc/functions` as the registry takes it (`registry.ts`): the contract's fields and its definitions. */
import type { FnDef } from '../args.js'
import type { FunctionGroup, FunctionGroupName } from '../contract.js'

export interface LoadedGroup extends FunctionGroup {
  readonly defs: Readonly<Record<string, FnDef>>
}

/** A group of the definitions of `parts`. */
export function group(name: FunctionGroupName, ...parts: Readonly<Record<string, FnDef>>[]): FunctionGroup {
  const defs: Record<string, FnDef> = Object.assign({}, ...parts)
  const g: LoadedGroup = { kind: 'xlcalc-functions', name, functions: Object.keys(defs).sort(), defs }
  return g
}
