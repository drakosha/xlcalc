import type { FunctionCategory } from './contract.js'

/** The categories of functions in the order Excel's Insert Function window lists them (their members are `function-info.ts`'s). */
export const FUNCTION_CATEGORIES: readonly FunctionCategory[] = ['financial', 'date', 'math', 'statistical', 'lookup', 'database', 'text', 'logical', 'information', 'engineering', 'compatibility']
