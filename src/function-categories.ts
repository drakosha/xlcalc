/** The categories of functions in the order Excel's Insert Function window lists them, apart from their members (`function-info.ts`), for a grid's window without them. */
import type { FunctionCategory } from './contract.js'

export const FUNCTION_CATEGORIES: readonly FunctionCategory[] = ['financial', 'date', 'math', 'statistical', 'lookup', 'database', 'text', 'logical', 'information', 'engineering', 'compatibility']
