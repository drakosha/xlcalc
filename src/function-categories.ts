import type { FunctionCategory } from './contract.js'

/** The function categories in the order Excel's Insert Function dialog lists them (their members come from `function-info.ts`). */
export const FUNCTION_CATEGORIES: readonly FunctionCategory[] = ['financial', 'date', 'math', 'statistical', 'lookup', 'database', 'text', 'logical', 'information', 'engineering', 'compatibility']
