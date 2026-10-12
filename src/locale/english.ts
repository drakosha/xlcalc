/** A message in English, as en-US writes it: what the parser and the checks say outside a locale (`validateFormula`, input errors). */
import { EN_MESSAGES } from '../locales/en-US.js'
import type { EngineKey, MessageParams } from './messages.js'

export const english = <K extends EngineKey>(m: { key: K } & MessageParams[K]): string => EN_MESSAGES[m.key](m)
