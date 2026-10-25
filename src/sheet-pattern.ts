/** A sheet name a reference holds without quotes (the tokenizer and `quoteSheetIfNeeded` read the same): a letter or `_`, then letters, digits, `_` and `.`. */
export const PLAIN_SHEET = '[\\p{L}_][\\p{L}\\p{N}_.]*'

/** The same pattern as the whole name (`quoteSheetIfNeeded`, `ADDRESS`, the name of a link). */
export const PLAIN_NAME = new RegExp(`^${PLAIN_SHEET}$`, 'u')

/** A sheet name in quotes, each apostrophe inside doubled. */
export const quoteSheet = (s: string) => `'${s.replaceAll("'", "''")}'`
