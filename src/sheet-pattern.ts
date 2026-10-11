/** A sheet name a reference holds without quotes (the tokenizer and `quoteSheetIfNeeded` read the same): a letter or `_`, then letters, digits, `_` and `.`. */
export const PLAIN_SHEET = '[\\p{L}_][\\p{L}\\p{N}_.]*'
