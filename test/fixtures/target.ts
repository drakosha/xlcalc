/** What child processes of the tests load: the sources (default) or the built `dist` (`FORMULA_TARGET=dist`). */
export const TARGET = process.env.FORMULA_TARGET === 'dist' ? 'dist' : 'src'

/** URL of a library module in the build under test: `libUrl('index.js')`. */
export const libUrl = (file: string): string => new URL(`../../${TARGET}/${file}`, import.meta.url).href
