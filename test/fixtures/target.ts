/** Что проверяют дочерние процессы тестов: исходники (по умолчанию) или собранный `dist` (`FORMULA_TARGET=dist`). */
export const TARGET = process.env.FORMULA_TARGET === 'dist' ? 'dist' : 'src'

/** URL модуля библиотеки в проверяемой сборке: `libUrl('index.js')`. */
export const libUrl = (file: string): string => new URL(`../../${TARGET}/${file}`, import.meta.url).href
