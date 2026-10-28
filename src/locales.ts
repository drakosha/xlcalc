/**
 * The `./locales` subpath: `loadLocale(tag)`, one of the package's locales loaded by its tag, for a host that knows a
 * locale by its tag alone (a file's, a user's choice). Its module (`xlcalc/locales/ru-RU`) comes as a chunk of its own
 * where the bundler splits, and only into the bundles that import this subpath; the subpath takes none of the engine.
 */
import { strict } from './strict-args.js'
import * as packaged from './locale/package.js'

export const loadLocale = /* @__PURE__ */ strict(packaged.loadLocale)
