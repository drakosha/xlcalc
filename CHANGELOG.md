# Изменения

## 0.1.0 (2026-10-05)

Первая версия: движок формул вынесен из the app (`packages/formula`, пакет `@app/formula`, коммит the app `7bae0fb`) вместе с историей каталога (`git subtree split`). Код движка не менялся, значения и хэши те же, что в the app: синтетика (сид 1, 3 400 строк) `9d7414efbda34a63…`, версия конфигурации `app-formula/1;compiled-js;formulajs@4.6.1;canon15`.

Что в библиотеке с переезда:
- контракт: `compute`, `validateFormula`, `validateName`, `openSession` (`setCells`, `setNames`, `setClock`, `explain`, `result`, `close`), коды проверки, ошибки ячеек с причиной, волатильные ячейки, хэш значений и входа, внешние функции через `externals`;
- белый список: SUM, SUMIF, SUMIFS, COUNTIFS, IF, IFERROR, XLOOKUP, ROUND, EOMONTH, DATE, TEXT, TRUE, FALSE, TODAY, NOW, RAND, RANDBETWEEN; из formulajs 4.6.1 только ROUND, EOMONTH, DATE, SUM;
- тесты: семантика против документации Excel и оракула LibreOffice, детерминизм в трёх часовых зонах, сессия (200 правок против полного расчёта), бюджеты на синтетике, независимость исходников.

Новое при выносе:
- сборка `tsc` в `dist` с `.d.ts`, подпути `exports`: `.`, `./contract` (только контракт), `./synth` (генератор синтетики, переехал из `test/fixtures` в `src/synth.ts`), `./internal` (внутренности для харнесса замера в the app, не контракт);
- относительные импорты с расширением `.js` (NodeNext), без этого собранный `dist` не грузится в node;
- `npm run test:dist`: весь набор тестов против собранного `dist`, включая дочерние процессы; дымовой тест `test/dist.test.ts` импортирует подпути по имени пакета из чистого node;
- тест пакета проверяет, что единственная зависимость formulajs и HyperFormula нет ни в одной зависимости и в lock.
