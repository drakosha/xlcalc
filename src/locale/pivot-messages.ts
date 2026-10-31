/**
 * en-US's text of an edit's messages about PivotTables, slicers and timelines (`PivotEditKey`): a module of its own, out
 * of the xlsx subpaths' bundles, which take an edit's other messages with the copy of shared formulas.
 */
import { sayEdit } from './edit-messages.js'
import { sayBy, type PivotMessages } from './messages.js'

export const PIVOT_MESSAGES: PivotMessages = {
  pivotUnknown: (p) => `there is no PivotTable ${p.pivot} on ${p.sheet}`,
  pivotSourceKept: (p) => `the PivotTable ${p.pivot} keeps the data it has: ${p.why === 'outside' ? 'its source is in another workbook' : p.why === 'kind' ? 'its source is a consolidation, an external connection, the data model or scenarios, which are not refreshed' : 'its cache does not allow refreshing'}`,
  pivotSourceGone: (p) => `the source of the PivotTable ${p.pivot} is not in the workbook (Reference isn't valid)`,
  pivotFieldName: (p) => `the PivotTable field name is not valid: the source of ${p.pivot} must be a list with a label at the top of every column`,
  pivotNotLoaded: () => "The PivotTable engine is not loaded: registerPivot(pivotEngine) loads it",
  pivotOverData: (p) => `there's already data in ${p.sheet}!${p.range}, where the PivotTable ${p.pivot} would be laid out`,
  pivotOverPivot: (p) => `the PivotTable ${p.pivot} would overlap the PivotTable ${p.other} (a PivotTable report cannot overlap another PivotTable report)`,
  pivotOverTable: (p) => `the PivotTable ${p.pivot} would overlap the table ${p.table}`,
  pivotPart: (p) => `${p.sheet}!${p.cell} is part of the PivotTable ${p.pivot}: we can't change this part of the PivotTable`,
  pivotAffect: (p) => `${p.sheet}!${p.cell} is part of the PivotTable ${p.pivot}: we can't make this change for the selected cells because it will affect a PivotTable`,
  pivotFormula: (p) => `${p.sheet}!${p.cell} is part of the PivotTable ${p.pivot}: cannot enter a formula for an item or field name in a PivotTable report`,
  pivotNameTaken: (p) => `PivotTable field name already exists: ${JSON.stringify(p.name)}`,
  pivotNotItem: (p) => `${JSON.stringify(p.item)} is not an item of this field`,
  pivotCantGroup: () => 'Cannot group that selection: a field of dates or numbers holds text, or a field of groups is grouped again',
  pivotNotGrouped: (p) => `The field ${JSON.stringify(p.field)} is not grouped`,
  pivotGroupedItems: () => 'Calculated items cannot be added to a grouped field, and a field with calculated items cannot be grouped',
  pivotItemsAverage: () => 'Averages, standard deviations, and variances are not supported when a PivotTable report has calculated items',
  pivotFormulaRefs: () => 'References, names, and arrays are not supported in PivotTable formulas',
  pivotFormulaBad: (p) => `The formula ${JSON.stringify(p.formula)} contains an error`,
  pivotItemTaken: (p) => `PivotTable item name already exists: ${JSON.stringify(p.item)}`,
  slicerField: (p) => `there is no field ${JSON.stringify(p.field)} in the PivotTables' source`,
  slicerNameTaken: (p) => `the name ${JSON.stringify(p.name)} is already used by a slicer or a timeline`,
  slicerNameEmpty: () => 'a slicer or a timeline needs a name',
  timelineField: (p) => `a timeline needs a field of dates, and ${JSON.stringify(p.field)} holds none`,
  pivotCut: (p) => `the operation would change part of the PivotTable ${p.pivot} on '${p.sheet}' (Excel moves or deletes a PivotTable whole)`,
  pivotTableNameTaken: (p) => `a PivotTable named ${JSON.stringify(p.name)} already exists on '${p.sheet}' (names are case-insensitive)`,
  pivotNameEmpty: () => 'a PivotTable needs a name',
  pivotSourceRows: (p) => `the source ${p.source} needs a header row and at least one row below it`,
}

/** A message of the engine, an edit's of PivotTables among them, in the language of a locale (`ownText`), else in English. */
export const sayPivot = /* @__PURE__ */ sayBy(PIVOT_MESSAGES, sayEdit)
