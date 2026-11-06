/** Criteria and sorting shared by worksheet and table writers; stored hidden rows are left as supplied. */
import type { ColumnFilter, ConditionalStyle, FilterColumn, FilterCriteria, FilterSort, SortCondition } from '../../contract.js'
import { escapeAttr } from '../xml/text.js'
import type { CellFormats } from './styles.js'

type Styles = Pick<CellFormats, 'dxf'>
const attrs = (a: Record<string, string | number | boolean | undefined>) => Object.entries(a).filter(([,v])=>v!==undefined).map(([k,v])=>` ${k}="${escapeAttr(typeof v==='boolean'?(v?'1':'0'):String(v))}"`).join('')
const tag = (name:string,a:Record<string,string|number|boolean|undefined>,body?:string) => `<${name}${attrs(a)}${body?`>${body}</${name}>`:'/>'}`
const x14 = (body:string,fallback='') => `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><mc:Choice Requires="x14">${body}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

function color(style:ConditionalStyle,font:boolean,styles:Styles):number {
  const c=font?style.font?.color??style.fill?.color:style.fill?.color??style.font?.color
  return styles.dxf(c?{fill:{color:c}}:{},true)
}

function criterion(f:ColumnFilter,styles:Styles):string {
  switch(f.type){
    case 'values':{
      const values=f.values??[]
      const list=values.some(v=>v.length>255)?x14(values.map(val=>tag('x14:filter',{val})).join(''),values.filter(v=>v.length<=255).map(val=>tag('filter',{val})).join('')):values.map(val=>tag('filter',{val})).join('')
      return tag('filters',{blank:f.blank,calendarType:f.calendar},list+(f.dates??[]).map(d=>tag('dateGroupItem',{year:d.year,month:d.month,day:d.day,hour:d.hour,minute:d.minute,second:d.second,dateTimeGrouping:d.level})).join(''))
    }
    case 'custom':{
      const long=f.conditions.some(c=>c.value.length>255);const prefix=long?'x14:':''
      const text=tag(prefix+'customFilters',{and:f.and},f.conditions.map(c=>tag(prefix+'customFilter',{operator:c.operator,val:c.value})).join(''))
      return long?x14(text):text
    }
    case 'top':return tag('top10',{top:f.bottom===undefined?undefined:!f.bottom,percent:f.percent,val:f.value,filterVal:f.filterValue})
    case 'dynamic':return tag('dynamicFilter',{type:f.kind,val:f.value,maxVal:f.maxValue,valIso:f.valueIso,maxValIso:f.maxValueIso})
    case 'color':return tag('colorFilter',{dxfId:color(f.style,!!f.font,styles),cellColor:f.font===undefined?undefined:!f.font})
    case 'icon':{
      const extended=['3Stars','3Triangles','5Boxes'].includes(f.set);const prefix=extended?'x14:':''
      const text=tag(prefix+'iconFilter',{iconSet:f.set,iconId:f.icon})
      return extended?x14(text):text
    }
  }
}

function condition(c:SortCondition,styles:Styles):string {
  const extended=c.iconSet!==undefined&&['3Stars','3Triangles','5Boxes'].includes(c.iconSet)
  const text=tag((extended?'x14:':'')+'sortCondition',{ref:c.range,descending:c.descending,sortBy:c.by,customList:c.customList,dxfId:c.style===undefined?undefined:color(c.style,c.by==='fontColor',styles),iconSet:c.iconSet,iconId:c.icon})
  return extended?x14(text):text
}

/** The last sort, without changing the order of any cells. */
export const sortXml = (s:FilterSort|undefined,styles:Styles):string => s?tag('sortState',{ref:s.range,caseSensitive:s.caseSensitive,columnSort:s.byColumns,sortMethod:s.method},(s.conditions??[]).map(c=>condition(c,styles)).join('')):''

/** A filter and its real criteria, or the legacy filter-mode marker when no criteria were supplied. */
export function filterXml(range:string,criteria:FilterCriteria|undefined,styles:Styles,fallback=''):string {
  const columns=(criteria?.columns??[]).map((c:FilterColumn)=>tag('filterColumn',{colId:c.column,hiddenButton:c.hiddenButton,showButton:c.showButton},c.filter?criterion(c.filter,styles):'')).join('')
  return tag('autoFilter',{ref:range},(criteria===undefined?fallback:columns)+sortXml(criteria?.sort,styles))
}
