/**
 * Held-out display cases of the Hijri months and weekdays under `B2` (Excel en-US, `display calendar-hijri-held`), each answer predicted and
 * committed before Excel showed the case, after the set `calendar-names-held`, where Excel for Mac wrote them in Arabic in
 * both locales: predicted by the Unicode CLDR Arabic names of the tabular Hijri calendar with a no-break space between
 * words (as Excel wrote جمادى الأولى), `ddd` as CLDR's short weekday, `mmm` as `mmmm`, `mmmmm` the name's first letter;
 * and the days around the end of 1445, a leap year of 355 days, by the tabular calendar.
 */
import type { DisplayCase } from './cases.js'

export function buildDisplayCases(): DisplayCase[] {
  return [
    {"id":"hijri#1","cell":{"value":45480.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"محرم"}},
    {"id":"hijri#2","cell":{"value":45480.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الأحد"}},
    {"id":"hijri#3","cell":{"value":45510.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"صفر"}},
    {"id":"hijri#4","cell":{"value":45510.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الثلاثاء"}},
    {"id":"hijri#5","cell":{"value":45539.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"ربيع الأول"}},
    {"id":"hijri#6","cell":{"value":45539.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الأربعاء"}},
    {"id":"hijri#7","cell":{"value":45569.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"ربيع الآخر"}},
    {"id":"hijri#8","cell":{"value":45569.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الجمعة"}},
    {"id":"hijri#9","cell":{"value":45598.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"جمادى الأولى"}},
    {"id":"hijri#10","cell":{"value":45598.5,"format":"B2dddd"},"width":40,"predicted":{"text":"السبت"}},
    {"id":"hijri#11","cell":{"value":45628.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"جمادى الآخرة"}},
    {"id":"hijri#12","cell":{"value":45628.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الاثنين"}},
    {"id":"hijri#13","cell":{"value":45657.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"رجب"}},
    {"id":"hijri#14","cell":{"value":45657.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الثلاثاء"}},
    {"id":"hijri#15","cell":{"value":45687.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"شعبان"}},
    {"id":"hijri#16","cell":{"value":45687.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الخميس"}},
    {"id":"hijri#17","cell":{"value":45716.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"رمضان"}},
    {"id":"hijri#18","cell":{"value":45716.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الجمعة"}},
    {"id":"hijri#19","cell":{"value":45746.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"شوال"}},
    {"id":"hijri#20","cell":{"value":45746.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الأحد"}},
    {"id":"hijri#21","cell":{"value":45775.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"ذو القعدة"}},
    {"id":"hijri#22","cell":{"value":45775.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الاثنين"}},
    {"id":"hijri#23","cell":{"value":45805.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"ذو الحجة"}},
    {"id":"hijri#24","cell":{"value":45805.5,"format":"B2dddd"},"width":40,"predicted":{"text":"الأربعاء"}},
    {"id":"hijri#25","cell":{"value":45480.5,"format":"B2ddd"},"width":40,"predicted":{"text":"الأحد"}},
    {"id":"hijri#26","cell":{"value":45510.5,"format":"B2ddd"},"width":40,"predicted":{"text":"الثلاثاء"}},
    {"id":"hijri#27","cell":{"value":45539.5,"format":"B2ddd"},"width":40,"predicted":{"text":"الأربعاء"}},
    {"id":"hijri#28","cell":{"value":45569.5,"format":"B2ddd"},"width":40,"predicted":{"text":"الجمعة"}},
    {"id":"hijri#29","cell":{"value":45598.5,"format":"B2ddd"},"width":40,"predicted":{"text":"السبت"}},
    {"id":"hijri#30","cell":{"value":45628.5,"format":"B2ddd"},"width":40,"predicted":{"text":"الاثنين"}},
    {"id":"hijri#31","cell":{"value":45657.5,"format":"B2ddd"},"width":40,"predicted":{"text":"الثلاثاء"}},
    {"id":"hijri#32","cell":{"value":45480.5,"format":"B2mmm"},"width":40,"predicted":{"text":"محرم"}},
    {"id":"hijri#33","cell":{"value":45569.5,"format":"B2mmm"},"width":40,"predicted":{"text":"ربيع الآخر"}},
    {"id":"hijri#34","cell":{"value":45805.5,"format":"B2mmm"},"width":40,"predicted":{"text":"ذو الحجة"}},
    {"id":"hijri#35","cell":{"value":45480.5,"format":"B2mmmmm"},"width":40,"predicted":{"text":"م"}},
    {"id":"hijri#36","cell":{"value":45569.5,"format":"B2mmmmm"},"width":40,"predicted":{"text":"ر"}},
    {"id":"hijri#37","cell":{"value":45805.5,"format":"B2mmmmm"},"width":40,"predicted":{"text":"ذ"}},
    {"id":"hijri#38","cell":{"value":45478.5,"format":"B2dd/mm/yyyy"},"width":40,"predicted":{"text":"29/12/1445"}},
    {"id":"hijri#39","cell":{"value":45479.5,"format":"B2dd/mm/yyyy"},"width":40,"predicted":{"text":"30/12/1445"}},
    {"id":"hijri#40","cell":{"value":45480.5,"format":"B2dd/mm/yyyy"},"width":40,"predicted":{"text":"01/01/1446"}},
  ]
}
