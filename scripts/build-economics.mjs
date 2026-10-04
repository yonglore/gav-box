import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const out = fileURLToPath(new URL('../outputs/gav-box-2026-10-04/', import.meta.url));
const work = fileURLToPath(new URL('../.work/', import.meta.url));
await fs.mkdir(out, {recursive:true});
await fs.mkdir(work, {recursive:true});
const wb = Workbook.create();
const s = wb.worksheets.add('Экономика');
s.showGridLines = false;
s.tabColor = '#192B4B';
s.getRange('A1:E57').format.font.name = 'Arial';
s.getRange('A1:E57').format.font.size = 11;
s.getRange('A1:E57').format.rowHeight = 23;
s.getRange('A1:E57').format.verticalAlignment = 'center';
s.getRange('A1:A57').format.columnWidth = 53;
s.getRange('B1:B57').format.columnWidth = 18;
s.getRange('C1:C57').format.columnWidth = 3;
s.getRange('D1:D57').format.columnWidth = 76;
s.getRange('E1:E57').format.columnWidth = 3;
const num = '#,##0.00;(#,##0.00);"—"';
s.getRange('B5:B54').setNumberFormat(num);
const value = (row,label,v,note='') => {
  s.getRange(`A${row}`).values=[[label]];
  if(v!==null) s.getRange(`B${row}`).values=[[v]];
  if(note) s.getRange(`D${row}`).values=[[note]];
};
const formula = (row,label,f,note='') => {
  value(row,label,null,note); s.getRange(`B${row}`).formulas=[[f]];
};
const section = (row,label) => {
  value(row,label,null);
  s.getRange(`A${row}:B${row}`).format.fill='#192B4B';
  s.getRange(`A${row}:B${row}`).format.font.color='#FFFFFF';
  s.getRange(`A${row}:B${row}`).format.font.bold=true;
};
value(2,'Гав Бокс: экономика одного заказа',null);
s.getRange('A2').format.font.size=16;
s.getRange('A2').format.font.bold=true;
s.getRange('A2:D2').format.borders={bottom:{style:'thin',color:'#192B4B'}};
value(3,'Москва · расчётные допущения · 04.10.2026',null);
s.getRange('A3').format.font.italic=true;
formula(5,'Доход после привлечения, ₽','=B42','До постоянных расходов; это не чистая прибыль.');
formula(6,'Доля от выручки','=B43');
formula(7,'Заказов/месяц для покрытия всех постоянных','=B45','Все заказы новые; включает заданную оплату основателям.');
formula(8,'Резерв денег на пилот до выручки, ₽','=B54','Консервативно: все расходы оплачены до первых поступлений.');
formula(9,'Результат пилота после подготовки, ₽','=B53','До ежемесячных постоянных расходов; партия продаётся целиком.');
s.getRange('A5:B9').format.fill='#EDF1F7';
s.getRange('A5:B9').format.font.bold=true;
s.getRange('B6').setNumberFormat('0.0%');
s.getRange('B7').setNumberFormat('#,##0');
value(11,'Меняйте синие ячейки с жёлтой заливкой.',null,'Все входные числа — допущения, не тарифы и не полученные прайсы.');
value(12,'Единицы: рубли, если не указано иначе.',null,'Модель без НДС. При возникновении НДС требуется перерасчёт.');
section(14,'Допущения на заказ и пилот');
const inputs=[
[15,'Цена набора',2490,'Проверить оплатой реального заказа.'],
[16,'Доставка, оплаченная покупателем',299,'Поступление продавцу; входит в базу резерва налога и эквайринга.'],
[17,'Закупка содержимого',950,'Целевой бюджет. Заменить котировкой поставщика.'],
[18,'Коробка и вкладыш',150,'Допущение на единицу.'],
[19,'Входящая доставка на единицу',50,'Распределить стоимость поставки по партии.'],
[20,'Труд сборки и обработки заказа',100,'Включён даже при самостоятельной сборке.'],
[21,'Доставка покупателю',350,'Не тариф службы. Требуется маршрут, вес и размер.'],
[22,'Эквайринг, доля всей выручки',0.025,'Заменить условиями договора.'],
[23,'Резерв налога, доля всей выручки',0.06,'Гипотеза УСН «доходы»; не индивидуальный налоговый расчёт.'],
[24,'Резерв потерь, доля цены набора',0.04,'Чистые ожидаемые расходы на претензии, а не частота возвратов.'],
[25,'CAC нового покупателя',500,'Реклама, подаренные образцы и относимые расходы привлечения.'],
[26,'Постоянные расходы в месяц',15000,'Бухгалтерия, касса, хранение, взносы и прочее: оценка.'],
[27,'Оплата основателям в месяц',80000,'Дополнительно к сборке. Не дублировать труд, включённый в CAC.'],
[28,'Количество новых заказов в пилоте',20,'Не факт спроса; весь объём предполагается проданным.'],
[29,'Разовые расходы подготовки',12000,'Оценка. Проверить реальную стоимость оформления запуска.'],
[30,'Прототипы и образцы',4000,'Отдельно от продаваемой партии.'],
[31,'Дополнительный денежный резерв',3000,'Не расход в расчёте результата до фактического использования.']
];
for(const row of inputs) value(...row);
s.getRange('B15:B31').format.font.color='#0000FF';
s.getRange('B15:B31').format.fill='#FFF2CC';
s.getRange('B22:B24').setNumberFormat('0.0%');
s.getRange('B28').setNumberFormat('#,##0');
s.getRange('B15:B31').dataValidation={rule:{type:'decimal',operator:'greaterThanOrEqual',formula1:0}};
s.getRange('B22:B24').dataValidation={rule:{type:'decimal',operator:'between',formula1:0,formula2:1}};
s.getRange('B28').dataValidation={rule:{type:'whole',operator:'greaterThanOrEqual',formula1:1}};
value(32,'Справка по налоговой гипотезе',null,'https://www.nalog.gov.ru/rn77/taxation/taxes/usn/');
section(34,'Расчёт на один новый заказ');
formula(35,'Выручка с доставкой','=IF(COUNT(B15:B16)=2,SUM(B15:B16),"n.a.")');
formula(36,'Содержимое, упаковка, приёмка и сборка','=SUM(B17:B20)');
formula(37,'Эквайринг','=B35*B22');
formula(38,'Налоговый резерв','=B35*B23');
formula(39,'Резерв потерь','=B15*B24');
formula(40,'Все переменные расходы до привлечения','=IF(COUNT(B15:B24)=10,SUM(B36:B39)+B21,"n.a.")');
formula(41,'Маржинальный доход до привлечения','=IF(ISNUMBER(B40),B35-B40,"n.a.")');
formula(42,'Маржинальный доход после привлечения','=IF(AND(ISNUMBER(B41),ISNUMBER(B25)),B41-B25,"n.a.")');
formula(43,'Доля дохода после привлечения','=IF(AND(ISNUMBER(B42),ISNUMBER(B35)),IF(B35=0,"n.a.",B42/B35),"n.a.")');
formula(44,'Заказов для покрытия расходов без основателей','=IF(AND(ISNUMBER(B42),ISNUMBER(B26)),IF(B42<=0,"n.a.",ROUNDUP(B26/B42,0)),"n.a.")');
formula(45,'Заказов для покрытия расходов с основателями','=IF(AND(ISNUMBER(B42),COUNT(B26:B27)=2),IF(B42<=0,"n.a.",ROUNDUP(SUM(B26:B27)/B42,0)),"n.a.")');
s.getRange('B43').setNumberFormat('0.0%');
s.getRange('B44:B45').setNumberFormat('#,##0');
section(48,'Пилот: все заказы новые');
formula(49,'Выручка пилота','=IF(AND(ISNUMBER(B35),ISNUMBER(B28)),B35*B28,"n.a.")');
formula(50,'Переменные расходы пилота с привлечением','=IF(AND(ISNUMBER(B40),ISNUMBER(B25),ISNUMBER(B28)),(B40+B25)*B28,"n.a.")');
formula(51,'Доход до расходов подготовки','=IF(COUNT(B49:B50)=2,B49-B50,"n.a.")');
formula(52,'Подготовка и прототипы','=IF(COUNT(B29:B30)=2,SUM(B29:B30),"n.a.")');
formula(53,'Результат после подготовки','=IF(COUNT(B51:B52)=2,B51-B52,"n.a.")','Ежемесячные расходы строк 26–27 не вычтены.');
formula(54,'Денежный резерв до получения выручки','=IF(COUNT(B50,B52,B31)=3,SUM(B50,B52,B31),"n.a.")','Предоплата может снизить потребность, но создаёт обязательства.');
value(56,'n.a. — результата нет: проверьте входные данные.',null,'Нулевая или отрицательная маржа не даёт точки безубыточности.');
for(const row of [41,42,45,51,53,54]) {
  s.getRange(`A${row}:B${row}`).format.font.bold=true;
  s.getRange(`A${row}:B${row}`).format.borders={top:{style:'thin',color:'#A8B2C2'}};
}
for(const r of ['B5','B9','B41:B42','B51','B53']) s.getRange(r).conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{fill:'#FCE4D6',font:{color:'#9C0006',bold:true}}});
s.getRange('D3:D54').format.font.color='#56657A';
s.getRange('D3:D54').format.wrapText=false;
s.freezePanes.freezeRows(14);

// Independent numerical check, then perturb a material driver and restore.
wb.recalculate();
const base=Number(s.getRange('B42').values[0][0]);
if(Math.abs(base-352.335)>1e-6) throw new Error(`Base CM2 mismatch: ${base}`);
s.getRange('B25').values=[[1000]];
wb.recalculate();
if(Math.abs(Number(s.getRange('B5').values[0][0])-(-147.665))>1e-6) throw new Error('CAC dependency failed');
if(s.getRange('B7').values[0][0]!=='n.a.') throw new Error('Negative-margin break-even guard failed');
s.getRange('B25').values=[[500]];
s.getRange('B16').values=[[0]];
wb.recalculate();
if(Math.abs(Number(s.getRange('B5').values[0][0])-78.75)>1e-6) throw new Error('Shipping dependency failed');
s.getRange('B16').values=[[299]];
s.getRange('B17').clear({applyTo:'contents'});
wb.recalculate();
if(s.getRange('B5').values[0][0]!=='n.a.' || s.getRange('B8').values[0][0]!=='n.a.') throw new Error('Missing input guard failed');
s.getRange('B17').values=[[950]];
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'Экономика!A5:B9',include:'values,formulas',tableMaxRows:5,tableMaxCols:2})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},summary:'Formula errors'})).ndjson);
for(const [name,range] of [['top','A1:D32'],['bottom','A34:D56']]) {
 const img=await wb.render({sheetName:s.name,range,scale:1.5,format:'png'});
 await fs.writeFile(`${work}/economics-${name}.png`,new Uint8Array(await img.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(wb)).save(`${out}/gav-box-economics.xlsx`);
await fs.writeFile(`${work}/economics-checked.json`,JSON.stringify({cm2:base,pilot:s.getRange('B49:B54').values},null,2));
console.log('Exported workbook; base, CAC, shipping and negative-margin checks passed.');
