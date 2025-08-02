import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';

const data = [
  ['序号', '姓名', '成绩'],
  [1, '张三', 100],
  [2, '李四', 90],
  [3, '王五', 80],
];

const sheet = new WorkSheet(
  'sheet1',
  {
    origin: 'A1',
    data,
  },
  {},
);

const workbook = new WorkBook('simple', [sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('simple.xlsx', buffer);
