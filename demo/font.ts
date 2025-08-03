import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Styles } from '../src/render/renderStyles.ts';

const data = [
  [
    { value: '序号', style: 'header' },
    { value: '姓名', style: 'header' },
    { value: '成绩', style: 'header' },
  ],
  [1, '张三', 100],
  [2, '李四', 90],
  [3, '王五', { value: 59, style: 'fail' }],
  ['', { value: '平均分', style: 'summary' }, { value: 83, style: 'summary' }],
];

const style: Styles = {
  header: {
    font: {
      sz: 16,
      name: '微软雅黑',
      b: true,
    },
  },
  fail: {
    font: {
      sz: 18,
      i: true,
      strike: true,
      color: 'FFFF0000',
    },
  },
  summary: {
    font: {
      sz: 16,
      b: true,
      u: 'doubleAccounting',
      vertAlign: 'subscript',
    },
  }
};

const sheet = new WorkSheet(
  'sheet1',
  {
    origin: 'A1',
    data,
  },
  style,
);

const workbook = new WorkBook('header', [sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/font.xlsx', buffer);
