import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Styles } from '../src/render/renderStyles.ts';

const data = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, '', { value: 12345.6789 }],
  [2, '0', { value: 12345.6789, style: 'fmt0' }],
  [3, '0.00', { value: 12345.6789, style: 'fmt2' }],
  [4, '#,##0', { value: 12345.6789, style: 'fmtComma' }],
  [5, '#,##0.00', { value: 12345.6789, style: 'fmtComma2' }],
  [6, '0%', { value: 12345.6789, style: 'fmtPercent' }],
  [7, '0.00%', { value: 12345.6789, style: 'fmtPercent2' }],
  [8, '0.00E+00', { value: 12345.6789, style: 'fmtScientific' }],
  [9, '$#,##0.00', { value: 12345.6789, style: 'fmtFinance' }],
];

const style: Styles = {
  default: {
    alignment: {
      horizontal: 'right',
    },
  },
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    alignment: {
      horizontal: 'right',
    },
  },
  fmt0: {
    formatCode: '0',
  },
  fmt2: {
    formatCode: '0.00',
  },
  fmtComma: {
    formatCode: '#,##0',
  },
  fmtComma2: {
    formatCode: '#,##0.00',
  },
  fmtPercent: {
    formatCode: '0%',
  },
  fmtPercent2: {
    formatCode: '0.00%',
  },
  fmtScientific: {
    formatCode: '0.00E+00',
  },
  fmtFinance: {
    formatCode: '$#,##0.00',
  },
};

const sheet = new WorkSheet(
  'format',
  {
    origin: 'A1',
    data,
  },
  style,
  {
    colWidths: [{ min: 1, size: [10, 20, 20] }],
  },
);

const workbook = new WorkBook([sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/format.xlsx', buffer);
