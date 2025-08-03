import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Alignment, type Border, type Styles } from '../src/render/renderStyles.ts';

const data = [
  [
    { value: '成绩单', style: 'caption' },
    { value: '', style: 'caption' },
    { value: '', style: 'caption' },
  ],
  [
    { value: '序号', style: 'header' },
    { value: '姓名', style: 'header' },
    { value: '成绩', style: 'header' },
  ],
  [1, '张三', 100],
  [2, '李四', 90],
  [3, '王五', 80],
];

const border: Border = {
  style: 'thin',
  color: 'FF000000',
};

const alignment: Alignment = {
  horizontal: 'center',
  vertical: 'center',
};

const style: Styles = {
  default: {
    border,
    alignment,
  },
  caption: {
    font: {
      sz: 24,
      name: '微软雅黑',
      b: true,
    },
    border,
    alignment,
  },
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    border,
    alignment,
  },
};

const sheet = new WorkSheet('options', { origin: 'B2', data }, style, {
  mergeCells: ['B2:D2'],
  colWidths: [{ min: 2, size: [10, 20, 20] }],
  rowHeights: [
    { min: 2, size: 40 },
    { min: 3, max: 6, size: 30 },
  ],
});

const workbook = new WorkBook([sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/option.xlsx', buffer);
