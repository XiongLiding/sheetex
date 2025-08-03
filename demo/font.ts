import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Styles } from '../src/render/renderStyles.ts';

const data = [
  [
    { value: '序号', style: 'header' },
    { value: 'font.* ', style: 'header' },
    { value: '效果', style: 'header' },
  ],
  [1, 'sz=24', { value: '演示文本', style: 'fontSz' }],
  [2, 'name=微软雅黑', { value: '演示文本', style: 'fontName' }],
  [3, 'color=FFFF0000', { value: '演示文本', style: 'fontColor' }],
  [4, 'b=true', { value: '演示文本', style: 'fontB' }],
  [5, 'i=true', { value: '演示文本', style: 'fontI' }],
  [6, 'u=single', { value: '演示文本', style: 'fontUSingle' }],
  [7, 'u=double', { value: '演示文本', style: 'fontUDouble' }],
  [8, 'u=singleAccounting', { value: '演示文本', style: 'fontUSingleAccounting' }],
  [9, 'u=doubleAccounting', { value: '演示文本', style: 'fontUDoubleAccounting' }],
  [10, 'strike=true', { value: '演示文本', style: 'fontStrike' }],
  [11, 'vertAlign=superscript', { value: '演示文本', style: 'fontVertAlignSuperscript' }],
  [12, 'vertAlign=subscript', { value: '演示文本', style: 'fontVertAlignSubscript' }],
  [13, '综合展示', { value: '演示文本', style: 'fontMixed' }],
];

const style: Styles = {
  default: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
  },
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    }
  },
  fontSz: {
    font: {
      sz: 24,
    },
  },
  fontName: {
    font: {
      name: '微软雅黑',
    },
  },
  fontColor: {
    font: {
      color: 'FFFF0000',
    },
  },
  fontB: {
    font: {
      b: true,
    },
  },
  fontI: {
    font: {
      i: true,
    },
  },
  fontUSingle: {
    font: {
      u: 'single',
    },
  },
  fontUDouble: {
    font: {
      u: 'double',
    },
  },
  fontUAccountingSingle: {
    font: {
      u: 'singleAccounting',
    },
  },
  fontUAccountingDouble: {
    font: {
      u: 'doubleAccounting',
    },
  },
  fontStrike: {
    font: {
      strike: true,
    },
  },
  fontVertAlignSuperscript: {
    font: {
      vertAlign: 'superscript',
    },
  },
  fontVertAlignSubscript: {
    font: {
      vertAlign: 'subscript',
    },
  },
  fontMixed: {
    font: {
      sz: 24,
      name: '微软雅黑',
      color: 'FFFF0000',
      b: true,
      i: true,
      u: 'single',
      strike: true,
      vertAlign: 'superscript',
    },
  },
};

const sheet = new WorkSheet(
  'font',
  {
    origin: 'A1',
    data,
  },
  style,
  {
    colWidths: [
      {
        min: 1,
        size: [10, 20, 20],
      },
    ],
  },
);

const workbook = new WorkBook([sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/font.xlsx', buffer);
