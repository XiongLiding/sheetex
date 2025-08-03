import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Styles } from '../src/render/renderStyles.ts';

const data1 = [
  [
    { value: '序号', style: 'header' },
    { value: '水平对齐', style: 'header' },
    { value: '效果展示', style: 'header' },
    { value: '支持缩进', style: 'header' },
  ],
  [1, 'left', { value: '演示文本', style: 'alignHorizontalLeft' }, '是'],
  [2, 'center', { value: '演示文本', style: 'alignHorizontalCenter' }, '否'],
  [3, 'right', { value: '演示文本', style: 'alignHorizontalRight' }, '是'],
  [4, 'fill', { value: '演示文本', style: 'alignHorizontalFill' }, '否'],
  [5, 'justify', { value: '演示文本', style: 'alignHorizontalJustify' }, '否'],
  [6, 'centerContinuous', { value: '演示文本', style: 'alignHorizontalCenterContinuous' }, '否'],
  [7, 'distributed', { value: '演示文本', style: 'alignHorizontalDistributed' }, '是'],
];

const data2 = [
  [
    { value: '序号', style: 'header' },
    { value: '垂直对齐', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, 'top', { value: '演示文本', style: 'alignVerticalTop' }],
  [2, 'center', { value: '演示文本', style: 'alignVerticalCenter' }],
  [3, 'bottom', { value: '演示文本', style: 'alignVerticalBottom' }],
  [4, 'justify', { value: '演\n示\n文\n本', style: 'alignVerticalJustify' }],
  [5, 'distributed', { value: '演\n示\n文\n本', style: 'alignVerticalDistributed' }],
];

const data3 = [
  [
    { value: '序号', style: 'header' },
    { value: '旋转角度', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, 0, { value: '演示文本', style: 'alignRotate0' }],
  [2, 45, { value: '演示文本', style: 'alignRotate45' }],
  [3, 90, { value: '演示文本', style: 'alignRotate90' }],
  [4, 135, { value: '演示文本', style: 'alignRotate135' }],
  [5, 180, { value: '演示文本', style: 'alignRotate180' }],
];

const data4 = [
  [
    { value: '序号', style: 'header' },
    { value: '特殊效果', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, '自动换行', { value: '当演示文本的长度超过列宽', style: 'alignWrapText' }],
  [2, '缩小字体填充', { value: '当演示文本的长度超过列宽', style: 'alignShrinkToFit' }],
];

const style: Styles = {
  default: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    }
  },
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
  },
  alignHorizontalLeft: {
    alignment: {
      horizontal: 'left',
      vertical: 'center',
      indent: 1,
    },
  },
  alignHorizontalCenter: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
  },
  alignHorizontalRight: {
    alignment: {
      horizontal: 'right',
      vertical: 'center',
      indent: 1,
    },
  },
  alignHorizontalFill: {
    alignment: {
      horizontal: 'fill',
      vertical: 'center',
    },
  },
  alignHorizontalJustify: {
    alignment: {
      horizontal: 'justify',
      vertical: 'center',
    },
  },
  alignHorizontalCenterContinuous: {
    alignment: {
      horizontal: 'centerContinuous',
      vertical: 'center',
    },
  },
  alignHorizontalDistributed: {
    alignment: {
      horizontal: 'distributed',
      vertical: 'center',
      indent: 1,
    },
  },
  alignVerticalTop: {
    alignment: {
      horizontal: 'center',
      vertical: 'top',
    },
  },
  alignVerticalCenter: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
  },
  alignVerticalBottom: {
    alignment: {
      horizontal: 'center',
      vertical: 'bottom',
    },
  },
  alignVerticalJustify: {
    alignment: {
      horizontal: 'center',
      vertical: 'justify',
    },
  },
  alignVerticalDistributed: {
    alignment: {
      horizontal: 'center',
      vertical: 'distributed',
    },
  },
  alignRotate0: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 0,
    },
  },
  alignRotate45: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 45,
    },
  },
  alignRotate90: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 90,
    },
  },
  alignRotate135: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 135,
    },
  },
  alignRotate180: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 180,
    },
  },
  alignWrapText: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
  },
  alignShrinkToFit: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      shrinkToFit: true,
    },
  },
};

const sheet = new WorkSheet(
  'alignment',
  [
    { origin: 'A1', data: data1 },
    { origin: 'A10', data: data2 },
    { origin: 'F10', data: data3 },
    { origin: 'F1', data: data4 },
  ],
  style,
  {
    colWidths: [{ min: 1, size: [10, 20, 20, 10]}, { min: 6, size: [10, 20, 20] }],
    rowHeights: [{ min: 10, max: 20, size: 80 }],
  },
);

const workbook = new WorkBook([sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/alignment.xlsx', buffer);
