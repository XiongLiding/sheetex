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
  [4, 'justify', { value: '演示文本', style: 'alignVerticalJustify' }],
  [5, 'distributed', { value: '演示文本', style: 'alignVerticalDistributed' }],
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
  [1, '自动换行', { value: '超过列宽的演示文本', style: 'alignWrapText' }],
  [2, '缩小字体填充', { value: '超过列宽演示文本', style: 'alignShrinkToFit' }],
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
    },
  },
  alignHorizontalLeft: {
    alignment: {
      horizontal: 'left',
      indent: 1,
    },
  },
  alignHorizontalCenter: {
    alignment: {
      horizontal: 'center',
    },
  },
  alignHorizontalRight: {
    alignment: {
      horizontal: 'right',
      indent: 1,
    },
  },
  alignHorizontalFill: {
    alignment: {
      horizontal: 'fill',
    },
  },
  alignHorizontalJustify: {
    alignment: {
      horizontal: 'justify',
    },
  },
  alignHorizontalCenterContinuous: {
    alignment: {
      horizontal: 'centerContinuous',
    },
  },
  alignHorizontalDistributed: {
    alignment: {
      horizontal: 'distributed',
      indent: 1,
    },
  },
  alignVerticalTop: {
    alignment: {
      vertical: 'top',
    },
  },
  alignVerticalCenter: {
    alignment: {
      vertical: 'center',
    },
  },
  alignVerticalBottom: {
    alignment: {
      vertical: 'bottom',
    },
  },
  alignVerticalJustify: {
    alignment: {
      vertical: 'justify',
    },
  },
  alignVerticalDistributed: {
    alignment: {
      vertical: 'distributed',
    },
  },
  alignRotate0: {
    alignment: {
      textRotation: 0,
    },
  },
  alignRotate45: {
    alignment: {
      textRotation: 45,
    },
  },
  alignRotate90: {
    alignment: {
      textRotation: 90,
    },
  },
  alignRotate135: {
    alignment: {
      textRotation: 135,
    },
  },
  alignRotate180: {
    alignment: {
      textRotation: 180,
    },
  },
  alignRotate225: {
    alignment: {
      textRotation: 225,
    },
  },
  alignRotate270: {
    alignment: {
      textRotation: 270,
    },
  },
  alignRotate315: {
    alignment: {
      textRotation: 315,
    },
  },
  alignWrapText: {
    alignment: {
      wrapText: true,
    },
  },
  alignShrinkToFit: {
    alignment: {
      shrinkToFit: true,
    }
  }
};

const sheet = new WorkSheet(
  'alignment1',
  [
    { origin: 'A1', data: data1, },
    { origin: 'A10', data: data2, },
    { origin: 'F10', data: data3 },
    { origin: 'F1', data: data4 },
  ],
  style,
);

const workbook = new WorkBook('alignment', [sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/alignment.xlsx', buffer);
