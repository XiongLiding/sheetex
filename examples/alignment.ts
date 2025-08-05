import { WorkSheet, WorkBook, type Styles, type Alignment } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件演示水平对齐、垂直对齐、旋转角度、自动换行、缩小字体填充功能
 * 请打开 xlsx/alignment.xlsx 查看最终效果
 */

// 水平对齐演示数据
const dataVertical = [
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

// 垂直对齐演示数据
const dataHorizontal = [
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

// 文本旋转角度演示数据
const dataRotate = [
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

// 自动换行和缩小字体填充演示数据
const dataMisc = [
  [
    { value: '序号', style: 'header' },
    { value: '特殊效果', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, '自动换行', { value: '当演示文本的长度超过列宽', style: 'alignWrapText' }],
  [2, '缩小字体填充', { value: '当演示文本的长度超过列宽', style: 'alignShrinkToFit' }],
];

// 定义居中对齐，方便复用
const alignment: Alignment = {
  horizontal: 'center',
  vertical: 'center',
};

// 样式定义，为了清晰展示各种效果，本示例代码特意避免逻辑运算，实际工作中可以通过循环、函数生成、对象合并等方式提高样式管理效率
const styles: Styles = {
  default: {
    alignment,
  },
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    alignment,
  },

  // 以下是水平对齐样式
  alignHorizontalLeft: {
    alignment: {
      horizontal: 'left', // 水平居左
      vertical: 'center',
      indent: 1, // 缩进1格，只有居左、居右和两端对齐支持缩进
    },
  },
  alignHorizontalCenter: {
    alignment: {
      horizontal: 'center', // 水平居中
      vertical: 'center',
    },
  },
  alignHorizontalRight: {
    alignment: {
      horizontal: 'right', // 水平居右
      vertical: 'center',
      indent: 1, // 缩进1格，只有居左、居右和两端对齐支持缩进
    },
  },
  alignHorizontalFill: {
    alignment: {
      horizontal: 'fill', // 填充，重复文字填充单元格剩余空间
      vertical: 'center',
    },
  },
  alignHorizontalJustify: {
    alignment: {
      horizontal: 'justify', // 两端对齐
      vertical: 'center',
    },
  },
  alignHorizontalCenterContinuous: {
    alignment: {
      horizontal: 'centerContinuous', // 跨列居中
      vertical: 'center',
    },
  },
  alignHorizontalDistributed: {
    alignment: {
      horizontal: 'distributed', // 分散对齐
      vertical: 'center',
      indent: 1, // 缩进1格，只有居左、居右和分散对齐支持缩进
    },
  },

  // 以下是垂直对齐样式
  alignVerticalTop: {
    alignment: {
      horizontal: 'center',
      vertical: 'top', // 垂直居上
    },
  },
  alignVerticalCenter: {
    alignment: {
      horizontal: 'center',
      vertical: 'center', // 垂直居中
    },
  },
  alignVerticalBottom: {
    alignment: {
      horizontal: 'center',
      vertical: 'bottom', // 垂直居下
    },
  },
  alignVerticalJustify: {
    alignment: {
      horizontal: 'center',
      vertical: 'justify', // 两端对齐
    },
  },
  alignVerticalDistributed: {
    alignment: {
      horizontal: 'center',
      vertical: 'distributed', // 分散对齐
    },
  },

  // 以下是旋转角度样式，请注意旋转并不是连续的，0-90，90-180，180-255对应三套不同的规则，并且临界值在不同软件中可能有不同的行为
  alignRotate0: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 0, // 水平
    },
  },
  alignRotate45: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 45, // 文字逆时针旋转 45 度
    },
  },
  alignRotate90: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 90, // 文字逆时针旋转 90 度
    },
  },
  alignRotate135: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 135, // 文字顺时针旋转 45 度
    },
  },
  alignRotate180: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      textRotation: 180, // 文字顺时针旋转 90度
    },
  },

  // 以下是自动换行和缩小字体填充样式，定义文字超过当前单元格大小时的行为
  alignWrapText: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true, // 自动换行
    },
  },
  alignShrinkToFit: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      shrinkToFit: true, // 缩小字体填充
    },
  },
};

// 将 4 个数据块放到对应的位置
const dataBlocks = [
  { origin: 'A1', data: dataVertical },
  { origin: 'A10', data: dataHorizontal },
  { origin: 'F10', data: dataRotate },
  { origin: 'F1', data: dataMisc },
];

// 定义工作表，名称为 alignment 里面包含 4 个数据块，套用 styles 中定义的样式，并设置合并单元格、列宽和行高
const sheet = new WorkSheet('alignment', dataBlocks, styles, {
  colWidths: [
    { min: 1, size: [10, 20, 20, 10] },
    { min: 6, size: [10, 20, 20] },
  ],
  rowHeights: [{ min: 10, max: 20, size: 80 }], // 将10-20行高度设为80以便展示出垂直对齐和旋转角度的效果
});

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/alignment.xlsx', buffer);
