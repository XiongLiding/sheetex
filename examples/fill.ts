import { WorkSheet, WorkBook, type Styles, type Alignment } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件演示填充样式（单元格背景色）
 * 请打开 xlsx/fill.xlsx 查看最终效果
 * 图案样式排列顺序与 Microsoft Excel 填充设置界面中一致，分成 6 列，每列 3 种样式
 * 填充由两种颜色组成，可以理解成先在下层涂满 bgColor，然后在上层用 fgColor 画出 patternType 对应纹理，上下叠加就是最终填充效果
 * 注意当 bgColor 或 fgColor 未指定时，默认值为 auto，就本库而言表示透明
 *
 * 由于即时通讯软件中的预览功能支持有限，在这些软件中看不到纯色以外的填充效果属于正常现象
 */

// 第 1 列图案的演示数据
const data1 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, 'solid', { value: '', style: 'fillSolid' }],
  [2, 'darkHorizontal', { value: '', style: 'fillDarkHorizontal' }],
  [3, 'lightHorizontal', { value: '', style: 'fillLightHorizontal' }],
];

// 第 2 列图案的演示数据
const data2 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [4, 'darkGray', { value: '', style: 'fillDarkGray' }],
  [5, 'darkVertical', { value: '', style: 'fillDarkVertical' }],
  [6, 'lightVertical', { value: '', style: 'fillLightVertical' }],
];

// 第 3 列图案的演示数据
const data3 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [7, 'mediumGray', { value: '', style: 'fillMediumGray' }],
  [8, 'darkDown', { value: '', style: 'fillDarkDown' }],
  [9, 'lightDown', { value: '', style: 'fillLightDown' }],
];

// 第 4 列图案的演示数据
const data4 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [10, 'lightGray', { value: '', style: 'fillLightGray' }],
  [11, 'darkUp', { value: '', style: 'fillDarkUp' }],
  [12, 'lightUp', { value: '', style: 'fillLightUp' }],
];

// 第 5 列图案的演示数据
const data5 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [13, 'gray125', { value: '', style: 'fillGray125' }],
  [14, 'darkGrid', { value: '', style: 'fillDarkGrid' }],
  [15, 'lightGrid', { value: '', style: 'fillLightGrid' }],
];

// 第 6 列图案的演示数据
const data6 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [16, 'gray0625', { value: '', style: 'fillGray0625' }],
  [17, 'darkTrellis', { value: '', style: 'fillDarkTrellis' }],
  [18, 'lightTrellis', { value: '', style: 'fillLightTrellis' }],
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
      b: true,
    },
    fill: {
      patternType: 'none',
    },
    alignment,
  },

  // 纯色填充，最常用的填充形式，Microsoft Excel 顶部快捷工具中的油漆桶对应的就是纯色填充
  fillSolid: {
    fill: {
      patternType: 'solid', // 指定图案样式为纯色背景，纯色背景比较特殊，上层把颜色涂满了
      fgColor: 'CCCCCC', // 前景色，就是最终看到的颜色
      bgColor: 'FF0000', // 背景色，在纯色填充时可以忽略
    },
  },

  // 剩余 17 种图案样式，很少用，而且其他软件对这些图案的支持都比较差
  fillDarkHorizontal: {
    fill: {
      patternType: 'darkHorizontal',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightHorizontal: {
    fill: {
      patternType: 'lightHorizontal',
      fgColor: '000000',
      bgColor: 'FF0000',
    },
  },
  fillDarkGray: {
    fill: {
      patternType: 'darkGray',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillDarkVertical: {
    fill: {
      patternType: 'darkVertical',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightVertical: {
    fill: {
      patternType: 'lightVertical',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillMediumGray: {
    fill: {
      patternType: 'mediumGray',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillDarkDown: {
    fill: {
      patternType: 'darkDown',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightDown: {
    fill: {
      patternType: 'lightDown',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightGray: {
    fill: {
      patternType: 'lightGray',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillDarkUp: {
    fill: {
      patternType: 'darkUp',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightUp: {
    fill: {
      patternType: 'lightUp',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillGray125: {
    fill: {
      patternType: 'gray125',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillDarkGrid: {
    fill: {
      patternType: 'darkGrid',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightGrid: {
    fill: {
      patternType: 'lightGrid',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillGray0625: {
    fill: {
      patternType: 'gray0625',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillDarkTrellis: {
    fill: {
      patternType: 'darkTrellis',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
  fillLightTrellis: {
    fill: {
      patternType: 'lightTrellis',
      fgColor: 'CCCCCC',
      bgColor: 'FF0000',
    },
  },
};

// 将 6 个数据块放到对应的地址
const dataBlocks = [
  { origin: 'A1', data: data1 },
  { origin: 'D1', data: data2 },
  { origin: 'G1', data: data3 },
  { origin: 'J1', data: data4 },
  { origin: 'M1', data: data5 },
  { origin: 'P1', data: data6 },
];

// 定义工作表，名称为 fill，里面包含 6 个数据块，套用 styles 中定义的样式，并设置列宽
const sheet = new WorkSheet('fill', dataBlocks, styles, {
  colWidths: [{ min: 1, max: 18, size: [10, 20, 20] }], // 列宽定义，size 不足，自动循环补足 18 列
});

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/fill.xlsx', buffer);
