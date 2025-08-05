import { WorkSheet, WorkBook, type Styles, type Alignment } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件演示字体相关设置
 * 请打开 xlsx/font.xlsx 查看最终效果
 */

// 演示数据
const data = [
  [
    { value: '序号', style: 'header' },
    { value: 'font.* ', style: 'header' },
    { value: '效果', style: 'header' },
  ],
  [1, 'sz=24', { value: '演示文本', style: 'fontSz' }],
  [2, 'name=微软雅黑', { value: '演示文本', style: 'fontName' }],
  [3, 'color=FF0000', { value: '演示文本', style: 'fontColor' }],
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

  // 字号（字体大小）
  fontSz: {
    font: {
      sz: 24,
    },
  },

  // 字体名称
  fontName: {
    font: {
      name: '微软雅黑',
    },
  },

  // 颜色
  fontColor: {
    font: {
      color: 'FF0000',
    },
  },

  // 粗体
  fontB: {
    font: {
      b: true,
    },
  },

  // 斜体
  fontI: {
    font: {
      i: true,
    },
  },

  // 以下是 4 种下划线样式
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
  fontUSingleAccounting: {
    font: {
      u: 'singleAccounting',
    },
  },
  fontUDoubleAccounting: {
    font: {
      u: 'doubleAccounting',
    },
  },

  // 删除线
  fontStrike: {
    font: {
      strike: true,
    },
  },

  // 以下是上标与下标
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

  // 字体样式的综合运用，不同的属性可以任意组合
  fontMixed: {
    font: {
      sz: 24,
      name: '微软雅黑',
      color: 'FF0000',
      b: true,
      i: true,
      u: 'single',
      strike: true,
      vertAlign: 'superscript',
    },
  },
};

// 定义工作表，名称为 font，里面包含 1 个数据块，套用 styles 中定义的样式，并设置列宽
const sheet = new WorkSheet('font', [{ origin: 'A1', data }], styles, {
  colWidths: [
    { min: 1, size: [10, 20, 20] }, // 未定义 max，会根据 size 的长度自动生成
  ],
});

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/font.xlsx', buffer);
