import { WorkSheet, WorkBook, type Styles } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件演示数字格式化
 * 请打开 xlsx/format.xlsx 查看最终效果
 *
 * 这里只演示了最常用的数字格式化代码，实际能实现的效果远超此处所展示
 * 关于格式代码的完整知识，请参考：
 * https://support.microsoft.com/zh-cn/office/-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5
 */

// 演示数据，这里将演示 12345.6789 这个数字在各种格式化下的效果
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

// 样式定义，为了清晰展示各种效果，本示例代码特意避免逻辑运算，实际工作中可以通过循环、函数生成、对象合并等方式提高样式管理效率
const styles: Styles = {
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
    formatCode: '0', // 取整
  },
  fmt2: {
    formatCode: '0.00', // 2 位小数
  },
  fmtComma: {
    formatCode: '#,##0', // 带千分位（每 3 位一个逗号）
  },
  fmtComma2: {
    formatCode: '#,##0.00', // 带千分位（每 3 位一个逗号），同时保留2位小数
  },
  fmtPercent: {
    formatCode: '0%', // 百分比
  },
  fmtPercent2: {
    formatCode: '0.00%', // 百分比，同时保留2位小数
  },
  fmtScientific: {
    formatCode: '0.00E+00', // 科学计数法
  },
  fmtFinance: {
    formatCode: '$#,##0.00', // 金额
  },
};

// 定义工作表，名称为 format，里面包含 1 个数据块，套用 styles 中定义的样式，并设置列宽
const sheet = new WorkSheet('format', [{ origin: 'A1', data }], styles, {
  colWidths: [{ min: 1, size: [10, 20, 20] }],
});

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/format.xlsx', buffer);
