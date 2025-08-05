import { WorkSheet, WorkBook, type Styles, type Border, type Alignment } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件演示行高、列宽、合并单元格功能的使用
 * 请打开 xlsx/option.xlsx 查看最终效果
 */

// 演示数据，一份学生成绩单
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

// 定义细线边框，方便复用
const border: Border = {
  style: 'thin',
  color: '000000',
};

// 定义居中对齐，方便复用
const alignment: Alignment = {
  horizontal: 'center',
  vertical: 'center',
};

// 样式定义，为了清晰展示各种效果，本示例代码特意避免逻辑运算，实际工作中可以通过循环、函数生成、对象合并等方式提高样式管理效率
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

// 定义工作表，名称为 option，里面包含一个位于 B2 的数据块，套用 styles 中定义的样式，并设置合并单元格、列宽和行高
const sheet = new WorkSheet('option', [{ origin: 'B2', data }], style, {
  mergeCells: ['B2:D2'], // 合并标题“成绩单”
  colWidths: [{ min: 2, size: [10, 20, 20] }], // 设置列宽，让序号比另两列窄
  rowHeights: [
    { min: 2, size: 40 }, // 设置标题行高度
    { min: 3, max: 6, size: 30 }, // 设置表头和数据行高度
  ],
});

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/option.xlsx', buffer);
