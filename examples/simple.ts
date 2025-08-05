import { WorkBook, WorkSheet } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件是一个最简单的示例，生成 Excel 和导出 CSV 一样简单
 * 请打开 xlsx/simple.xlsx 查看最终效果
 */

// 演示数据，一份学生成绩单
const data = [
  ['序号', '姓名', '成绩'],
  [1, '张三', 100],
  [2, '李四', 90],
  [3, '王五', 80],
];

// 创建工作表，名称为 simple，数据为 data，从 A1 单元格开始展示，不添加任何样式和其他选项
const sheet = new WorkSheet('simple', [{ origin: 'A1', data }]);

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/simple.xlsx', buffer);
