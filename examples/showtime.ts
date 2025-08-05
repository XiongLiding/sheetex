import { WorkBook, WorkSheet, type DataRow, type Alignment, type Border, type Styles } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件将以贴近现实需求的“周考勤汇总表”来演示 sheetex 的使用方法
 * 请打开 xlsx/showtime.xlsx 查看最终效果
 */

// 定义签到格，重复出现的单元格可以通过预先定义，来简化数据表达
const sign = {
  value: '√',
  style: 'sign',
};

// 定义未签到格
const notSign = {
  value: '×',
  style: 'notSign',
};

// 定义表格数据
const data: DataRow[] = [
  // 标题，如果单元格合并后无需边框，可以忽略其他被合并的格子
  [{ value: '2025年度第30周考勤汇总', style: 'caption' }],

  // 表头
  [
    { value: '序号', style: 'header' },
    { value: '工号', style: 'header' },
    { value: '打卡情况', style: 'header' }, // 打卡情况横跨7列，后面被合并的6列用空字符串填充，并使用相同的样式，确保合并后边框正常展示
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '出勤天数', style: 'header' },
  ],
  [
    { value: '', style: 'header' }, // 被上方的序号合并，用空字符串填充
    { value: '', style: 'header' }, // 被上方的工号合并，用空字符串填充
    { value: '一', style: 'header' },
    { value: '二', style: 'header' },
    { value: '三', style: 'header' },
    { value: '四', style: 'header' },
    { value: '五', style: 'header' },
    { value: '六', style: 'header' },
    { value: '日', style: 'header' },
    { value: '', style: 'header' }, // 被上方的出勤天数合并，用空字符串填充
  ],

  // 数据
  [1, 'A01001', sign, sign, sign, sign, sign, notSign, notSign, 5],
  [2, 'A01002', sign, sign, sign, sign, sign, notSign, notSign, 5],
  [
    { value: 3, style: 'fail' }, // 考勤不合格，使用专门的样式突出显示
    { value: 'A01003', style: 'fail' },
    sign,
    sign,
    sign,
    sign,
    notSign,
    notSign,
    notSign,
    { value: 4, style: 'fail' },
  ],
  [4, 'A01004', sign, sign, sign, sign, notSign, notSign, sign, 5],
  [5, 'A01005', sign, sign, sign, sign, sign, notSign, notSign, 5],
  [6, 'B02001', sign, sign, notSign, sign, sign, sign, notSign, 5],
  [7, 'B02002', sign, sign, sign, sign, sign, notSign, notSign, 5],
  [8, 'B02003', sign, sign, sign, sign, sign, sign, notSign, 6],
  [9, 'B02004', notSign, sign, sign, sign, sign, sign, notSign, 5],
  [10, 'B02005', sign, notSign, sign, sign, sign, sign, notSign, 5],
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
const styles: Styles = {
  // default 为系统保留样式名，所有未明确指定 style 的格子，都会套用此样式
  default: {
    alignment,
    border,
  },

  // 标题样式
  caption: {
    font: {
      sz: 18, // 需要比其他文字更大
      name: '微软雅黑',
      b: true,
    },
    alignment,
  },

  // 表头样式
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    alignment,
    border,
  },

  // 考勤情况样式
  sign: {
    font: {
      name: '宋体',
      color: '296011',
    },
    alignment,
    border,
    fill: {
      patternType: 'solid',
      fgColor: 'CFEED0',
    },
  },
  notSign: {
    font: {
      name: '宋体',
      color: '8E1813',
    },
    alignment,
    border,
    fill: {
      patternType: 'solid',
      fgColor: 'F6C9CE',
    },
  },

  // 考勤不合格样式
  fail: {
    font: {
      color: 'FF0000',
      b: true,
    },
    alignment,
    border,
  },
};

// 定义工作表，名称为 showtime，里面包含一个位于 A1 的数据块，套用 styles 中定义的样式，并设置合并单元格、列宽和行高
const sheet = new WorkSheet('showtime', [{ origin: 'A1', data }], styles, {
  mergeCells: ['A1:J1', 'A2:A3', 'B2:B3', 'C2:I2', 'J2:J3'], // 合并单元格，依次是标题、序号、工号、打卡情况、出勤天数
  colWidths: [{ min: 1, size: [8, 10, 8, 8, 8, 8, 8, 8, 8, 10] }], // 定义 A-J 列的列宽
  rowHeights: [
    { min: 1, size: 40 }, // 标题行高度
    { min: 2, max: 13, size: 24 }, // 表头与数据行高度
  ],
});

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/showtime.xlsx', buffer);
