import WorkSheet, { type DataRow } from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Alignment, type Border, type Styles } from '../src/render/renderStyles.ts';

const sign = {
  value: '√',
  style: 'sign',
};

const notSign = {
  value: '×',
  style: 'notSign',
};

const data: DataRow[] = [
  [{ value: '2025年度第30周考勤汇总', style: 'caption' }],
  [
    { value: '序号', style: 'header' },
    { value: '工号', style: 'header' },
    { value: '打卡情况', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '出勤天数', style: 'header' },
  ],
  [
    { value: '', style: 'header' },
    { value: '', style: 'header' },
    { value: '一', style: 'header' },
    { value: '二', style: 'header' },
    { value: '三', style: 'header' },
    { value: '四', style: 'header' },
    { value: '五', style: 'header' },
    { value: '六', style: 'header' },
    { value: '日', style: 'header' },
    { value: '', style: 'header' },
  ],
  [1, 'A01001', sign, sign, sign, sign, sign, notSign, notSign, 5],
  [2, 'A01002', sign, sign, sign, sign, sign, notSign, notSign, 5],
  [
    { value: 3, style: 'fail' },
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

const border: Border = {
  style: 'thin',
  color: 'FF000000',
};

const alignment: Alignment = {
  horizontal: 'center',
  vertical: 'center',
};

const style: Styles = {
  default: {
    border,
    alignment,
  },
  caption: {
    font: {
      sz: 18,
      name: '微软雅黑',
      b: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
    border,
  },
  header: {
    font: {
      name: '微软雅黑',
      b: true,
    },
    border,
    alignment,
  },
  sign: {
    font: {
      name: '宋体',
      color: 'FF296011',
    },
    border,
    alignment,
    fill: {
      patternType: 'solid',
      fgColor: 'FFCFEED0',
    },
  },
  notSign: {
    font: {
      name: '宋体',
      color: 'FF8E1813',
    },
    border,
    alignment,
    fill: {
      patternType: 'solid',
      fgColor: 'FFF6C9CE',
    },
  },
  fail: {
    alignment,
    border,
    font: {
      color: 'FFFF0000',
      b: true,
    },
  },
};

const sheet = new WorkSheet('options', { origin: 'A1', data }, style, {
  mergeCells: ['A1:J1', 'A2:A3', 'B2:B3', 'C2:I2', 'J2:J3'],
  colWidths: [{ min: 1, size: [8, 10, 8, 8, 8, 8, 8, 8, 8, 10] }],
  rowHeights: [
    { min: 1, size: 40 },
    { min: 2, max: 13, size: 24 },
  ],
});

const workbook = new WorkBook([sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/showtime.xlsx', buffer);
