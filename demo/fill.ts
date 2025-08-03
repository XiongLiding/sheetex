import WorkSheet from '../src/process/WorkSheet.ts';
import WorkBook from '../src/process/WorkBook.ts';
import { writeFileSync } from 'node:fs';
import { type Styles } from '../src/render/renderStyles.ts';

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

const style: Styles = {
  header: {
    font: {
      b: true,
    },
    fill: {
      patternType: 'none',
    }
  },
  fillSolid: {
    fill: {
      patternType: 'solid',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkHorizontal: {
    fill: {
      patternType: 'darkHorizontal',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightHorizontal: {
    fill: {
      patternType: 'lightHorizontal',
      fgColor: 'FF000000',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkGray: {
    fill: {
      patternType: 'darkGray',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkVertical: {
    fill: {
      patternType: 'darkVertical',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightVertical: {
    fill: {
      patternType: 'lightVertical',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillMediumGray: {
    fill: {
      patternType: 'mediumGray',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkDown: {
    fill: {
      patternType: 'darkDown',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightDown: {
    fill: {
      patternType: 'lightDown',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightGray: {
    fill: {
      patternType: 'lightGray',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkUp: {
    fill: {
      patternType: 'darkUp',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightUp: {
    fill: {
      patternType: 'lightUp',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillGray125: {
    fill: {
      patternType: 'gray125',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkGrid: {
    fill: {
      patternType: 'darkGrid',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightGrid: {
    fill: {
      patternType: 'lightGrid',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillGray0625: {
    fill: {
      patternType: 'gray0625',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillDarkTrellis: {
    fill: {
      patternType: 'darkTrellis',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    },
  },
  fillLightTrellis: {
    fill: {
      patternType: 'lightTrellis',
      fgColor: 'FFCCCCCC',
      bgColor: 'FFFF0000',
    }
  }
};

const sheet = new WorkSheet(
  'fill1',
  [
    {
      origin: 'A1',
      data: data1,
    },
    {
      origin: 'D1',
      data: data2,
    },
    {
      origin: 'G1',
      data: data3,
    },
    {
      origin: 'J1',
      data: data4,
    },
    {
      origin: 'M1',
      data: data5,
    },
    {
      origin: 'P1',
      data: data6,
    }
  ],
  style,
);

const workbook = new WorkBook('fill', [sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/fill.xlsx', buffer);
