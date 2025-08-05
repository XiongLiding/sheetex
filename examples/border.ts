import { WorkSheet, WorkBook, type Styles, type Alignment } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 本文件演示边框样式
 * 请打开 xlsx/border.xlsx 查看最终效果
 * 线型排列顺序与 Microsoft Excel 边框设置界面中一致，分成左右两列，每列 7 种样式
 * 在超分辨率屏幕上，部分线型看起来会与设置界面不同，属于正常现象，实际工作中尽量避免使用这些线型
 */

// 左边 7 种线型样式演示数据
const data1 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [1, 'none', { value: '无', style: 'borderNone' }],
  [2, 'hair', { value: '', style: 'borderHair' }],
  [3, 'dotted', { value: '', style: 'borderDotted' }],
  [4, 'dashDotDot', { value: '', style: 'borderDashDotDot' }],
  [5, 'dashDot', { value: '', style: 'borderDashDot' }],
  [6, 'dashed', { value: '', style: 'borderDashed' }],
  [7, 'thin', { value: '', style: 'borderThin' }],
];

// 右边 7 种线型样式演示数据
const data2 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [8, 'mediumDashDotDot', { value: '', style: 'borderMediumDashDotDot' }],
  [9, 'slantDashDot', { value: '', style: 'borderSlantDashDot' }],
  [10, 'mediumDashDot', { value: '', style: 'borderMediumDashDot' }],
  [11, 'mediumDashed', { value: '', style: 'borderMediumDashed' }],
  [12, 'medium', { value: '', style: 'borderMedium' }],
  [13, 'thick', { value: '', style: 'borderThick' }],
  [14, 'double', { value: '', style: 'borderDouble' }],
];

// 对角线以及边框综合应用的演示数据
const data3 = [
  [
    { value: '序号', style: 'header' },
    { value: '格式代码', style: 'header' },
    { value: '效果展示', style: 'header' },
  ],
  [''],
  [1, 'diagonalUp', { value: '', style: 'borderDiagonalUp' }],
  [''],
  [2, 'diagonalDown', { value: '', style: 'borderDiagonalDown' }],
  [''],
  [3, '显示双对角线', { value: '', style: 'borderDiagonalBoth' }],
  [''],
  [4, '混合相同样式', { value: '', style: 'borderSame' }],
  [''],
  [5, '混合不同样式', { value: '', style: 'borderDifferent' }],
  [''],
  [6, '定义所有边框', { value: '', style: 'borderAll' }],
  [''],
  [7, '覆盖特定位置', { value: '', style: 'borderOverride' }],
];

// 定义居中对齐，方便复用
const alignment: Alignment = {
  horizontal: 'center',
  vertical: 'center',
};

// 样式定义，为了清晰展示各种效果，本示例代码特意避免逻辑运算，实际工作中可以通过循环、函数生成、对象合并等方式提高样式管理效率
const style: Styles = {
  default: {
    alignment,
  },
  header: {
    font: {
      b: true,
    },
    alignment,
  },

  // 以下是 14 种边框样式的定义，所有边框都被定义为下边框，以防上下单元格之间重叠，颜色为黑色（RGB格式）
  borderNone: {
    border: {
      bottom: {
        style: 'none',
        color: '000000',
      },
    },
    alignment,
  },
  borderHair: {
    border: {
      bottom: {
        style: 'hair',
        color: '000000',
      },
    },
  },
  borderDotted: {
    border: {
      bottom: {
        style: 'dotted',
        color: '000000',
      },
    },
  },
  borderDashDotDot: {
    border: {
      bottom: {
        style: 'dashDotDot',
        color: '000000',
      },
    },
  },
  borderDashDot: {
    border: {
      bottom: {
        style: 'dashDot',
        color: '000000',
      },
    },
  },
  borderDashed: {
    border: {
      bottom: {
        style: 'dashed',
        color: '000000',
      },
    },
  },
  borderThin: {
    border: {
      bottom: {
        style: 'thin',
        color: '000000',
      },
    },
  },

  borderMediumDashDotDot: {
    border: {
      bottom: {
        style: 'mediumDashDotDot',
        color: '000000',
      },
    },
  },
  borderSlantDashDot: {
    border: {
      bottom: {
        style: 'slantDashDot',
        color: '000000',
      },
    },
  },
  borderMediumDashDot: {
    border: {
      bottom: {
        style: 'mediumDashDot',
        color: '000000',
      },
    },
  },
  borderMediumDashed: {
    border: {
      bottom: {
        style: 'mediumDashed',
        color: '000000',
      },
    },
  },
  borderMedium: {
    border: {
      bottom: {
        style: 'medium',
        color: '000000',
      },
    },
  },
  borderThick: {
    border: {
      bottom: {
        style: 'thick',
        color: '000000',
      },
    },
  },
  borderDouble: {
    border: {
      bottom: {
        style: 'double',
        color: '000000',
      },
    },
  },

  // 以下是对角线的定义
  borderDiagonalUp: {
    border: {
      diagonalUp: true, // 需要同时定义对角线的方向和对角线的样式
      diagonal: {
        style: 'thin',
        color: '000000',
      },
    },
  },
  borderDiagonalDown: {
    border: {
      diagonalDown: true,
      diagonal: {
        style: 'thin',
        color: '000000',
      },
    },
  },
  borderDiagonalBoth: {
    border: {
      diagonalUp: true, // 两条对角线可以同时展示
      diagonalDown: true,
      diagonal: {
        style: 'thin',
        color: '000000',
      },
    },
  },

  // 多个边框可以任意组合
  borderSame: {
    border: {
      top: {
        style: 'thin',
        color: '000000',
      },
      right: {
        style: 'thin',
        color: '000000',
      },
    },
  },

  // 同一个单元格的边框可以定义成不同的线性和颜色
  borderDifferent: {
    border: {
      top: {
        style: 'thin',
        color: '000000',
      },
      right: {
        style: 'thick',
        color: 'FF0000',
      },
    },
  },

  // 如果同一个单元格上下左右边框样式全部相同，可以直接在 border 下定义统一的样式
  borderAll: {
    border: {
      style: 'thin',
      color: '000000',
    },
  },

  // 如果有需要可以先统一定义，然后覆盖掉特定位置
  borderOverride: {
    border: {
      style: 'thin',
      color: '000000',
      bottom: {
        // 覆盖掉默认的边框样式
        style: 'thick',
        color: 'FF0000',
      },
    },
  },
};

// 定义工作表，名称为 border，里面包含 2 个数据块，套用 styles 中定义的样式，并设置列宽
const sheet = new WorkSheet(
  'border',
  [
    { origin: 'A1', data: data1 }, // 左列从 A1 开始
    { origin: 'E1', data: data2 }, // 右列从 E1 开始
    { origin: 'A10', data: data3 }, // 在下方展示对角线和混合样式
  ],
  style,
  {
    colWidths: [{ min: 1, max: 8, size: [10, 20, 20, 10] }], // 列宽定义，size 不足，自动重复成 [10, 20, 20, 10, 10, 20, 20, 10]
    rowHeights: [{ min: 10, max: 24, size: [20, 10] }],
  },
);

// 定义工作簿
const workbook = new WorkBook([sheet]);

// 生成 Excel 文件的 buffer
const buffer = await workbook.getZipBuffer();

// 写入到指定位置
writeFileSync('./xlsx/border.xlsx', buffer);
