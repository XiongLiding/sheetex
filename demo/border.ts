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
  [1, 'none', { value: '无', style: 'borderNone' }],
  [2, 'hair', { value: '', style: 'borderHair' }],
  [3, 'dotted', { value: '', style: 'borderDotted' }],
  [4, 'dashDotDot', { value: '', style: 'borderDashDotDot' }],
  [5, 'dashDot', { value: '', style: 'borderDashDot' }],
  [6, 'dashed', { value: '', style: 'borderDashed' }],
  [7, 'thin', { value: '', style: 'borderThin' }],
];

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

const style: Styles = {
  default: {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    }
  },
  header: {
    font: {
      b: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    }
  },
  borderNone: {
    border: {
      bottom: {
        style: 'none', color: '00000000',
      }
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    }
  },
  borderDashDot: {
    border: {
      bottom: {
        style: 'dashDot', color: '00000000',
      }
    },
  },
  borderDashDotDot: {
    border: {
      bottom: {
        style: 'dashDotDot', color: '00000000',
      }
    },
  },
  borderDashed: {
    border: {
      bottom: {
        style: 'dashed', color: '00000000',
      }
    },
  },
  borderDotted: {
    border: {
      bottom: {
        style: 'dotted', color: '00000000',
      }
    },
  },
  borderDouble: {
    border: {
      bottom: {
        style: 'double', color: '00000000',
      }
    },
  },
  borderHair: {
    border: {
      bottom: {
        style: 'hair', color: '00000000',
      }
    },
  },
  borderMedium: {
    border: {
      bottom: {
        style: 'medium', color: '00000000',
      }
    },
  },
  borderMediumDashDot: {
    border: {
      bottom: {
        style: 'mediumDashDot', color: '00000000',
      }
    },
  },
  borderMediumDashDotDot: {
    border: {
      bottom: {
        style: 'mediumDashDotDot', color: '00000000',
      }
    },
  },
  borderMediumDashed: {
    border: {
      bottom: {
        style: 'mediumDashed', color: '00000000',
      }
    },
  },
  borderSlantDashDot: {
    border: {
      bottom: {
        style: 'slantDashDot', color: '00000000',
      }
    },
  },
  borderThick: {
    border: {
      bottom: {
        style: 'thick', color: '00000000',
      }
    },
  },
  borderThin: {
    border: {
      bottom: {
        style: 'thin', color: '00000000',
      }
    },
  },
};

const sheet = new WorkSheet(
  'border',
  [{
    origin: 'A1',
    data: data1,
  },
  {
    origin: 'E1',
    data: data2,
  }],
  style,
  {
    colWidths: [{ min: 1, max: 8, size: [10, 20, 20, 10] }]
  }
);

const workbook = new WorkBook([sheet]);
const buffer = await workbook.getZipBuffer();
writeFileSync('./xlsx/border.xlsx', buffer);
