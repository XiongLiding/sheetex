import { type Style } from '../render/renderStyles.ts';
import renderSheet, { type RenderSize } from '../render/renderWorkSheet.ts';

export type DataCell =
  | number
  | string
  | {
      value: string | number;
      style?: string;
    };
export type DataRow = DataCell[];
export type DataBlock = {
  origin: string;
  data: DataRow[];
};

export type SheetCell = {
  value: number | string;
  style: number;
};

export function aton(column: string) {
  return column
    .split('')
    .reverse()
    .reduce((acc, char, i) => {
      return acc + (char.charCodeAt(0) - 'A'.charCodeAt(0) + (i >= 1 ? 1 : 0)) * 26 ** i;
    }, 1);
}

export function ntoa(column: number) {
  let result = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    column = Math.floor((column - 1) / 26);
  }
  return result;
}

export interface Size {
  min: number;
  max?: number;
  size: number | number[];
}

export interface SheetOptions {
  mergeCells?: string[];
  colWidths?: Size[];
  rowHeights?: Size[];
}

export default class WorkSheet {
  public readonly name: string;
  public readonly styles: Record<string, Style>;
  public styleIndex?: Record<string, number>;
  private readonly blocks: DataBlock | DataBlock[];
  private readonly cells: Record<number, Record<number, SheetCell>> = {};
  private readonly mergeCells: string[];
  private readonly colWidths: Size[];
  private readonly rowHeights: Size[];

  /**
   * @param name 工作表名称
   * @param blocks 数据块
   * @param styles 样式表
   * @param options 工作表选项
   */
  constructor(
    name: string,
    blocks: DataBlock | DataBlock[],
    styles: Record<string, Style> = {},
    options: SheetOptions = {},
  ) {
    if (!name) throw new Error('工作表名称不能为空');
    const invalidName = /[:\\\/?*\[\]]/.test(name);
    if (invalidName) throw new Error('工作表名称包含非法字符');
    this.name = name.substring(0, 31);
    this.blocks = blocks;
    this.styles = styles;
    this.mergeCells = options.mergeCells ?? [];
    this.colWidths = options.colWidths ?? [];
    this.rowHeights = options.rowHeights ?? [];
  }

  private processCell(cell: number | string | DataCell): SheetCell {
    if (typeof cell === 'number') {
      return {
        value: cell,
        style: this.styleIndex?.default ?? 0,
      };
    }
    if (typeof cell === 'string') {
      return {
        value: cell,
        style: this.styleIndex?.default ?? 0,
      };
    }
    if (typeof cell === 'object') {
      return {
        value: cell.value ?? '',
        style: cell.style ? (this.styleIndex[cell.style ?? 'default'] ?? 0) : 0,
      };
    }
  }

  private consumeDataBlock(block: DataBlock) {
    const origin = block.origin.toUpperCase();

    const regex = /^([A-Z]+)([1-9]\d*)$/;
    const match = origin.match(regex);
    if (!match) {
      throw new Error('单元格坐标格式错误');
    }

    const [, column, row] = match;
    const rowOffset = parseInt(row, 10) - 1;
    const columnOffset = aton(column) - 1;

    block.data.forEach((row, y) => {
      row.forEach((cell, x) => {
        if (!this.cells[y + rowOffset]) {
          this.cells[y + rowOffset] = {};
        }
        this.cells[y + rowOffset][x + columnOffset] = this.processCell(cell);
      });
    });
  }

  private processSize(colWidths: Size[]) {
    const renderSize: RenderSize[] = [];

    colWidths.forEach((col) => {
      // size 是单个数字，可忽略 max，此时 max = min
      if (typeof col.size === 'number') {
        renderSize.push({
          min: col.min,
          max: col.max ?? col.min,
          size: col.size,
        });
      }

      // size 是数组
      if (Array.isArray(col.size)) {
        if (!col.max || col.max < col.min) {
          // max 未定义或无效，根据数组长度生成一批列定义
          col.size.forEach((size, i) => {
            if (typeof size !== 'number' || size < 0) return;

            renderSize.push({
              min: col.min + i,
              max: col.min + i,
              size,
            });
          });
        } else {
          // max 有效，在范围内用数组循环填充
          for (let i = col.min; i <= col.max; i++) {
            const size = col.size[(i - col.min) % col.size.length];
            if (typeof size !== 'number' || size < 0) continue;

            renderSize.push({
              min: i,
              max: i,
              size,
            });
          }
        }
      }
    });

    return renderSize;
  }

  public render() {
    if (!this.styleIndex) {
      throw new Error('调用顺序错误：完成样式预处理才能进行渲染');
    }

    if (Array.isArray(this.blocks)) {
      this.blocks.forEach((block) => {
        this.consumeDataBlock(block);
      });
    } else {
      this.consumeDataBlock(this.blocks);
    }

    const colWidths = this.processSize(this.colWidths);
    const rowHeights = this.processSize(this.rowHeights);

    const rowHeightIndex: Record<number, number> = {};
    rowHeights.forEach((v) => {
      for (let i = v.min; i <= v.max; i++) {
        rowHeightIndex[i] = v.size;
      }
    });

    const rows = Object.entries(this.cells).map(([row, cells]) => {
      const rowNumber = parseInt(row, 10) + 1;
      const ht = rowHeightIndex[rowNumber] ? ` ht="${rowHeightIndex[rowNumber]}" customHeight="1"` : '';
      const renderCells = Object.entries(cells).map(([column, cell]) => {
        return {
          value: cell.value,
          style: cell.style,
          position: `${ntoa(parseInt(column, 10) + 1)}${rowNumber}`,
          isString: typeof cell.value === 'string',
        };
      });

      return { number: rowNumber, ht, cells: renderCells };
    });

    return renderSheet(rows, this.mergeCells, colWidths);
  }
}
