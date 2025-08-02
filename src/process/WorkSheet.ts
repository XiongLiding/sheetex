import { type Style } from '../render/renderStyles.ts';
import renderSheet from '../render/renderWorkSheet.ts';

export type DataCell =
  | number
  | string
  | {
      value: string | number;
      style: string;
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

export function columnAlphabetToNumber(column: string) {
  return column.split('').reduce((acc, char) => {
    return acc + char.charCodeAt(0) - 'A'.charCodeAt(0);
  }, 0);
}

export function columnNumberToAlphabet(column: number) {
  if (column === 0) {
    return 'A';
  }
  let result = '';
  while (column > 0) {
    const remainder = column % 26;
    result = String.fromCharCode(65 + remainder) + result;
    column = Math.floor(column / 26);
  }
  return result;
}

export default class WorkSheet {
  public readonly name: string;
  public readonly styles: Record<string, Style>;
  public styleIndex?: Record<string, number>;
  private readonly blocks: DataBlock | DataBlock[];
  private readonly cells: Record<number, Record<number, SheetCell>> = {};

  constructor(name: string, blocks: DataBlock | DataBlock[], styles: Record<string, Style>) {
    if (!name) throw new Error('工作表名称不能为空');
    const invalidName = /[:\\\/?*\[\]]/.test(name);
    if (invalidName) throw new Error('工作表名称包含非法字符');
    this.name = name.substring(0, 31);
    this.blocks = blocks;
    this.styles = styles;
  }

  private processCell(cell: number | string | DataCell): SheetCell {
    if (typeof cell === 'number') {
      return {
        value: cell,
        style: 0,
      };
    }
    if (typeof cell === 'string') {
      return {
        value: cell,
        style: 0,
      };
    }
    if (typeof cell === 'object' && cell.value) {
      return {
        value: cell.value,
        style: cell.style ? this.styleIndex[cell.style] : 0,
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
    const columnOffset = columnAlphabetToNumber(column);

    block.data.forEach((row, y) => {
      row.forEach((cell, x) => {
        if (!this.cells[y + rowOffset]) {
          this.cells[y + rowOffset] = {};
        }
        this.cells[y + rowOffset][x + columnOffset] = this.processCell(cell);
      });
    });
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

    const rows = Object.entries(this.cells).map(([row, cells]) => {
      const rowNumber = parseInt(row, 10) + 1;
      const renderCells = Object.entries(cells).map(([column, cell]) => {
        return {
          value: cell.value,
          style: cell.style,
          position: `${columnNumberToAlphabet(parseInt(column, 10))}${rowNumber}`,
          isString: typeof cell.value === 'string',
        };
      });

      return { number: rowNumber, cells: renderCells };
    });

    return renderSheet(rows);
  }
}
