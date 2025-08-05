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

/**
 * 将列名转换为列序号，A -> 1
 *
 * @param column - 列名
 * @returns 列序号
 */
export function aton(column: string) {
  return column
    .split('')
    .reverse()
    .reduce((acc, char, i) => {
      return acc + (char.charCodeAt(0) - 'A'.charCodeAt(0) + (i >= 1 ? 1 : 0)) * 26 ** i;
    }, 1);
}

/**
 * 将列序号转换为列名，1 -> A
 *
 * @param column
 * @returns 列名
 */
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

/**
 * 工作表管理
 */
export default class WorkSheet {
  // 工作表名称
  public readonly name: string;

  // 工作表样式
  public readonly styles: Record<string, Style>;

  // 样式名称与样式索引的映射关系
  public styleIndex?: Record<string, number>;

  // 工作表数据块
  private readonly blocks: DataBlock[];

  // 工作表内的单元格
  private readonly cells: Record<number, Record<number, SheetCell>> = {};

  // 合并的单元格
  private readonly mergeCells: string[];

  // 列宽
  private readonly colWidths: Size[];

  // 行高
  private readonly rowHeights: Size[];

  /**
   * 构造函数
   *
   * @constructor
   * @param name - 工作表名称
   * @param blocks - 数据块
   * @param styles - 工作表样式
   * @param options - 工作表选项
   */
  constructor(name: string, blocks: DataBlock[], styles: Record<string, Style> = {}, options: SheetOptions = {}) {
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

  /**
   * 处理单元格，统一转换成对象形式
   * 当前单元格若未定义样式，则让其使用默认样式
   *
   * @param cell - 单元格，可以是数字、字符串或对象
   * @private
   * @return 统一的对象形式的单元格
   */
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

  /**
   * 批量处理数据单元格
   *
   * @param block 数据块
   * @private
   */
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

  /**
   * 将行高和列宽处理成标准格式
   *
   * @param sizes - 行高和列宽
   * @private
   * @return 标准化处理后的行高和列宽
   */
  private processSize(sizes: Size[]) {
    const renderSize: RenderSize[] = [];

    sizes.forEach((v) => {
      // size 是单个数字，可忽略 max，此时 max = min
      if (typeof v.size === 'number') {
        renderSize.push({
          min: v.min,
          max: v.max ?? v.min,
          size: v.size,
        });
      }

      // v.size 是数组
      if (Array.isArray(v.size)) {
        if (!v.max || v.max < v.min) {
          // max 未定义或无效，根据数组长度生成一批列定义
          v.size.forEach((size, i) => {
            if (typeof size !== 'number' || size < 0) return;

            renderSize.push({
              min: v.min + i,
              max: v.min + i,
              size,
            });
          });
        } else {
          // max 有效，在范围内用数组循环填充
          for (let i = v.min; i <= v.max; i++) {
            const size = v.size[(i - v.min) % v.size.length];
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

  /**
   * 生成 sheet{n}.xml 文件
   *
   * @return - sheet{n}.xml 文件内容
   */
  public render() {
    if (!this.styleIndex) {
      throw new Error('调用顺序错误：完成样式预处理才能进行渲染');
    }

    this.blocks.forEach((block) => {
      this.consumeDataBlock(block);
    });

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
