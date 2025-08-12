import { type Styles } from '../render/renderStyles.ts';
import renderSheet, { type RenderSize } from '../render/renderWorkSheet.ts';

/**
 * 数据单元格
 *
 * @example
 * ```javascript
 * const cellA1 = 82;
 * const cellA2 = 'Hello';
 * const cellA3 = {value: 'Hello', style: 'bold'};
 * ```
 *
 * @remarks
 * 可以是简单的数字、字符串，也可以是带样式名称的对象
 */
export type DataCell =
  | number
  | string
  | {
      value: number | string;
      style?: string;
    };

/**
 * 数据行
 *
 * 数据行是数据单元格组成的数组
 *
 * @example
 * ```javascript
 * const row1 = ['hello', 82];
 * const row2 = [{value: 'hello', style: 'bold'}, 'world'];
 * ```
 */
export type DataRow = DataCell[];

/**
 * 数据块
 *
 * 包含数据行组成的数组和一个起始地址
 *
 * @example
 * ```javascript
 * const block = {
 *   origin: 'A1',
 *   data: [
 *     ['hello', 82],
 *     [{value: 'hello', style: 'bold'}, 'world']
 *   ]
 * };
 * ```
 */
export type DataBlock = {
  /**
   * 数据块的起始地址，数据块左上角的单元格落在这个位置
   *
   * @example
   * ```javascript
   * const cellAddress = 'A1',
   * ```
   */
  origin: string;

  /**
   * 数据块中的内容，由多个数据行组成
   */
  data: DataRow[];
};

/**
 * 工作表单元格
 */
export type SheetCell = {
  /**
   * 单元格值，最终写入 Excel 单元格的数据，可以是数字和字符串
   */
  value: number | string;

  /**
   * 单元格样式名称，单元格根据样式名称套用相应的样式规则进行展现
   */
  style: number;
};

/**
 * 将列名转换为列序号，A -> 1
 *
 * @example
 * ```javascript
 * const number = aton('AA'); // number = 27
 * ```
 *
 * @param column - 列名
 * @return - 列序号
 *
 * @version 1.1.0
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
 * @example
 * ```javascript
 * const name = ntoa(27); // column = 'AA'
 * ```
 *
 * @param column - 列序号
 * @return - 列名
 *
 * @version 1.1.0
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

/**
 * 定义一组行高或列宽
 *
 * - 当 max 填写且 size 为数字时，表示从 min 到 max 的行高/列宽都是 size；
 * - 当 max 填写且 size 为数组时，表示从 min 到 max 的行高/列宽依次如 size 数组所述，数组过长会被截断，过短会从头开始重复继续填充；
 * - 当 max 不填且 size 为数字时，表示 min 列的行高/列宽是 size；
 * - 当 max 不填且 size 为数组时，表示从 min 开始，行高/列宽依次如 size 数组中所述，直到数组用尽。
 *
 * @example
 * ```javascript
 * const size = {min: 1, max: 5, size: 20};
 * ```
 *
 * @version 1.0.0
 */
export interface Size {
  /**
   * 从第几行/列（含）开始
   */
  min: number;
  /**
   * 到第几行/列（含）结束
   */
  max?: number;
  /**
   * 行高或列宽
   */
  size: number | number[];
}

/**
 * 工作表选项
 *
 * @example
 * ```typescript
 * const options: SheetOptions = {
 *   mergeCells: ['A1:B2', 'C2:F8'],
 *   colWidths: [{min: 1, max: 5, size: 20}],
 *   rowHeights: [{min: 1, max: 5, size: 20}]
 * }
 * ```
 */
export interface SheetOptions {
  /**
   * 需要合并的范围
   *
   * @defaultValue []
   */
  mergeCells?: string[];

  /**
   * 工作表列宽定义
   *
   * @defaultValue []
   */
  colWidths?: Size[];

  /**
   * 工作表行高定义
   *
   * @defaultValue []
   */
  rowHeights?: Size[];
}

/**
 * 工作表管理
 * @version 1.0.0
 */
export default class WorkSheet {
  /**
   * 工作表名称
   * @internal
   */
  public readonly name: string;

  /**
   * 工作表样式
   * @internal
   */
  public readonly styles: Styles;

  /**
   * 样式名称与样式索引的映射关系
   * @internal
   */
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
   * @param name - 工作表名称
   * @param blocks - 数据块
   * @param styles - 工作表样式
   * @param options - 工作表选项
   *
   * @example
   * 1. 生成一个只有数据，不带任何样式的工作表
   * ```javascript
   * const block = {
   *   origin: 'A1',
   *   data: [['hello', 'world']]
   * }
   * const ws = new WorkSheet('Sheet1', [block]);
   * ```
   *
   * 2. 生成一个带样式的工作表（hello 被加粗）
   * ```javascript
   * const block = {
   *   origin: 'A1',
   *   data: [[{value: 'hello', style: 'bold'}, 'world']]
   * }
   * cost styles = {
   *   bold: {
   *     font: {
   *       b: true
   *     }
   *   }
   * }
   * const ws = new WorkSheet('Sheet2', [block], styles);
   * ```
   *
   * 3. 生成一个配置了列宽的工作表
   * ```javascript
   * const block = {
   *   origin: 'A1',
   *   data: [['hello', 'world']]
   * };
   * const options = {
   *   colWidths: {min: 1, max: 2, size: 40}
   * };
   * const ws = new WorkSheet('Sheet3', [block], {}, options);
   * ```
   *
   * @version 1.0.0
   */
  constructor(name: string, blocks: DataBlock[], styles: Styles = {}, options: SheetOptions = {}) {
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
   * @return 统一的对象形式的单元格
   * @private
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
      throw new Error('单元格地址格式错误');
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
   * @internal
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
