import WorkBook from './process/WorkBook.ts';
import WorkSheet, { type DataRow, ntoa, type SheetOptions } from './process/WorkSheet.ts';
import { type Alignment, type Border, type Styles } from './render/renderStyles.ts';

/**
 * 需要快速保存的数据
 *
 * @example
 * ```javascript
 * const data = {
 *   caption: '学生成绩',
 *   headers: ['序号', '姓名', '成绩'],
 *   data: [
 *     [1, '张三', 100],
 *     [2, '李四', 90],
 *     [3, '王五', 80],
 *   ],
 * }
 * ```
 */
export interface JustData {
  /**
   * 表标题
   *
   * @remarks
   * 表标题将被以更大的字体加粗显示在第一行，并且合并居中显示，合并的列数参照表头（如果存在）或表数据第一行的列数
   */
  caption?: string;
  /**
   * 表头
   *
   * @remarks
   * 表头将被显示在第一或第二行（根据 caption 是否存在），文字加粗居中显示
   */
  headers?: string[];
  /**
   * 表数据
   */
  data: DataRow[];
}

/**
 * 快速生成工作簿
 *
 * @param data - 数据
 * @param colWidths - 列宽
 * @return - 工作簿
 */
function justWorkBook(data: DataRow[] | JustData, colWidths: number[] = []) {
  const blocks = {
    origin: 'A1',
    data: [],
  };

  const options: SheetOptions = {};

  if (Array.isArray(data)) {
    blocks.data = data;
  } else {
    if (data.caption) {
      const titleRow = (data.headers || data.data[0] || ['']).map(() => ({ value: '', style: 'caption' }));
      options.mergeCells = [`A1:${ntoa(titleRow.length)}1`];
      options.rowHeights = [{ min: 1, size: 24 }];
      titleRow[0].value = data.caption;
      blocks.data.push(titleRow);
    }
    if (data.headers) {
      blocks.data.push(data.headers.map((v) => ({ value: v, style: 'header' })));
    }
    blocks.data.push(...data.data);
  }

  if (colWidths && colWidths.length) {
    options.colWidths = [{ min: 1, size: colWidths }];
  }

  const alignment: Alignment = {
    horizontal: 'center',
    vertical: 'center',
    wrapText: true,
  };

  const border: Border = {
    style: 'thin',
    color: '000000',
  };

  const styles: Styles = {
    default: {
      font: {
        sz: 12,
      },
      alignment,
      border,
    },
    caption: {
      font: {
        b: true,
        sz: 18,
      },
      alignment,
      border,
    },
    header: {
      font: {
        b: true,
        sz: 12,
      },
      alignment,
      border,
    },
  };

  const ws = new WorkSheet('Sheet1', [blocks], styles, options);
  return new WorkBook([ws]);
}

/**
 * 快速生成工作簿 Buffer
 *
 * 根据用户提供的数据以及列宽配置，生成一个带有默认样式的 Excel 文件并得到相应的 Buffer
 *
 * @param data - 表格数据
 * @param colWidths - 列宽
 * @return - 生成的 Excel 文件的 Buffer
 *
 * @remarks
 * 默认样式是指：
 * 1. 标题加大加粗合并单元格居中显示
 * 2. 表头加粗居中显示
 * 3. 数据居中显示
 * 4. 所有内容加上黑色细线边框
 *
 * @example
 * 数据
 * ```javascript
 * const data = [
 *   [1, '张三', 100],
 *   [2, '李四', 90],
 *   [3, '王五', 80],
 * ];
 * ```
 * 1. 仅数据
 * ```javascript
 * const buffer = await justBuffer(data);
 * ```
 * 2. 指定标题、表头并设置列宽
 * ```javascript
 * const buffer = await justBuffer({
 *   caption: '成绩单',
 *   header: ['学号', '姓名', '成绩'],
 *   data,
 * }, [20, 20, 20]);
 * ```
 *
 * @since 1.3.0
 */
export async function justBuffer(data: DataRow[] | JustData, colWidths: number[] = []) {
  const wb = justWorkBook(data, colWidths);
  return await wb.getZipBuffer();
}

/**
 * 快速生成工作簿 Blob
 *
 * 根据用户提供的数据以及列宽配置，生成一个带有默认样式的 Excel 文件并得到相应的 Blob
 *
 * @param data - 表格数据
 * @param colWidths - 列宽
 * @return - 生成的 Excel 文件的 Blob
 *
 * @remarks
 * 默认样式是指：
 * 1. 标题加大加粗合并单元格居中显示
 * 2. 表头加粗居中显示
 * 3. 数据居中显示
 * 4. 所有内容加上黑色细线边框
 *
 * @example
 * 数据
 * ```javascript
 * const data = [
 *   [1, '张三', 100],
 *   [2, '李四', 90],
 *   [3, '王五', 80],
 * ];
 * ```
 * 1. 仅数据
 * ```javascript
 * const blob = await justBlob(data);
 * ```
 * 2. 指定标题、表头并设置列宽
 * ```javascript
 * const blob = await justBlob({
 *   caption: '成绩单',
 *   header: ['学号', '姓名', '成绩'],
 *   data,
 * }, [20, 20, 20]);
 * ```
 *
 * @since 1.3.0
 */
export async function justBlob(data: DataRow[] | JustData, colWidths: number[] = []) {
  const wb = justWorkBook(data, colWidths);
  return await wb.getZipBlob();
}

/**
 * 在浏览器中快速保存工作簿
 *
 * 根据用户提供的文件名、数据以及列宽配置，生成一个带有默认样式的 Excel 文件并保存
 *
 * @param name - 文件保存名称
 * @param data - 表格数据
 * @param colWidths - 列宽
 *
 * @remarks
 * 默认样式是指：
 * 1. 标题加大加粗合并单元格居中显示
 * 2. 表头加粗居中显示
 * 3. 数据居中显示
 * 4. 所有内容加上黑色细线边框
 *
 * @example
 * 数据
 * ```javascript
 * const data = [
 *   [1, '张三', 100],
 *   [2, '李四', 90],
 *   [3, '王五', 80],
 * ];
 * ```
 * 1. 仅数据
 * ```javascript
 * await justSaveAs('成绩单.xlsx', data);
 * ```
 * 2. 指定标题、表头并设置列宽
 * ```javascript
 * await justSaveAs('成绩单.xlsx', {
 *   caption: '成绩单',
 *   header: ['学号', '姓名', '成绩'],
 *   data,
 * }, [20, 20, 20]);
 *```
 *
 * @since 1.3.0
 */
export async function justSaveAs(name: string, data: DataRow[] | JustData, colWidths: number[] = []) {
  const wb = justWorkBook(data, colWidths);
  return await wb.saveAs(name);
}
