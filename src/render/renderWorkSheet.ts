import Mustache from 'mustache';
import templateString from '../templates/xl/worksheets/sheet.xml.ts';

export interface RenderCell {
  value: string | number;
  style: number;
  position: string;
  isString: boolean;
}

export interface RenderRow {
  number: number;
  cells: RenderCell[];
}

export interface RenderSize {
  min: number;
  max: number;
  size: number;
}

/**
 * 生成 xl/worksheets/sheet{n}.xml 的内容
 *
 * @param rows - 表格数据
 * @param mergeCells - 需要合并的单元格
 * @param cols - 列宽
 *
 * @returns xl/worksheets/sheet{n}.xml 的内容
 */
export default (rows: RenderRow[], mergeCells: string[], cols: RenderSize[]): string => {
  return Mustache.render(templateString, {
    rows,
    mergeCells: mergeCells,
    mergeCellsCount: mergeCells.length,
    cols: cols,
    colsCount: cols.length,
  });
};
