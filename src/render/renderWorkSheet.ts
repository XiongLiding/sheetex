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

export default (rows: RenderRow[], mergeCells: string[], cols: RenderSize[]): string => {
  return Mustache.render(templateString, {
    rows,
    mergeCells: mergeCells,
    mergeCellsCount: mergeCells.length,
    cols: cols,
    colsCount: cols.length,
  });
};
