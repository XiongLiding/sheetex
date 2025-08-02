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

export default (rows: RenderRow[]): string => {
  return Mustache.render(templateString, {
    rows,
  });
};
