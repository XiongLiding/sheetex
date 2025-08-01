import Mustache from 'mustache'
import templateString from '../templates/xl/worksheets/sheet.xml.ts'

export interface Cell {
  column: string;
  isString: boolean;
  isNumber: boolean;
  value: string | number;
}

export interface Row {
  number: number;
  cells: Cell[];
}

export default (rows: Row[]) => {
  return Mustache.render(templateString, {
    rows,
  })
};
