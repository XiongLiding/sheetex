import Mustache from 'mustache'
import templateString from '../templates/xl/_rels/workbook.xml.rels.ts'
import { type SheetInfo } from './types.ts';

export default (sheets: SheetInfo[]): string => {
  return Mustache.render(templateString, {
    sheets,
    styleId: sheets.length + 1,
  });
};
