import Mustache from 'mustache';
import templateString from '../templates/xl/_rels/workbook.xml.rels.ts';
import { type SheetInfo } from './types.ts';

/**
 * 生成 xl/_rels/workbook.xml.rels 的内容
 *
 * @param sheets - 工作表信息
 * @returns - xl/_rels/workbook.xml.rels 的内容
 */
export default (sheets: SheetInfo[]): string => {
  return Mustache.render(templateString, {
    sheets,
    styleId: sheets.length + 1,
  });
};
