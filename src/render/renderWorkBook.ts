import Mustache from 'mustache'
import templateString from '../templates/xl/workbook.xml.ts'
import { type SheetInfo } from './types.ts';

/**
 * 生成 xl/workbook.xml 的内容
 *
 * @param sheets - 工作表信息
 * @return - xl/workbook.xml 的内容
 */
export default (sheets: SheetInfo[]): string => {
  return Mustache.render(templateString, {
    sheets: sheets,
  });
};
