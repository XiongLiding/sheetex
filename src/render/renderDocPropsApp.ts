import Mustache from 'mustache'
import templateString from '../templates/docProps/app.xml.ts'
import { type SheetInfo } from './types.ts';

/**
 * 生成 docProps/app.xml 的内容
 *
 * @param sheets - 工作表信息
 * @returns docProps/app.xml 的内容
 */
export default (sheets: SheetInfo[]): string => {
  return Mustache.render(templateString, {
    sheets: sheets,
    sheetsCount: sheets.length,
  });
};
