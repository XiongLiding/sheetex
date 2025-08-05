import Mustache from 'mustache'
import templateString from '../templates/[Content_Types].xml.ts'
import { type SheetInfo } from './types.ts';

/**
 * 生成 [Content_Types].xml 的内容
 *
 * @param sheets - 工作表信息
 * @return - [Content_Types].xml 的内容
 */
export default (sheets: SheetInfo[]): string => {
  return Mustache.render(templateString, {
    sheets: sheets,
  });
};
