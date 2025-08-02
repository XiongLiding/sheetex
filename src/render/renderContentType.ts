import Mustache from 'mustache'
import templateString from '../templates/[Content_Types].xml.ts'
import { type SheetInfo } from './types.ts';

export default (sheets: SheetInfo[]): string => {
  return Mustache.render(templateString, {
    sheets: sheets,
  });
};
