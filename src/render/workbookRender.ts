import Mustache from 'mustache'
import templateString from '../templates/xl/workbook.xml.ts'

export default (sheets: number[]) => {
  return Mustache.render(templateString, {
    sheets: sheets,
  });
};
