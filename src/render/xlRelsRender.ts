import Mustache from 'mustache'
import templateString from '../templates/xl/_rels/workbook.xml.rels.ts'

export default (sheets: number []) => {
  return Mustache.render(templateString, {
    sheets,
    styleId: sheets.length + 1,
  });
};
