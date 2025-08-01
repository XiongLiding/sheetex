import Mustache from 'mustache'
import templateString from '../templates/docProps/app.xml.ts'

export default (sheets: number[]) => {
  return Mustache.render(templateString, {
    sheets: sheets,
    sheetsCount: sheets.length,
  });
};
