import Mustache from 'mustache'
import templateString from '../templates/[Content_Types].xml.ts'

export default (sheets: number[]) => {
  return Mustache.render(templateString, {
    sheets: sheets,
  });
};
