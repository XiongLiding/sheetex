import Mustache from 'mustache'
import templateString from '../templates/[Content_Types].xml.ts'

export default (sheets: number[]): string => {
  return Mustache.render(templateString, {
    sheets: sheets,
  });
};
