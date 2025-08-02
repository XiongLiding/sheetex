import Mustache from 'mustache'
import templateString from '../templates/_rels/.rels.ts'

export default (): string => {
  return Mustache.render(templateString, {});
};
