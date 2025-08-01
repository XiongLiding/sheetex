import Mustache from 'mustache'
import templateString from '../templates/_rels/.rels.ts'

export default () => {
  return Mustache.render(templateString, {});
};
