import Mustache from 'mustache'
import templateString from '../templates/_rels/.rels.ts'

/**
 * 生成 _rels/.rels 的内容
 *
 * @returns - _rels/.rels 的内容
 */
export default (): string => {
  return Mustache.render(templateString, {});
};
