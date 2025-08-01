import relsRender from '../src/render/relsRender.ts'
import docPropsAppRender from '../src/render/docPropsAppRender.ts'
import xlRelsRender from '../src/render/xlRelsRender.ts'
import sheetRender from '../src/render/sheetRender.ts'
import stylesRender from '../src/render/stylesRender.ts'
import workbookRender from '../src/render/workbookRender.ts'
import contentTypeRender from '../src/render/contentTypeRender.ts'

console.log(relsRender())
console.log(docPropsAppRender([1, 2, 3]))
console.log(xlRelsRender([1, 2, 3]))
console.log(
  sheetRender([{ number: 1, cells: [{ column: 'A', isString: true, isNumber: false, value: 'Hello World!' }] }]))
console.log(stylesRender([
  { fontSize: 12, fontFamily: '宋体' }], [
  {
    color: '#cccccc',
  }], [
  {
    left: '', right: '', top: '', bottom: '', diagonal: '',
  }], [], []))
console.log(workbookRender([1, 2, 3]))
console.log(contentTypeRender([1, 2, 3]))
