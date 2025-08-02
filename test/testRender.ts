import relsRender from '../src/render/renderRels.ts';
import docPropsAppRender from '../src/render/renderDocPropsApp.ts';
import xlRelsRender from '../src/render/renderXlRels.ts';
import sheetRender from '../src/render/renderWorkSheet.ts';
import stylesRender from '../src/render/renderStyles.ts';
import workbookRender from '../src/render/renderWorkBook.ts';
import contentTypeRender from '../src/render/renderContentType.ts';

console.log(relsRender());
console.log(docPropsAppRender([1, 2, 3]));
console.log(xlRelsRender([1, 2, 3]));
console.log(sheetRender([]));
console.log(stylesRender([], [], [], [], []));
console.log(workbookRender([1, 2, 3]));
console.log(contentTypeRender([1, 2, 3]));
