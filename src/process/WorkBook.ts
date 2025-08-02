import WorkSheet from './WorkSheet.ts';
import StyleStack from './StyleStack.ts';
import renderRels from '../render/renderRels.ts';
import renderDocPropsApp from '../render/renderDocPropsApp.ts';
import renderContentType from '../render/renderContentType.ts';
import renderXlRels from '../render/renderXlRels.ts';
import renderWorkBook from '../render/renderWorkBook.ts';
import renderStyles from '../render/renderStyles.ts';
import { BlobWriter, TextReader, ZipWriter } from '@zip.js/zip.js';

export default class WorkBook {
  private readonly workSheets: WorkSheet[];
  private readonly styleStack = new StyleStack();
  private readonly sheetIndex: number[];

  constructor(name: string, workSheets: WorkSheet[]) {
    this.workSheets = workSheets;
    this.sheetIndex = workSheets.map((v, i) => i + 1);
  }

  private processStyles() {
    this.workSheets.forEach((sheet) => {
      sheet.styleIndex = this.styleStack.consume(sheet.styles);
    });
  }

  public async getZipBlob() {
    this.processStyles();
    const rels = renderRels();
    const docPropsApp = renderDocPropsApp(this.sheetIndex);
    const xlRels = renderXlRels(this.sheetIndex);
    const workBook = renderWorkBook(this.sheetIndex);
    const rules = this.styleStack.getRules();
    const styles = renderStyles(
      rules.numFmtRules,
      rules.fontRules,
      rules.fillRules,
      rules.borderRules,
      rules.cellXfsRules,
    );
    const workSheets = this.workSheets.map((sheet) => sheet.render());
    const contentType = renderContentType(this.sheetIndex);

    const zipWriter = new ZipWriter(new BlobWriter());
    await zipWriter.add('_rels/.rels', new TextReader(rels));
    await zipWriter.add('docProps/app.xml', new TextReader(docPropsApp));
    await zipWriter.add('xl/_rels/workbook.xml.rels', new TextReader(xlRels));
    await zipWriter.add('xl/workbook.xml', new TextReader(workBook));
    await zipWriter.add('xl/styles.xml', styles);
    for (let i = 0; i < workSheets.length; i++) {
      await zipWriter.add(`xl/worksheets/sheet${i + 1}.xml`, new TextReader(workSheets[i]));
    }
    await zipWriter.add('[Content_Types].xml', new TextReader(contentType));
    return await zipWriter.close();
  }

  async downloadAs(name: string) {
    const blob = await this.getZipBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    URL.revokeObjectURL(url);
  }
}
