import WorkSheet from './WorkSheet.ts';
import StyleStack from './StyleStack.ts';
import renderRels from '../render/renderRels.ts';
import renderDocPropsApp from '../render/renderDocPropsApp.ts';
import renderContentType from '../render/renderContentType.ts';
import renderXlRels from '../render/renderXlRels.ts';
import renderWorkBook from '../render/renderWorkBook.ts';
import renderStyles from '../render/renderStyles.ts';
import { strToU8, zip } from 'fflate';
import { type SheetInfo } from '../render/types.ts';

export default class WorkBook {
  private readonly workSheets: WorkSheet[];
  private readonly styleStack = new StyleStack();
  private readonly sheets: SheetInfo[];

  constructor(name: string, workSheets: WorkSheet[]) {
    this.workSheets = workSheets;
    this.sheets = workSheets.map((v, i) => {
      return {
        index: i + 1,
        name: v.name,
      };
    });
  }

  private processStyles() {
    this.workSheets.forEach((sheet) => {
      sheet.styleIndex = this.styleStack.consume(sheet.styles);
    });
  }

  public async getZipBuffer(): Promise<Uint8Array> {
    this.processStyles();
    const rels = renderRels();
    const docPropsApp = renderDocPropsApp(this.sheets);
    const xlRels = renderXlRels(this.sheets);
    const workBook = renderWorkBook(this.sheets);
    const rules = this.styleStack.getRules();
    const styles = renderStyles(
      rules.numFmtRules,
      rules.fontRules,
      rules.fillRules,
      rules.borderRules,
      rules.cellXfsRules,
    );
    const workSheets = this.workSheets.map((sheet) => sheet.render());
    const contentType = renderContentType(this.sheets);

    return new Promise((resolve, reject) => {
      zip(
        {
          _rels: {
            '.rels': strToU8(rels),
          },
          docProps: {
            'app.xml': strToU8(docPropsApp),
          },
          xl: {
            _rels: {
              'workbook.xml.rels': strToU8(xlRels),
            },
            'workbook.xml': strToU8(workBook),
            'styles.xml': strToU8(styles),
            worksheets: workSheets.reduce((acc, workSheet, i) => {
              acc[`sheet${i + 1}.xml`] = strToU8(workSheet);
              return acc;
            }, {}),
          },
          '[Content_Types].xml': strToU8(contentType),
        },
        (err, data: Uint8Array) => {
          if (err) {
            return reject(err);
          }

          return resolve(data);
        },
      );
    });
  }

  async getZipBlob() {
    const buffer = await this.getZipBuffer();
    const newBuffer = new ArrayBuffer(buffer.byteLength);
    new Uint8Array(newBuffer).set(buffer);
    return new Blob([newBuffer], {
      type: 'application/zip',
    });
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
