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

/**
 * 工作簿管理
 */
export default class WorkBook {
  // 工作表，一个工作簿包含一个或多个工作表
  private readonly workSheets: WorkSheet[];

  // 样式栈，用于统一处理样式，所有工作表的样式最终都会入到这个统一的栈中
  private readonly styleStack = new StyleStack();

  // 工作表信息
  private readonly sheets: SheetInfo[];

  /**
   * 创建一个工作簿
   *
   * @constructor
   * @param workSheets - 一个或多个工作表
   */
  constructor(workSheets: WorkSheet[]) {
    this.workSheets = workSheets;
    this.sheets = workSheets.map((v, i) => {
      return {
        index: i + 1,
        name: v.name,
      };
    });
  }

  /**
   * 处理样式，将所有工作表的样式逐个丢给样式栈处理
   *
   * @private
   */
  private processStyles() {
    this.workSheets.forEach((sheet) => {
      sheet.styleIndex = this.styleStack.consume(sheet.styles);
    });
  }

  /**
   * 生成 Buffer 形式的 Excel 文件数据
   *
   * @return Promise<Uint8Array> - Excel 文件 Buffer 形式
   */
  public async getZipBuffer(): Promise<Uint8Array> {
    // 处理样式
    this.processStyles();

    // 生成 Excel 内部文件
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

    // 用 fflate 根据 Excel 文件内部格式压缩上述文件并得到 Buffer
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

  /**
   * 获取 blob 格式的 Excel 文件
   *
   * @returns {Promise<Blob>} - Excel 文件的 blob 形式
   */
  async getZipBlob() {
    const buffer = await this.getZipBuffer();
    const newBuffer = new ArrayBuffer(buffer.byteLength);
    new Uint8Array(newBuffer).set(buffer);
    return new Blob([newBuffer], {
      type: 'application/zip',
    });
  }

  /**
   * 在浏览器端保存为文件
   *
   * @param name - 文件名
   */
  async saveAs(name: string) {
    const blob = await this.getZipBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    URL.revokeObjectURL(url);
  }

  /**
   * 在浏览器端下载为文件
   *
   * @deprecated
   * @param name - 文件名
   */
  async downloadAs(name: string) {
    await this.saveAs(name);
  }
}
