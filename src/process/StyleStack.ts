import Stack from './Stack.ts';
import {
  renderBorderRule,
  renderCellXfsRule,
  renderFillRule,
  renderFontRule,
  renderFormatRule,
  type Style,
  type Styles,
} from '../render/renderStyles.ts';

/**
 * 样式栈，管理一个 WorkBook 中的所有样式
 */
export default class StyleStack {
  // 格式栈
  public readonly numFmtStack = new Stack(164); // 0-163 是系统保留的，自定义样式从 164 开始

  // 字体栈
  public readonly fontStack = new Stack();

  // 边框栈
  public readonly borderStack = new Stack();

  // 填充栈
  public readonly fillStack = new Stack(2); // 0-1 是系统保留的，自定义样式从 2 开始

  // 单元格样式栈
  public readonly cellXfsStack = new Stack();

  /**
   * 构造函数
   */
  constructor() {}

  /**
   * 处理一个样式，将其中的格式、字体、边框、填充等分别入栈并获取索引，
   * 然后用这些索引与对齐一起生成单元格样式规则并入栈，最终获取样式索引
   *
   * @param style - 样式
   * @private
   * @return - 样式索引
   */
  private process(style: Style): number {
    const numFmtId = style.formatCode ? this.numFmtStack.push(renderFormatRule(style.formatCode)) : 0;
    const fontId = style.font ? this.fontStack.push(renderFontRule(style.font)) : 0;
    const borderId = style.border ? this.borderStack.push(renderBorderRule(style.border)) : 0;
    const fillId = style.fill ? this.fillStack.push(renderFillRule(style.fill)) : 0;
    const cellXfsRule = renderCellXfsRule(numFmtId, fontId, fillId, borderId, style.alignment);
    return this.cellXfsStack.push(cellXfsRule);
  }

  /**
   * 处理一批样式
   *
   * @param styles - 一批样式
   * @return - 样式名称和单元格索引的映射关系
   */
  public consume(styles: Styles): Record<string, number> {
    return Object.fromEntries(
      Object.entries(styles).map(([k, v]) => {
        return [k, this.process(v)];
      }),
    );
  }

  /**
   * 获取所有样式规则
   *
   * @return - 格式、字体、边框、填充、单元格样式规则组成的对象
   */
  public getRules() {
    return {
      numFmtRules: this.numFmtStack.getRules(),
      fontRules: this.fontStack.getRules(),
      borderRules: this.borderStack.getRules(),
      fillRules: this.fillStack.getRules(),
      cellXfsRules: this.cellXfsStack.getRules(),
    };
  }
}
