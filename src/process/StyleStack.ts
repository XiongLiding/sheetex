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

export default class StyleStack {
  public readonly numFmtStack = new Stack(164);
  public readonly fontStack = new Stack();
  public readonly borderStack = new Stack();
  public readonly fillStack = new Stack();
  public readonly cellXfsStack = new Stack();

  constructor() {}

  private process(style: Style): number {
    const numFmtId = style.formatCode ? this.numFmtStack.push(renderFormatRule(style.formatCode)) : 0;
    const fontId = style.font ? this.fontStack.push(renderFontRule(style.font)) : 0;
    const borderId = style.border ? this.borderStack.push(renderBorderRule(style.border)) : 0;
    const fillId = style.fill ? this.fillStack.push(renderFillRule(style.fill)) : 0;
    const cellXfsRule = renderCellXfsRule(numFmtId, fontId, fillId, borderId, style.alignment);
    return this.cellXfsStack.push(cellXfsRule);
  }

  public consume(styles: Styles) {
    return Object.fromEntries(
      Object.entries(styles).map(([k, v]) => {
        return [k, this.process(v)];
      }),
    );
  }

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
