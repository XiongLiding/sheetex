import Stack from './Stack.ts';
import {
  renderBorderRule,
  renderCellXfsRule,
  renderFillRule,
  renderFontRule, renderFormatRule,
  type Style,
} from '../render/renderStyles.ts';

export default class StyleStack {
  public readonly numFmtStack = new Stack(164);
  public readonly fontStack = new Stack();
  public readonly borderStack = new Stack();
  public readonly fillStack = new Stack();
  public readonly cellXfsStack = new Stack();

  constructor () {
  }

  private process (style: Style): number {
    const numFmtRule = renderFormatRule(style.formatCode);
    const numFmtId = this.numFmtStack.push(numFmtRule);
    const fontRule = renderFontRule(style.font);
    const fontId = this.fontStack.push(fontRule);
    const borderRule = renderBorderRule(style.border);
    const borderId = this.borderStack.push(borderRule);
    const fillRule = renderFillRule(style.fill);
    const fillId = this.fillStack.push(fillRule);
    const cellXfsRule = renderCellXfsRule(numFmtId, fontId, borderId, fillId, style.alignment);
    return this.cellXfsStack.push(cellXfsRule);
  }

  public consume (styles: Record<string, Style>) {
    return Object.fromEntries(Object.entries(styles).map(([k, v]) => {
      return [k, this.process(v)];
    }));
  }

  public getRules () {
    return {
      numFmtRules: this.numFmtStack.getRules(),
      fontRules: this.fontStack.getRules(),
      borderRules: this.borderStack.getRules(),
      fillRules: this.fillStack.getRules(),
      cellXfsRules: this.cellXfsStack.getRules(),
    };
  }
}
