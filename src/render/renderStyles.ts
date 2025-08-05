import Mustache from 'mustache';
import templateString from '../templates/xl/styles.xml.ts';

const alignmentHorizontal = {
  left: 1,
  center: 2,
  right: 3,
  fill: 4,
  justify: 5,
  centerContinuous: 6,
  distributed: 7,
};

const alignmentVertical = {
  top: 1,
  center: 2,
  bottom: 3,
  justify: 4,
  distributed: 5,
};

const fontUnderline = {
  single: 1,
  double: 2,
  singleAccounting: 3,
  doubleAccounting: 4,
};

const fontVertAlign = {
  superscript: 1,
  subscript: 2,
};

const borderLineStyle = {
  none: 1,
  hair: 2,
  dotted: 3,
  dashDotDot: 4,
  dashDot: 5,
  dashed: 6,
  thin: 7,
  mediumDashDotDot: 8,
  slantDashDot: 9,
  mediumDashDot: 10,
  mediumDashed: 11,
  medium: 12,
  thick: 13,
  double: 14,
};

const fillPatternType = {
  solid: 1,
  darkHorizontal: 2,
  lightHorizontal: 3,
  darkGray: 4,
  darkVertical: 5,
  lightVertical: 6,
  mediumGray: 7,
  darkDown: 8,
  lightDown: 9,
  lightGray: 10,
  darkUp: 11,
  lightUp: 12,
  gray125: 13,
  darkGrid: 14,
  lightGrid: 15,
  gray0625: 16,
  darkTrellis: 17,
  lightTrellis: 18,
};

export interface Alignment {
  horizontal?: keyof typeof alignmentHorizontal;
  vertical?: keyof typeof alignmentVertical;
  indent?: number;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
}

export interface Font {
  sz?: number;
  name?: string;
  color?: string;
  b?: boolean;
  i?: boolean;
  u?: keyof typeof fontUnderline | false;
  strike?: boolean;
  vertAlign?: keyof typeof fontVertAlign | false;
}

export interface BorderStyle {
  style: keyof typeof borderLineStyle;
  color: string;
}

export interface Border {
  style?: keyof typeof borderLineStyle;
  color?: string;
  diagonalUp?: boolean;
  diagonalDown?: boolean;

  left?: BorderStyle;
  top?: BorderStyle;
  right?: BorderStyle;
  bottom?: BorderStyle;
  diagonal?: BorderStyle;
}

export interface Fill {
  patternType: keyof typeof fillPatternType | 'none';
  fgColor?: string;
  bgColor?: string;
}

export interface Style {
  formatCode?: string;
  font?: Font;
  border?: Border;
  fill?: Fill;
  alignment?: Alignment;
}

export type Styles = Record<string, Style>;

export interface CellXfs {
  numFmtId: number;
  numFontId: number;
  fillId: number;
  borderId: number;
  xfId: number;
}

/**
 * 判断是否为颜色，合法的颜色格式为 RRGGBB 或 AARRGGBB
 * @param color
 */
function isColor(color: string) {
  return /^([0-9A-F]{6}|[0-9A-F]{8})$/.test(color);
}

/**
 * 渲染格式规则
 *
 * @param formatCode 格式码
 * @return - 格式规则
 */
export function renderFormatRule(formatCode: string) {
  return Mustache.render('<numFmt numFmtId="$$id$$" formatCode="{{.}}"/>', formatCode);
}

/**
 * 渲染字体规则
 *
 * @param font - 字体样式
 * @return - 字体规则
 */
export function renderFontRule(font: Font) {
  const sz = font.sz && typeof font.sz === 'number' ? `<sz val="${font.sz}"/>` : '';
  const name = font.name ? Mustache.render('<name val="{{.}}"/>', font.name) : '';
  const color =
    font.color && isColor(font.color) ? `<color rgb="${(font.color.length === 6 ? 'FF' : '') + font.color}"/>` : '';
  const b = font.b ? '<b/>' : '';
  const i = font.i ? '<i/>' : '';
  const u = font.u && fontUnderline[font.u] ? `<u val="${font.u}"/>` : '';
  const strike = font.strike ? '<strike/>' : '';
  const vertAlign = font.vertAlign && fontVertAlign[font.vertAlign] ? `<vertAlign val="${font.vertAlign}"/>` : '';
  return `<font>${sz}${name}${color}${b}${i}${u}${strike}${vertAlign}</font>`;
}

/**
 * 渲染边框样式
 *
 * @param border - 边框样式
 * @return - 边框规则
 */
export function renderBorderRule(border: Border) {
  const inner = ['left', 'right', 'top', 'bottom', 'diagonal']
    .map((v) => {
      if (!border.diagonalUp && !border.diagonalDown && v === 'diagonal') {
        return `<${v}/>`;
      }

      const style = border[v]?.style ?? border.style ?? '';

      if (!style || style === 'none' || !borderLineStyle[style]) {
        return `<${v}/>`;
      }

      const styleAttr = style ? ` style="${style}"` : '';

      const color = border[v]?.color ?? border.color ?? '';
      const colorTag = color && isColor(color) ? `<color rgb="${(color.length === 6 ? 'FF' : '') + color}" />` : '';
      return `<${v}${styleAttr}>${colorTag}</${v}>`;
    })
    .join('');

  const diagonalUp = border.diagonalUp ? ' diagonalUp="1"' : '';
  const diagonalDown = border.diagonalDown ? ' diagonalDown="1"' : '';
  return `<border${diagonalUp}${diagonalDown}>${inner}</border>`;
}

/**
 * 渲染填充规则
 *
 * @param fill - 填充样式
 * @return 填充规则
 */
export function renderFillRule(fill: Fill) {
  const patternType = fill.patternType && fillPatternType[fill.patternType] ? fill.patternType : 'none';

  const fgColor =
    fill.fgColor && isColor(fill.fgColor)
      ? `<fgColor rgb="${(fill.fgColor.length === 6 ? 'FF' : '') + fill.fgColor}"/>`
      : '<fgColor auto="1"/>';
  const bgColor =
    fill.bgColor && isColor(fill.bgColor)
      ? `<bgColor rgb="${(fill.bgColor.length === 6 ? 'FF' : '') + fill.bgColor}"/>`
      : '<bgColor auto="1"/>';

  return `<fill><patternFill patternType="${patternType}">${fgColor}${bgColor}</patternFill></fill>`;
}

/**
 * 渲染对齐规则
 *
 * @param alignment 对齐样式
 * @return - 对齐规则
 */
export function renderAlignmentRule(alignment: Alignment) {
  const horizontal =
    alignment.horizontal && alignmentHorizontal[alignment.horizontal] ? ` horizontal="${alignment.horizontal}"` : '';
  const vertical =
    alignment.vertical && alignmentVertical[alignment.vertical] ? ` vertical="${alignment.vertical}"` : '';
  const indent =
    ['left', 'right', 'distributed'].includes(alignment.horizontal) &&
    alignment.indent &&
    typeof alignment.indent === 'number'
      ? ` indent="${alignment.indent}"`
      : '';
  const textRotation = alignment.textRotation ? ` textRotation="${alignment.textRotation}"` : '';
  const wrapText = alignment.wrapText ? ' wrapText="1"' : '';
  const shrinkToFit = alignment.shrinkToFit ? ' shrinkToFit="1"' : '';
  return `<alignment${horizontal}${vertical}${indent}${textRotation}${wrapText}${shrinkToFit}/>`;
}

/**
 * 渲染单元格样式规则
 *
 * @param numFmtId - 格式索引
 * @param fontId - 字体索引
 * @param fillId - 填充索引
 * @param borderId - 边框索引
 * @param alignment - 对齐样式
 * @return 样式规则
 */
export function renderCellXfsRule(
  numFmtId: number,
  fontId: number,
  fillId: number,
  borderId: number,
  alignment: Alignment,
) {
  const alignmentRule = alignment ? renderAlignmentRule(alignment) : '';
  return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0">${alignmentRule}</xf>`;
}

/**
 * 生成 style.xml 的内容
 *
 * @param numFmts - 所有格式规则
 * @param fonts - 所有字体规则
 * @param fills - 所有填充规则
 * @param borders - 所有边框规则
 * @param cellXfs - 所有单元格样式规则
 * @return - style.xml 的内容
 */
export default (numFmts: string[], fonts: string[], fills: string[], borders: string[], cellXfs: string[]) => {
  return Mustache.render(templateString, {
    numFmts,
    numFmtsCount: numFmts.length,
    fonts,
    fontsCount: fonts.length + 1,
    fills,
    fillsCount: fills.length + 2,
    borders,
    bordersCount: borders.length + 1,
    cellXfs,
    cellXfsCount: cellXfs.length + 1,
  });
};
