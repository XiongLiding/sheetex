import Mustache from 'mustache';
import templateString from '../templates/xl/styles.xml.ts';

export interface Alignment {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
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
  u?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | false;
  strike?: boolean;
  vertAlign?: 'superscript' | 'subscript' | false;
}

export interface BorderStyle {
  style:
    | 'none'
    | 'dashDot'
    | 'dashDotDot'
    | 'dashed'
    | 'dotted'
    | 'double'
    | 'hair'
    | 'medium'
    | 'mediumDashDot'
    | 'mediumDashDotDot'
    | 'mediumDashed'
    | 'slantDashDot'
    | 'thick'
    | 'thin';
  color: string;
}

export interface Border {
  style?:
    | 'none'
    | 'dashDot'
    | 'dashDotDot'
    | 'dashed'
    | 'dotted'
    | 'double'
    | 'hair'
    | 'medium'
    | 'mediumDashDot'
    | 'mediumDashDotDot'
    | 'mediumDashed'
    | 'slantDashDot'
    | 'thick'
    | 'thin';
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
  patternType:
    | 'none'
    | 'solid'
    | 'darkGray'
    | 'mediumGray'
    | 'lightGray'
    | 'gray125'
    | 'gray0625'
    | 'darkHorizontal'
    | 'darkVertical'
    | 'darkDown'
    | 'darkUp'
    | 'darkGrid'
    | 'darkTrellis'
    | 'lightHorizontal'
    | 'lightVertical'
    | 'lightDown'
    | 'lightUp'
    | 'lightGrid'
    | 'lightTrellis';
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

export function renderFormatRule(formatCode: string) {
  return `<numFmt numFmtId="{{id}}" formatCode="${formatCode}"/>`;
}

export function renderFontRule(font: Font) {
  const sz = font.sz ? `<sz val="${font.sz}"/>` : '';
  const name = font.name ? `<name val="${font.name}"/>` : '';
  const color = font.color ? `<color rgb="${(font.color.length === 6 ? 'FF' : '') + font.color}"/>` : '';
  const b = font.b ? '<b/>' : '';
  const i = font.i ? '<i/>' : '';
  const u = font.u ? `<u val="${font.u}"/>` : '';
  const strike = font.strike ? '<strike/>' : '';
  const vertAlign = font.vertAlign ? `<vertAlign val="${font.vertAlign}"/>` : '';
  return `<font>${sz}${name}${color}${b}${i}${u}${strike}${vertAlign}</font>`;
}

export function renderBorderRule(border: Border) {
  const inner = ['left', 'right', 'top', 'bottom', 'diagonal']
    .map((v) => {
      if (!border.diagonalUp && !border.diagonalDown && v === 'diagonal') {
        return `<${v}/>`;
      }

      const style = border[v]?.style ?? border.style ?? '';
      if (!style || style === 'none') return `<${v}/>`;
      const styleAttr = style ? ` style="${style}"` : '';

      const color = border[v]?.color ?? border.color ?? '';
      const colorTag = color ? `<color rgb="${(color.length === 6 ? 'FF' : '') + color}" />` : '';
      return `<${v}${styleAttr}>${colorTag}</${v}>`;
    })
    .join('');

  const diagonalUp = border.diagonalUp ? ' diagonalUp="1"' : '';
  const diagonalDown = border.diagonalDown ? ' diagonalDown="1"' : '';
  return `<border${diagonalUp}${diagonalDown}>${inner}</border>`;
}

export function renderFillRule(fill: Fill) {
  const fgColor = fill.fgColor ? `<fgColor rgb="${(fill.fgColor.length === 6 ? 'FF' : '') + fill.fgColor }"/>` : '<fgColor auto="1"/>';
  const bgColor = fill.bgColor ? `<bgColor rgb="${(fill.bgColor.length === 6 ? 'FF' : '') + fill.bgColor }"/>` : '<bgColor auto="1"/>';
  return `<fill><patternFill patternType="${fill.patternType}">${fgColor}${bgColor}</patternFill></fill>`;
}

export function renderAlignmentRule(alignment: Alignment) {
  const horizontal = alignment.horizontal ? ` horizontal="${alignment.horizontal}"` : '';
  const vertical = alignment.vertical ? ` vertical="${alignment.vertical}"` : '';
  const indent =
    ['left', 'right', 'distributed'].includes(alignment.horizontal) && alignment.indent
      ? ` indent="${alignment.indent}"`
      : '';
  const textRotation = alignment.textRotation ? ` textRotation="${alignment.textRotation}"` : '';
  const wrapText = alignment.wrapText ? ' wrapText="1"' : '';
  const shrinkToFit = alignment.shrinkToFit ? ' shrinkToFit="1"' : '';
  return `<alignment${horizontal}${vertical}${indent}${textRotation}${wrapText}${shrinkToFit}/>`;
}

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
