import Mustache from 'mustache';
import templateString from '../templates/xl/styles.xml.ts';

export interface Alignment {
  horizontal: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  indent: number;
  textRotation: number;
  wrapText: boolean;
  shrinkToFit: boolean;
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
  diagonalUp: boolean;
  diagonalDown: boolean;

  left: BorderStyle;
  top: BorderStyle;
  right: BorderStyle;
  bottom: BorderStyle;
  diagonal: BorderStyle;
}

export interface Fill {
  patternType:
    | 'none'
    | 'solid'
    | 'gradient'
    | 'path'
    | 'darkDown'
    | 'darkGray'
    | 'darkGrid'
    | 'darkHorizontal'
    | 'darkTrellis'
    | 'darkUp'
    | 'darkVertical'
    | 'gray0625'
    | 'gray125'
    | 'lightDown'
    | 'lightGray'
    | 'lightGrid'
    | 'lightHorizontal'
    | 'lightTrellis'
    | 'lightUp'
    | 'lightVertical'
    | 'mediumGray'
    | 'pct10'
    | 'pct12';
  fgColor: string;
  bgColor: string;
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
  const color = font.color ? `<color rgb="${font.color}"/>` : '';
  const b = font.b ? '<b/>' : '';
  const i = font.i ? '<i/>' : '';
  const u = font.u ? `<u val="${font.u}"/>` : '';
  const strike = font.strike ? '<strike/>' : '';
  const vertAlign = font.vertAlign ? `<vertAlign val="${font.vertAlign}"/>` : '';
  return `<font>${sz}${name}${color}${b}${i}${u}${strike}${vertAlign}</font>`;
}

export function renderBorderRule(border: Border) {
  if (border.style && border.color.length === 8 && border.style !== 'none') {
    border.left = border.left ?? { style: border.style, color: border.color };
    border.top = border.top ?? { style: border.style, color: border.color };
    border.right = border.right ?? { style: border.style, color: border.color };
    border.bottom = border.bottom ?? { style: border.style, color: border.color };
    border.diagonal = border.diagonal ?? { style: border.style, color: border.color };
  }
  const left = border.left ? `<left style="${border.left.style}" rgb="${border.left.color}"/>` : '';
  const top = border.top ? `<top style="${border.top.style}" rgb="${border.top.color}"/>` : '';
  const right = border.right ? `<right style="${border.right.style}" rgb="${border.right.color}"/>` : '';
  const bottom = border.bottom ? `<bottom style="${border.bottom.style}" rgb="${border.bottom.color}"/>` : '';
  const diagonal = border.diagonal ? `<diagonal style="${border.diagonal.style}" rgb="${border.diagonal.color}"/>` : '';
  const diagonalUp = border.diagonalUp ? ' diagonalUp="1"' : '';
  const diagonalDown = border.diagonalUp ? ' diagonalDown="1"' : '';
  return `<border${diagonalUp}${diagonalDown}>${left}${top}${right}${bottom}${diagonal}</border>`;
}

export function renderFillRule(fill: Fill) {
  const fgColor = fill.fgColor ? `<fgColor rgb="${fill.fgColor}"/>` : '';
  const bgColor = fill.bgColor ? `<bgColor rgb="${fill.bgColor}"/>` : '';
  return `<fill><patternFill patternType="${fill.patternType}">${fgColor}${bgColor}</patternFill></fill>`;
}

export function renderAlignmentRule(alignment: Alignment) {
  const horizontal = alignment.horizontal ? `<horizontal val="${alignment.horizontal}"/>` : '';
  const vertical = alignment.vertical ? `<vertical val="${alignment.vertical}"/>` : '';
  const indent = alignment.indent ? `<indent val="${alignment.indent}"/>` : '';
  const textRotation = alignment.textRotation ? `<textRotation val="${alignment.textRotation}"/>` : '';
  const wrapText = alignment.wrapText ? '<wrapText/>' : '';
  const shrinkToFit = alignment.shrinkToFit ? '<shrinkToFit/>' : '';
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
    fillsCount: fills.length + 1,
    borders,
    bordersCount: borders.length + 1,
    cellXfs,
    cellXfsCount: cellXfs.length + 1,
  });
};
