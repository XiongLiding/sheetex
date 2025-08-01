import Mustache from 'mustache'
import templateString from '../templates/xl/styles.xml.ts'

export interface Font {
  fontSize: number;
  fontFamily: string;
}

export interface Fill {
  color: string;
}

export interface Border {
  left: string,
  right: string,
  top: string,
  bottom: string,
  diagonal: string,
}

export interface CellStyleXfs {
  numFmtId: number;
  numFontId: number;
  fillId: number;
  borderId: number;
}

export interface CellXfs {
  numFmtId: number;
  numFontId: number;
  fillId: number;
  borderId: number;
  xfId: number;
}

export default (fonts: Font[], fills: Fill[], borders: Border[], cellStyleXfs: CellStyleXfs[], cellXfs: CellXfs[]) => {
  return Mustache.render(templateString, {
    fonts, fills, borders, cellStyleXfs, cellXfs,
  })
};
