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

/**
 * 对齐
 *
 * @example
 * 居中对齐
 * ```javascript
 * const alignment = {
 *   horizontal: 'center',
 *   vertical: 'center',
 * }
 * ```
 */
export interface Alignment {
  /**
   * 水平对齐
   *
   * @remarks
   * | 可用值 | 含义 |
   * | ---- | ---- |
   * | left | 水平居左 |
   * | center | 水平居中 |
   * | right | 水平居右 |
   * | fill | 填充 |
   * | justified | 两端对齐 |
   * | centerContinuous | 跨列居中 |
   * | distributed | 分散对齐 |
   */
  horizontal?: keyof typeof alignmentHorizontal;

  /**
   * 垂直对齐
   *
   * @remarks
   * | 可用值 | 含义 |
   * | ---- | ---- |
   * | top | 顶部对齐 |
   * | center | 垂直居中 |
   * | bottom | 底部对齐 |
   * | justified | 两端对齐 |
   * | distributed | 分散对齐 |
   */
  vertical?: keyof typeof alignmentVertical;

  /**
   * 缩进，只有当水平对齐为居左、居右和分散对齐时才有效
   */
  indent?: number;

  /**
   * 文字旋转，旋转并不是连续的，0-90，90-180，180-255对应三套不同的规则，并且临界值在不同软件中可能有不同的行为
   *
   * @remarks
   * | 范围 | 含义 |
   * | ---- | ---- |
   * | 0 <= n <= 90 | 逆时针旋转，从为水平到垂直 |
   * | 90 < n <= 180 | 顺时针旋转，从水平到垂直 |
   * | 180 < n <= 255 | 特殊含义 |
   */
  textRotation?: number;

  /**
   * 自动换行
   */
  wrapText?: boolean;

  /**
   * 缩小字体填充
   */
  shrinkToFit?: boolean;
}

/**
 * 字体
 *
 * @example
 * ```javascript
 * const font = {
 *   sz: 18,
 *   name: '微软雅黑',
 *   color: 'FF0000',
 *   b: true,
 * }
 * ```
 */
export interface Font {
  /**
   * 文字大小（字号）
   */
  sz?: number;
  /**
   * 字体名称
   */
  name?: string;
  /**
   * 文字颜色
   */
  color?: string;
  /**
   * 是否粗体
   */
  b?: boolean;
  /**
   * 是否斜体
   */
  i?: boolean;
  /**
   * 下划线样式
   */
  u?: keyof typeof fontUnderline | false;
  /**
   * 是否显示删除线
   */
  strike?: boolean;
  /**
   * 上标或下标
   */
  vertAlign?: keyof typeof fontVertAlign | false;
}

/**
 * 边框样式
 */
export interface BorderStyle {
  /**
   * 线型
   */
  style: keyof typeof borderLineStyle;
  /**
   * 颜色，ARGB 或 RGB 格式，Alpha 通道实际不起作用
   *
   * ```
   * color: 'FF0000';
   * ```
   */
  color: string;
}

/**
 * 边框
 *
 * @example
 * 1. 设置上下左右边框都为红色细线
 * ```javascript
 * const border = {
 *   style: 'thin',
 *   color: 'FF0000',
 * }
 * ```
 * 2. 设置左侧为红色细线，右侧为黑色粗线
 * ```javascript
 * const border = {
 *   left: {
 *     style: 'thin',
 *     color: 'FF0000',
 *   },
 *   right: {
 *     style: 'thick',
 *     color: '000000',
 *   }
 * }
 * ```
 *
 * @remarks
 * 1. 使用全局线型和颜色可以一次性设置所有方向上的边框
 * 2. 具体方向上的边框样式优先于全局样式，可在全局样式的基础上对特定位置进行覆盖
 * 3. 对角线需要同时设置样式和方向才能显示
 */
export interface Border {
  /**
   * 全局线型
   */
  style?: keyof typeof borderLineStyle;
  /**
   * 全局颜色
   */
  color?: string;
  /**
   * 是否显示从左下到右上的对角线
   */
  diagonalUp?: boolean;
  /**
   * 是否显示从左上到右下的对角线
   */
  diagonalDown?: boolean;

  /**
   * 左侧边框样式
   */
  left?: BorderStyle;
  /**
   * 上方边框样式
   */
  top?: BorderStyle;
  /**
   * 右侧边框样式
   */
  right?: BorderStyle;
  /**
   * 底部边框样式
   */
  bottom?: BorderStyle;
  /**
   * 对角线样式
   */
  diagonal?: BorderStyle;
}

/**
 * 填充：包含填充图样、前景色、背景色
 *
 * @example
 * 纯红色背景
 * ```javascript
 * const bgRed = {
 *   patternType: 'solid',
 *   fgColor: 'FF0000',
 * }
 * ```
 * @remarks
 * 使用纯色填充 solid 时，前景色完全覆盖背景色，所以设置纯色背景时，颜色要设置在 fgColor 上而非 bgColor
 */
export interface Fill {
  /**
   * 填充图样
   *
   * @remarks
   * 大多数在线预览功能仅支持 none 和 solid 类型
   */
  patternType: keyof typeof fillPatternType | 'none';
  /**
   * 前景色
   */
  fgColor?: string;
  /**
   * 背景色
   */
  bgColor?: string;
}

/**
 * 样式：由格式码、字体、边框、填充、对齐组合而成
 */
export interface Style {
  /**
   * 格式码：将数字按特定的格式进行展示
   *
   *
   * @remarks
   * 常用格式码
   * | 格式码 | 含义 |
   * | ---: | ---- |
   * | 0 | 取整 |
   * | 0.00 | 保留 2 位小数 |
   * | #,##0 | 带千分位的整数 |
   * | #,##0.00 | 带千分位并保留2位小数 |
   * | 0% | 百分数 |
   * | 0.00% | 百分数保留2位小数 |
   * | 0.00E+00 | 科学计数法 |
   * | $#,##0.00 | 金额 |
   *
   * 完整配置语法请参考：{@link https://support.microsoft.com/zh-cn/office/-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5 | 自定义数字格式的准则 }
   */
  formatCode?: string;
  /**
   * 字体
   */
  font?: Font;
  /**
   * 边框
   */
  border?: Border;
  /**
   * 填充
   */
  fill?: Fill;
  /**
   * 对齐
   */
  alignment?: Alignment;
}

/**
 * 样式表：包含多个样式规则，每个样式规则对应一个样式名称，样式名称作为键，样式规则作为值
 *
 * @example
 * ```javascript
 * const styles = {
 *   default: {
 *     alignment: {
 *       horizontal: 'center',
 *       vertical: 'center',
 *     }
 *   }
 *   title: {
 *     font: {
 *       sz: 18,
 *       b: true,
 *     }
 *     alignment: {
 *       horizontal: 'center',
 *       vertical: 'center',
 *     }
 *   }
 * }
 * ```
 *
 * @remarks
 * 样式名称 default 为系统保留名称，所有未明确设置 style 的单元格将套用此样式
 */
export interface Styles extends Record<string, Style> {}

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
 * @return - 样式规则
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
