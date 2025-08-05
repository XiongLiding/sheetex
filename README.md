# sheetex

<p align="center">
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License">
  <img src="https://img.shields.io/bundlephobia/min/sheetex" alt="Size">
  <img src="https://img.shields.io/npm/v/sheetex" alt="Version">
</p>

12KB 轻量 Excel 导出库，完整样式支持，开箱即用，浏览器 & Node.js

## 特性

- 🚀 高效生成 `.xlsx` 格式的 Excel 文件
- 📦 12KB
  极致轻量（gzip），从零开始编写，仅依赖 [fflate](https://github.com/101arrowz/fflate) & [mustache](https://github.com/janl/mustache)
- 💻 Office Open XML 标准子集，具备良好兼容性，可在 Microsoft Excel、WPS Office、Numbers、macOS/钉钉/微信预览中正确展示
- 🎨 单元格格式面板完整功能支持：
    - 格式（数字的各种格式化）
    - 对齐（水平、缩进、垂直、方向、自动换行、缩小字体填充）
    - 字体（字体、字形（加粗、斜体）、字号、颜色、下划线、删除线、上标、下标）
    - 边框（上下左右对角线、14种边框样式）
    - 背景色（18种填充样式）
    - 全局设置（单元格合并、行高、列宽）
- 🌐 同时支持 Node.js 和浏览器环境
- 😀 简洁明了的接口，易于理解和使用
- 📄 Learn by example，充分的示例，足以替代文档

## 效果展示

一个示例文件在 Microsoft Excel 中的展示效果

![showtime](./docs/image/showtime.jpg)

生成该文件的代码：[examples/showtime.ts](examples/showtime.ts)

## 适用场景与限制

本库的设计遵循以下原则，旨在为 Web 后台的 Excel 导出场景提供**简单、可靠、轻量**的解决方案，由于本库非常轻量，所以也非常适合移动端场景。

### 1. 聚焦典型场景

本库主要面向网页应用后台中常见的“查询结果导出 Excel”场景，所有设计均围绕该核心用例展开，力求在常见业务中提供简洁高效的解决方案。

### 2. 专注导出，不支持解析

本库专注于 Excel 文件的**生成与导出**，**不提供解析能力**。因此，以下开发模式**不适用**：

> ❌ 读取现有模板文件，并对特定单元格或区域进行修改。  
> （如需此类功能，建议使用 [SheetJS](https://sheetjs.com) 等支持读写操作的库。）

### 3. 支持常用样式，舍弃复杂格式

本库覆盖了 Excel 单元格格式面板中的所有样式（如字体、对齐、边框、填充、数字格式等），但**不支持高级格式特性**，例如：

- 在同一个单元格内对不同文字设置不同样式（富文本）
- 图表、条件格式、公式等高级功能

### 4. 尊重 Excel 原有命名与结构

在 API 设计上，尽可能复用 Excel 原生的术语和结构，降低用户学习成本，提升配置直观性。

### 5. 单向数据流，配置即不可变

为简化使用逻辑并避免运行时状态管理的复杂性，本库采用单向数据流设计：

- 所有配置必须在对象初始化时完成；
- 实例化后，对象状态不可修改；
- 确保输出可预测，易于调试和测试。

### 6. 轻量实现：模板渲染 + 流式压缩

为减小体积并降低系统复杂度，本库采用以下技术方案：

- 使用 `Mustache` 模板引擎渲染 Excel 内部 XML 文件；
- 使用 `fflate` 进行高效压缩，生成最终的 .xlsx 文件。

### 📦 优势：

- 无需依赖大型办公文档处理库，如 jszip + sheetjs（gzip 后 312KB）
- 包体积小，适合前端直出 Excel 文件
- 自带样式功能，sheetjs 需购买商业版或使用过时的样式插件

### 🤔 适合谁？

- ✅ 需要将表格数据导出为 Excel 的后台管理系统
- ✅ 移动端 Excel 文件生成
- ✅ 希望快速集成、不依赖 Node.js 环境的前端项目（当然，也可以在 Node.js 中使用）
- ✅ 不需要复杂格式或模板修改的轻量级场景

### 🚫 不适合谁？

- ❌ 需要读取或修改现有 Excel 文件（建议使用 SheetJS）
- ❌ 需要单元格内富文本、图表、公式等高级功能
- ❌ 需要支持 .xls 格式（本库仅输出 .xlsx）

## 安装

使用 npm 安装：

```bash
npm install sheetex
```

使用 yarn 安装：

```bash
yarn add sheetex
```

使用 pnpm 安装：

```bash
pnpm add sheetex
```

## 快速开始

### 基础用法

生成工作表和工作簿（在浏览器和Node.js中都一样）：

```javascript
import { WorkSheet, WorkBook } from 'sheetex';

// 生成一个工作表，里面有一个数据块，数据块里有一行数据，第一个列是"hello"，第二列是"world"，该数据块位于A1
const ws = new WorkSheet('sheet1', [{ data: [['hello', 'world']], origin: 'A1' }]);

// 生成一个工作簿，里面有一个工作表
const wb = new WorkBook([ws]);
```

在浏览器中保存

```javascript
wb.downloadAs('example.xlsx');
```

在服务器中保存

```javascript
// 获取文件的 buffer 形式
const buffer = await wb.getZipBuffer();

// 保存到服务器文件系统
fs.writeFileSync('example.xlsx', buffer);
```

在服务器上生成文件并返回给客户端
```javascript
// 获取文件的 buffer 形式
const buffer = await wb.getZipBuffer();

// 返回给客户端（以 express 为例）
res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
res.send(buffer);
```

## Learn By Example

您可以通过例子学习本库的使用方法：

**强烈建议在查看示例代码前，先打开同名的 xlsx 文件查看最终效果，可以很好的加快理解。文件位于 src/examples/xlsx/ 目录。**

1. [simple.ts](examples/simple.ts) 无样式快速导出 Excel ，适合作为 csv 替代
2. [format.ts](examples/format.ts) 数字的格式化（取整、保留小数、科学计数法、货币格式等）
3. [font.ts](examples/font.ts) 字体相关设置（字体、大小、颜色、加粗、斜体、下划线、删除线、上下标） 
4. [alignment.ts](examples/alignment.ts) 对齐（水平对齐、垂直对齐、缩进、文本旋转、自动换行、缩小字体填充）
5. [border.ts](examples/border.ts) 边框线型和颜色的设置，以及多边框组合使用的方法
6. [fill.ts](examples/fill.ts) 背景色填充，一般而言仅需了解其中的 solid 样式即可
7. [option.ts](examples/option.ts) 行高、列宽、单元格合并 
8. [showtime.ts](examples/showtime.ts) 根据一个真实项目中的案例编写，展示了对上面各类样式的综合应用

## API

本章节按照从顶层到细节的顺序对用户会接触的类型和方法进行说明。读者也可以顺着这里的顺序去阅读源代码了解内部的实现逻辑。

### 工作簿 WorkBook

一个 WorkBook 相当于一个 Excel 文件，其中可以包含多个工作表，主要职责是对个工作表进行组织，实现文件的生成与保存

```javascript
/**
 * 创建一个工作簿
 * @param workSheets 工作表
 */
new WorkBook(workSheets)

/**
 * 获得当前WorkBook对应的文件 buffer
 * @returns Promise<Buffer> 生成文件的 buffer
 */
WorkBook.getZipBuffer()

/**
 * 在浏览器中下载当前WorkBook对应的文件
 * @param filename 文件名
 */
WorkBook.downloadAs(filename)
```

### 工作表 WorkSheet

相当于 Excel 文件中一个工作表，每个工作表可以有自己的名称，包含一个或多个数据块，拥有一个样式表，以及一个配置项（包含行高、列宽、单元格合并）

```javascript
/**
 * @param name string 工作表名称
 * @param blocks DataBlock[] 数据块
 * @param styles Styles 样式表
 * @param options CellOptions 工作表选项
 */
new WorkSheet(name, data, style, options)
```

### 数据块 DataBlock

相当于工作表中的一块区域，比如 A1:C5（实际不一定是严谨的矩形），一个工作表可以有多个数据块，多个数据块区域若有重复，则后者的数据会覆盖前者，每个工作表都可以定义起始坐标（数据库中第一个单元格的坐标）。

```typescript
type DataBlock = {
  origin: string; // 数据块起始坐标
  data: DataRow[]; // 数据行
}
```

### 数据行 DataRow

一个数据块可以包含多行数据，每个数据行可以包含多个数据格

```typescript
type DataRow = DataCell[]; // 数据行是数据格的数组
```

### 数据格 DataCell

数据格可以是数字、文本或对象形式，当数据格为对象时，可以额外包含该格的样式信息

```typescript
type DataCell = string | number | {
  value: string | number; style: Style;
}
```

### 单元格格式 Style

类似前端样式中的 class ，针对每个单元格格式，可以定义样式信息，比如字体、颜色、背景色、边框、对齐、格式等，在数据格中可以引用样式信息

```typescript
interface Style {
  formatCode?: string; // 格式化代码
  font?: Font; // 字体
  border?: Border; // 边框
  fill?: Fill; // 填充（背景）
  alignment?: Alignment; // 对齐
}
```

### 格式化代码

可以定义单元格数据的显示格式，比如数字的千分位分隔符、日期的格式等。

```typescript
type FormatCode = string;
```

具体使用示例请查看 [examples/format.ts](examples/format.ts)

### 字体 Font

字体信息，包含字体名称、大小、颜色、加粗、斜体、删除线、下划线、上标、下标等

```typescript
interface Font {
  sz?: number; // 字体大小
  name?: string; // 字体名称
  color?: string; // 字体颜色
  b?: boolean; // 是否加粗
  i?: boolean; // 是否斜体
  u?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | false; // 下划线样式
  strike?: boolean; // 是否删除
  vertAlign?: 'superscript' | 'subscript' | false; // 上标、下标
}
```

具体使用示例请查看 [examples/font.ts](examples/font.ts)

### 边框 Border

边框信息，包含上边框、下边框、左边框、右边框、对角线、对角线方向、颜色、线型、线宽等

```typescript
interface Border {
  style?: // 14 种边框样式
    | 'none' | 'dashDot' | 'dashDotDot' | 'dashed' | 'dotted' | 'double' | 'hair' | 'medium' | 'mediumDashDot' | 'mediumDashDotDot' | 'mediumDashed' | 'slantDashDot' | 'thick' | 'thin';
  color?: string; // 边框颜色
  diagonalUp?: boolean; // 上对角线
  diagonalDown?: boolean; // 下对角线

  left?: BorderStyle; // 左边框样式
  top?: BorderStyle; // 上边框样式
  right?: BorderStyle; // 右边框样式
  bottom?: BorderStyle; // 下边框样式
  diagonal?: BorderStyle; // 对角线样式
}
```

BorderStyle 包含 style 与 color 两个属性；

设置 Border 上的 style 与 color 可以代替分别设置 left/top/right/bottom/diagonal

具体使用示例请查看 [examples/border.ts](examples/border.ts)

### 填充 Fill

填充信息，包含填充颜色与填充模式

```typescript
export interface Fill {
  patternType: // 18 种填充模式，只有 Microsoft Excel 和 WPS Office 兼容所有模式
    | 'none' | 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'gray125' | 'gray0625' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis';
  fgColor?: string;
  bgColor?: string;
}
```

注意纯色背景的 patternType 为 solid，填充颜色为 fgColor

具体使用示例请查看 [examples/fill.ts](examples/fill.ts)

### 对齐 Alignment

对齐信息，包含水平对齐、文本缩进、垂直对齐、文本旋转、自动换行、缩小字体填充等

```typescript
interface Alignment {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'; // 水平对齐
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'; // 垂直对齐
  indent?: number; // 文本缩进
  textRotation?: number; // 文本旋转
  wrapText?: boolean; // 自动换行
  shrinkToFit?: boolean; // 缩小字体填充
}
```

具体使用示例请查看 [examples/alignment.ts](examples/alignment.ts)

### SheetOptions

```typescript
export interface SheetOptions {
  mergeCells?: string[]; // 单元格合并
  colWidths?: Size[]; // 列宽
  rowHeights?: Size[]; // 行高
}
```

具体使用示例请查看 [examples/option.ts](examples/option.ts)

### 合并单元格 MergeCells

可以设置多组要合并的单元格范围，比如 ['A1:C5', 'D1:E2', 'F1:G3']

### 列宽 ColWidths

可以设置多列的列宽，比如 [{min: 1, max: 3, size: 10}, {min: 4, size: 20}]

min 是起始列，max 是结束列，size 是列宽，min 必填，max 可选：

- 当 max 填写且 size 为数字时，表示从 min 到 max 的列宽都是 size；
- 当 max 填写且 size 为数组时，表示从 min 到 max 的列宽依次如 size 数组中所述，数组长度不够时，从头开始循环；
- 当 max 不填且 size 为数字时，表示 min 列的列宽是 size；
- 当 max 不填且 size 为数组时，表示从 min 开始，列宽依次如 size 数组中所述，直到数组耗尽。

### 行高 RowHeights

行高的定义方式与列宽相同，请注意，Excel 中行高和列宽的单位是不同的，单元格对应的行高和列宽为相同的数字并不会显示成正方形

## 贡献

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目。

## 许可证

本项目采用 MIT 许可证。详情请见 [LICENSE](LICENSE) 文件。
