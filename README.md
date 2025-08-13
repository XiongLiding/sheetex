# sheetex

<p align="center">
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License">
  <img src="https://img.shields.io/npm/v/sheetex" alt="Version">
  <a href="https://gitee.com/xiongliding/sheetex/stargazers"><img alt="GitHub Repo stars" src="https://img.shields.io/github/stars/xiongliding/sheetex?style=flat&logo=github"/></a>
  <a href="https://gitee.com/xiongliding/sheetex/stargazers"><img src="https://gitee.com/xiongliding/sheetex/badge/star.svg?theme=dark" alt="star"/></a>
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
wb.saveAs('example.xlsx');
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

### 快速使用
如果标题居中对齐、表头文字加粗、给所有数据套上边框已经是你全部的需求。

尽管用 [just](examples/just.ts) 开头的三个函数实现你的需求吧！

等到对细节有更多要求时，再来查看深入学习部分。

### 深入学习 
1. [simple.ts](examples/simple.ts) 无样式快速导出 Excel
2. [format.ts](examples/format.ts) 数字的格式化（取整、保留小数、科学计数法、货币格式等）
3. [font.ts](examples/font.ts) 字体相关设置（字体、大小、颜色、加粗、斜体、下划线、删除线、上下标） 
4. [alignment.ts](examples/alignment.ts) 对齐（水平对齐、垂直对齐、缩进、文本旋转、自动换行、缩小字体填充）
5. [border.ts](examples/border.ts) 边框线型和颜色的设置，以及多边框组合使用的方法
6. [fill.ts](examples/fill.ts) 背景色填充，一般而言仅需了解其中的 solid 样式即可
7. [option.ts](examples/option.ts) 行高、列宽、单元格合并 
8. [showtime.ts](examples/showtime.ts) 根据一个真实项目中的案例编写，展示了对上面各类样式的综合应用

## API

完整的 [API文档](https://xiongliding.github.io/sheetex/api/)，包含所有对用户暴露的对象、方法和 TypeScript 类型

## 贡献

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目。

## 许可证

本项目采用 MIT 许可证。详情请见 [LICENSE](LICENSE) 文件。
