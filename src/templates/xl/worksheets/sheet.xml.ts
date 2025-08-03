export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15" />
  {{#colsCount}}
  <cols>
    {{#cols}}<col min="{{min}}" max="{{max}}" width="{{size}}" customWidth="1"/>{{/cols}}
  </cols>
  {{/colsCount}}
  
  <sheetData>
    {{#rows}}
    <row r="{{number}}"{{#ht}}{{& .}}{{/ht}}>
      {{#cells}}
        <c r="{{position}}"{{#isString}} t="str"{{/isString}}{{#style}} s="{{.}}"{{/style}}>
          <v>{{value}}</v>
        </c>
      {{/cells}}
    </row>
    {{/rows}}
  </sheetData>
  
  {{#mergeCellsCount}}
  <mergeCells count="{{mergeCellsCount}}">
    {{#mergeCells}}<mergeCell ref="{{.}}"/>{{/mergeCells}}
  </mergeCells>
  {{/mergeCellsCount}}
</worksheet>
`;
