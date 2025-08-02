export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    {{#rows}}
    <row r="{{number}}">
      {{#cells}}
        <c r="{{position}}"{{#isString}} t="str"{{/isString}}{{#style}} s="{{.}}"{{/style}}>
          <v>{{value}}</v>
        </c>
      {{/cells}}
    </row>
    {{/rows}}
  </sheetData>
</worksheet>
`
