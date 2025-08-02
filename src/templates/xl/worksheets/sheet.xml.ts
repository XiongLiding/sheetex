export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    {{#rows}}
    <row r="{{number}}">
      {{#cells}}
        {{#isString}}
        <c r="{{column}}" t="str" {{#style}}s="{{style}}"{{/style}>
          <t>{{value}}</t>
        </c>
        {{/isString}}
        {{#isNumber}}
        <c r="{{$id}}" {{#style}}s="{{style}}"{{/style}>
          <v>{{value}}</v>
        </c>
        {{/isNumber}}
      {{/cells}}
    </row>
    {{/rows}}
  </sheetData>
</worksheet>
`
