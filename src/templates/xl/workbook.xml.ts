export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    {{#sheets}}
    <sheet name="Sheet{{.}}" sheetId="{{.}}" r:id="rId{{.}}"/>
    {{/sheets}}
  </sheets>
</workbook>
`
