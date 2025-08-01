export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Excel</Application>
  <headings>
    <vt:vector xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" size="2" baseType="variant">
      <vt:variant>
        <vt:lpstr>Worksheets</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>{{sheetsCount}}</vt:i4>
      </vt:variant>
    </vt:vector>
  </headings>
  <titlesOfParts>
    <vt:vector xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" size="{{sheetsCount}}" baseType="lpstr">
      {{#sheets}}
      <vt:lpstr>Sheet{{.}}</vt:lpstr>
      {{/sheets}}
    </vt:vector>
  </titlesOfParts>
</Properties>
`
