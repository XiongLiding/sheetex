export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="{{numFmtsCount}}">
    {{#numFmts}}{{.}}{{/numFmts}}
  </numFmts> 
  <fonts count="{{fontsCount}}">
    {{#fonts}}{{.}}{{/fonts}}
  </fonts>
  <fills count="{{fillsCount}}">
    {{#fills}}{{.}}{{/fills}}
  </fills>
  <borders count="1">
    {{#borders}}{{.}}{{/borders}}
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    {{#cellXfs}}{{.}}{{/cellXfs}}
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
