export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    {{#fonts}}
    <font><sz val="{{fontSize}}"/><name val="{{fontFamily}}"/><family val="2"/></font>
    {{/fonts}}
  </fonts>
  <fills count="1">
    {{#fills}}
    <fill><patternFill patternType="none"/></fill>
    {{/fills}}
  </fills>
  <borders count="1">
    {{#borders}}
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    {{/borders}}
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
`
