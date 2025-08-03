export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {{#numFmtsCount}}
  <numFmts count="{{numFmtsCount}}">
    {{#numFmts}}{{& .}}{{/numFmts}}
  </numFmts> 
  {{/numFmtsCount}}
  
  <fonts count="{{fontsCount}}">
    <font><sz val="12" /><color theme="1" /><name val="Arial" /><family val="2" /><scheme val="minor" /></font>
    {{#fonts}}{{& .}}{{/fonts}}
  </fonts>
  
  <fills count="{{fillsCount}}">
    <fill><patternFill patternType="none" /></fill>
    <fill><patternFill patternType="gray125" /></fill>
    {{#fills}}{{& .}}{{/fills}}
  </fills>
  
  <borders count="{{bordersCount}}">
   <border><left/><right/><top/><bottom/><diagonal/></border>
    {{#borders}}{{& .}}{{/borders}}
  </borders>
  
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  
  <cellXfs count="{{cellXfsCount}}">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    {{#cellXfs}}{{& .}}{{/cellXfs}}
  </cellXfs>
  
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
