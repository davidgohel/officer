<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

  <xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />
  <xsl:param name="newdoc" select="document($filename)" />

  <xsl:variable name="updateItems" select="$newdoc" />

  <xsl:template match="@* | node()">
    <xsl:copy>
      <xsl:apply-templates select="@* | node()"/>
    </xsl:copy>
  </xsl:template>


  <xsl:template match="*" mode="copy-no-namespaces">
      <xsl:element name="{local-name()}">
          <xsl:copy-of select="@*"/>
          <xsl:apply-templates select="node()" mode="copy-no-namespaces"/>
      </xsl:element>
  </xsl:template>

  <xsl:template match="comment()| processing-instruction()" mode="copy-no-namespaces">
      <xsl:copy/>
  </xsl:template>

  <xsl:template match="node()[local-name()='sheetData']">
      <sheetData>
        <xsl:apply-templates select="@* | node()" mode="copy-no-namespaces"/>
        <xsl:apply-templates select="$newdoc/sheetData/row" mode="copy-no-namespaces"/>
      </sheetData>
  </xsl:template>

</xsl:stylesheet>
