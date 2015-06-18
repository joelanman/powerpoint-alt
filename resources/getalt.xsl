<?xml version="1.0"?>
<xsl:stylesheet 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:z="http://schemas.openxmlformats.org/package/2006/relationships"
  version="2.0">

  <xsl:output method="text"/>

  <xsl:param name="pptfile"/>
  <xsl:param name="slidefile"/>

  <xsl:variable name="quot">
    <xsl:text>"</xsl:text>
  </xsl:variable>

  <xsl:variable name="apos">
    <xsl:text>'</xsl:text>
  </xsl:variable>

  <xsl:variable name="slideno" 
    select="substring-before(
            substring-after($slidefile,'slides/slide'),'.xml')"/>

  <xsl:variable name="relsfile" 
    select="concat('ppt/slides/_rels/slide',$slideno,'.xml.rels')"/>
  
  <xsl:template match="/">
    <xsl:apply-templates select="descendant::p:pic |
                                 descendant::p:sp"/>
  </xsl:template>

  <!-- Images -->

  <xsl:template match="p:pic">
    <xsl:value-of select="$pptfile"/>
    <xsl:text>,</xsl:text>
    <xsl:value-of select="$slideno"/>
    <xsl:text>,</xsl:text>
    <xsl:variable name="id" select="descendant::a:blip/@r:embed"/>
    <!-- image filename, extracted from the relationships file -->
    <xsl:value-of select="substring-after(document($relsfile)/z:Relationships/
                          z:Relationship[@Id=$id]/@Target,'../media/')"/>
    <xsl:text>,</xsl:text>
    <xsl:value-of select="p:nvPicPr/p:cNvPr/@id"/>
    <xsl:text>,"</xsl:text>
    <xsl:value-of select="p:nvPicPr/p:cNvPr/@name"/>
    <xsl:text>","</xsl:text>
    <xsl:value-of select="translate(p:nvPicPr/p:cNvPr/@descr,$quot,$apos)"/>
    <xsl:text>","</xsl:text>
    <xsl:for-each select="p:txBody/a:p/*">
      <xsl:apply-templates/>
    </xsl:for-each>
    <xsl:text>"&#xa;</xsl:text>
  </xsl:template>

  <!-- Slide Paragraphs (text or drawing) -->

  <xsl:template match="p:sp">
    <xsl:value-of select="$pptfile"/>
    <xsl:text>,</xsl:text>
    <xsl:value-of select="$slideno"/>
    <xsl:text>,</xsl:text>
    <xsl:value-of select="count(preceding-sibling::p:sp)+1"/>
    <xsl:text>,</xsl:text>
    <xsl:value-of select="p:nvSpPr/p:cNvPr/@id"/>
    <xsl:text>,"</xsl:text>
    <xsl:value-of select="p:nvSpPr/p:cNvPr/@name"/>
    <xsl:text>","</xsl:text>
    <xsl:value-of select="translate(p:nvSpPr/p:cNvPr/@descr,$quot,$apos)"/>
    <xsl:text>","</xsl:text>
    <xsl:for-each select="p:txBody/a:p/*">
      <xsl:apply-templates/>
    </xsl:for-each>
    <xsl:text>"</xsl:text>
    <xsl:if test="ancestor::p:grpSp">
      <xsl:text>,</xsl:text>
      <xsl:value-of select="ancestor::p:grpSp/p:nvGrpSpPr/p:cNvPr/@id"/>
      <xsl:text>,"</xsl:text>
      <xsl:value-of select="ancestor::p:grpSp/p:nvGrpSpPr/p:cNvPr/@name"/>
      <xsl:text>","</xsl:text>
      <xsl:value-of select="translate(ancestor::p:grpSp/p:nvGrpSpPr/p:cNvPr/@descr,$quot,$apos)"/>
      <xsl:text>"</xsl:text>
    </xsl:if>
    <xsl:text>&#xa;</xsl:text>
 </xsl:template>

 <xsl:template match="text()">
   <xsl:value-of select="translate(.,$quot,$apos)"/>
 </xsl:template>

 <xsl:template match="a:br">
   <xsl:text>&#x23ce;</xsl:text>
 </xsl:template>

</xsl:stylesheet>

<!--
Slide number
- Slide title (if there is one, blank if not)
- Image number
- Image text 
(repeated for all the images)
- Shape number
- Shape text
(repeated for all the shapes)
- Any other text on the slide
- Slidenumber placeholder (if it appears on the slide)
Any notes text associated with the slide
-->
