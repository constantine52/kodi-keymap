<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" version="1.0" exclude-result-prefixes="office table text">
  <!--  -->
  <xsl:output method="xml" indent="yes" encoding="UTF-8" omit-xml-declaration="no"/>

<!--global variables-->
<xsl:variable name="sheetName" select="'keymap'" />
<xsl:variable name="lastSuplimentaryRow" select="2" />


<xsl:template match="/"><!-- Match the document base - start the process -->
  
  <xsl:for-each select="//table:table">
    <xsl:if test="./@table:name = $sheetName">
      <xsl:apply-templates select="."/><!--match first table of the book-->
    </xsl:if>
  </xsl:for-each>

</xsl:template>


<xsl:template match="table:table">

  <xsl:variable name="rawTable">

    <xsl:for-each select="table:table-row">

      <xsl:element name="row">

        <xsl:call-template name="CellLoop">
          <xsl:with-param name="totalCells" select="count(./table:table-cell)"/>
          <xsl:with-param name="currentColumn" select="1"/>
          <xsl:with-param name="currentCell" select="1"/>
          <xsl:with-param name="currentIteration" select="1"/>
        </xsl:call-template>

      </xsl:element><!-- row -->

    </xsl:for-each><!-- table:table-row -->

  </xsl:variable><!-- rawTable -->

  <xsl:variable name="expandedTable" select="exsl:node-set($rawTable)" xmlns:exsl="http://exslt.org/common"/>

  <xsl:element name="{string(./@table:name)}"><!-- sheet name -->

    <xsl:call-template name="rawTableParser">
      <xsl:with-param name="expandedTable" select="$expandedTable"/>
    </xsl:call-template>

  </xsl:element><!-- keymap -->

</xsl:template>


<xsl:template name="rawTableParser">
  <xsl:param name="expandedTable"/>

    <xsl:for-each select="$expandedTable/row[1]/column"><!-- global -->

      <xsl:variable name="functionColumn">
        <xsl:value-of select="position()"/>
      </xsl:variable>

      <xsl:if test=". != ''">

        <xsl:variable name="key_functionSet">

          <xsl:call-template name="columnsParser">
            <xsl:with-param name="expandedTable" select="$expandedTable"/>
            <xsl:with-param name="functionColumn" select="$functionColumn"/>
          </xsl:call-template>

        </xsl:variable><!--{string(.)}-->

        <xsl:if test="$key_functionSet != ''">
          <xsl:element name="{string(.)}">
            <xsl:copy-of select="exsl:node-set($key_functionSet)" xmlns:exsl="http://exslt.org/common" />
          </xsl:element>
        </xsl:if>

      </xsl:if><!--. != ''-->

    </xsl:for-each><!-- global -->

</xsl:template>


<xsl:template name="columnsParser">
  <xsl:param name="expandedTable" select="$expandedTable"/>
  <xsl:param name="functionColumn" select="$functionColumn"/>

  <xsl:variable name="keyColumn">

    <xsl:for-each select="$expandedTable/row[2]/column"> <!-- key column navigation -->

      <xsl:variable name="currentPosition" select="position()" />

      <xsl:if test="$expandedTable/row[1]/column[position() = $currentPosition] = '' and . = $expandedTable/row[2]/column[position() = $functionColumn]">
        <xsl:value-of select="$currentPosition" />
      </xsl:if>

    </xsl:for-each><!-- key column navigation -->

  </xsl:variable>

  <xsl:if test="$keyColumn &gt; 0">

    <xsl:element name="{string($expandedTable/row[2]/column[position() = $functionColumn])}"><!-- keyboard -->

      <xsl:call-template name="cellParser">
        <xsl:with-param name="expandedTable" select="$expandedTable"/>
        <xsl:with-param name="functionColumn" select="$functionColumn"/>
        <xsl:with-param name="keyColumn" select="$keyColumn"/>
      </xsl:call-template>

    </xsl:element><!-- keyboard -->

  </xsl:if>

</xsl:template>

<xsl:template name="cellParser">
  <xsl:param name="expandedTable" select="$expandedTable"/>
  <xsl:param name="functionColumn" select="$functionColumn"/>
  <xsl:param name="keyColumn" select="$keyColumn"/>

  <xsl:for-each select="$expandedTable/row"><!-- row -->

    <xsl:if test="position() &gt; $lastSuplimentaryRow and ./column[position() = $functionColumn] != '' and ./column[position() = $keyColumn] != ''">

      <xsl:element name="{string(./column[position() = $keyColumn])}">

        <xsl:if test="./column[position() = $keyColumn + 1] != ''">
          <xsl:attribute name="mod">
            <xsl:value-of select="string(./column[position() = $keyColumn + 1])"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:value-of select="string(./column[position() = $functionColumn])"/>

      </xsl:element>

    </xsl:if>

  </xsl:for-each><!-- row -->

</xsl:template><!-- table:table -->


  <xsl:template name="CellLoop">
    <xsl:param name="totalCells"/>
    <xsl:param name="currentColumn"/>
    <xsl:param name="currentCell"/>
    <xsl:param name="currentIteration"/>

    <xsl:if test="$currentCell &lt; $totalCells or ($currentCell = $totalCells and ./table:table-cell[position() = $currentCell] != '')">
      <!-- Output the cell contents -->
      <xsl:call-template name="MakeTag">

        <xsl:with-param name="tagContentPointer" select="./table:table-cell[position() = $currentCell]"/>
      </xsl:call-template>
      <!-- Decide how to recusively call this template -->
      <xsl:choose>
        <xsl:when test="$currentIteration &lt; ./table:table-cell[position() = $currentCell]/@table:number-columns-repeated">
          <!-- On a 'repeating' cell - call the template again, incrementing the column count, but keep the cell count the same -->
          <xsl:call-template name="CellLoop">
            <xsl:with-param name="totalCells" select="$totalCells"/>
            <xsl:with-param name="currentColumn" select="$currentColumn + 1"/>
            <xsl:with-param name="currentCell" select="$currentCell"/>
            <xsl:with-param name="currentIteration" select="$currentIteration + 1"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <!-- On a 'normal' cell - call the template again, incrementing both the column count and the cell count -->
          <xsl:call-template name="CellLoop">
            <xsl:with-param name="totalCells" select="$totalCells"/>
            <xsl:with-param name="currentColumn" select="$currentColumn + 1"/>
            <xsl:with-param name="currentCell" select="$currentCell + 1"/>
            <xsl:with-param name="currentIteration" select="1"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <xsl:template name="MakeTag">
<!--
    <xsl:param name="currentColumn"/>
-->
    <xsl:param name="tagContentPointer"/>

      <xsl:element name="column">
        <xsl:value-of select="$tagContentPointer/text:p"/>
      </xsl:element>

  </xsl:template>

</xsl:stylesheet>
