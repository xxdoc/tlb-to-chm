<?xml version="1.0" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output omit-xml-declaration="yes" method="xml" />
	<xsl:include href="common.xsl" />

	<xsl:template match="/">
		<HTML>
			<BODY>
				<UL>
					<xsl:apply-templates select="//TypeLibInfo|//CoclassInfo|//RecordInfo|//InterfaceInfo|//DeclarationInfo|//ConstantInfo" />
				</UL>
			</BODY>
		</HTML>
	</xsl:template>

	<xsl:template match="TypeLibInfo">
		<xsl:call-template name="liobjlink">
			<xsl:with-param name="elem" select="current()" />
		</xsl:call-template>
	</xsl:template>

	<xsl:template match="CoclassInfo">
		<xsl:if test="key('filtercoclasses',@guid)">
			<xsl:call-template name="liobjlink">
				<xsl:with-param name="elem" select="current()" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template match="InterfaceInfo">
		<xsl:if test="key('filterinterfaces',@guid)">
			<xsl:call-template name="liobjlink">
				<xsl:with-param name="elem" select="current()" />
			</xsl:call-template>
		</xsl:if>
		<xsl:if test="key('filterinterfaces',@guid)|key('class',@guid)">
			<xsl:for-each select="key('filterifmembers',@guid)">
				<xsl:call-template name="liobjlink">
					<xsl:with-param name="elem" select="current()" />
				</xsl:call-template>
			</xsl:for-each>
		</xsl:if>
	</xsl:template>

	<xsl:template match="DeclarationInfo">
		<xsl:if test="key('filterdeclarations',@name)">
			<xsl:call-template name="liobjlink">
				<xsl:with-param name="elem" select="current()" />
			</xsl:call-template>
			<xsl:for-each select="key('filterdcmembers',@name)">
				<xsl:call-template name="liobjlink">
					<xsl:with-param name="elem" select="current()" />
				</xsl:call-template>
			</xsl:for-each>
		</xsl:if>
	</xsl:template>

	<xsl:template match="ConstantInfo">
		<xsl:if test="key('filterconstants',@name)">
			<xsl:call-template name="liobjlink">
				<xsl:with-param name="elem" select="current()" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template match="RecordInfo">
		<xsl:if test="key('filterrecords',@name)">
			<xsl:call-template name="liobjlink">
				<xsl:with-param name="elem" select="current()" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template match="Coclasses|Interfaces|Declarations|Constants|Records">
		<xsl:call-template name="liobjlink">
			<xsl:with-param name="elem" select="current()" />
		</xsl:call-template>
	</xsl:template>
</xsl:stylesheet>
