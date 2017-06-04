<?xml version="1.0" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output omit-xml-declaration="yes" method="xml" />
	<xsl:include href="common.xsl" />

	<xsl:template match="/">
		<HTML>
			<BODY>
				<UL>
					<xsl:apply-templates select="TypeLibs/TypeLibInfo" />
				</UL>
				<xsl:call-template name="objmerge">
					<xsl:with-param name="elem" select="TypeLibs/TypeLibInfo" />
				</xsl:call-template>
			</BODY>
		</HTML>
	</xsl:template>

	<xsl:template match="TypeLibInfo">
		<LI>
			<xsl:call-template name="objlink">
				<xsl:with-param name="elem" select="current()" />
			</xsl:call-template>
			<UL>
				<xsl:call-template name="level1">
					<xsl:with-param name="elem" select="Coclasses" />
					<xsl:with-param name="col" select="key('filtercoclasses',@guid)" />
				</xsl:call-template>
				<xsl:call-template name="level1">
					<xsl:with-param name="elem" select="Interfaces" />
					<xsl:with-param name="col" select="key('filterinterfaces',@guid)" />
				</xsl:call-template>
				<xsl:call-template name="level1">
					<xsl:with-param name="elem" select="Declarations" />
					<xsl:with-param name="col" select="key('filterdeclarations',@guid)" />
				</xsl:call-template>
				<xsl:call-template name="level1">
					<xsl:with-param name="elem" select="Constants" />
					<xsl:with-param name="col" select="key('filterconstants',@guid)" />
				</xsl:call-template>
				<xsl:call-template name="level1">
					<xsl:with-param name="elem" select="Records" />
					<xsl:with-param name="col" select="key('filterrecords',@guid)" />
				</xsl:call-template>
				<!-- * Ignore Unions/UnionInfo here, because vb can not use union. -->
			</UL>
		</LI>
	</xsl:template>
	
	<xsl:template name="level1">
		<xsl:param name="elem" />
		<xsl:param name="col" />

		<xsl:if test="count($col)>0">
			<li>
				<xsl:call-template name="objlink">
					<xsl:with-param name="elem" select="$elem" />
				</xsl:call-template>
				<UL>
					<xsl:call-template name="level2">
						<xsl:with-param name="col" select="$col" />
					</xsl:call-template>
				</UL>
			</li>
		</xsl:if>
	</xsl:template>

	<xsl:template name="level2">
		<xsl:param name="col" />

		<xsl:for-each select="$col">
			<xsl:sort select="@name"/>
			<LI>
				<xsl:choose>
					<xsl:when test="name()='CoclassInfo'">
						<xsl:call-template name="objlink">
							<xsl:with-param name="elem" select="current()" />
						</xsl:call-template>
						<UL>
							<xsl:call-template name="level3">
								<xsl:with-param name="col" select="key('filterifmembers',@defaulteventinterfaceguid)|key('filterifmembers',@defaultinterfaceguid)" />
							</xsl:call-template>
						</UL>
					</xsl:when>
					<xsl:when test="name()='InterfaceInfo'">
						<xsl:call-template name="objlink">
							<xsl:with-param name="elem" select="current()" />
						</xsl:call-template>
						<UL>
							<xsl:call-template name="level3">
								<xsl:with-param name="col" select="key('filterifmembers',@guid)" />
							</xsl:call-template>
						</UL>
					</xsl:when>
					<xsl:when test="name()='DeclarationInfo'">
						<xsl:call-template name="objlink">
							<xsl:with-param name="elem" select="current()" />
						</xsl:call-template>
						<UL>
							<xsl:call-template name="level3">
								<xsl:with-param name="col" select="key('filterdcmembers',@name)" />
							</xsl:call-template>
						</UL>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="objlink">
							<xsl:with-param name="elem" select="current()" />
							<xsl:with-param name="image" select="'true'" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</LI>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="level3">
		<xsl:param name="col" />

		<xsl:for-each select="$col">
			<xsl:sort select="@name"/>
			<li>
				<xsl:call-template name="objlink">
					<xsl:with-param name="elem" select="current()" />
					<xsl:with-param name="image" select="'true'" />
				</xsl:call-template>
			</li>
		</xsl:for-each>
	</xsl:template>

</xsl:stylesheet>
