<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:key name="interface" match="/TypeLibs/TypeLibInfo/Interfaces/InterfaceInfo" use="@guid" />
	<xsl:key name="eventclass" match="/TypeLibs/TypeLibInfo/Coclasses/CoclassInfo" use="@defaulteventinterfaceguid" />
	<xsl:key name="class" match="CoclassInfo" use="@defaulteventinterfaceguid|@defaultinterfaceguid|@guid" />

	<!-- TypeLibInfo.GUID=>Filtered InterfaceInfo Collection; InterfaceInfo.GUID=>Item/Null -->
	<xsl:key name="filterinterfaces" match="/TypeLibs/TypeLibInfo/Interfaces/InterfaceInfo[not(contains(@attributestrings,'hidden') or substring(@name,1,1)='_' or @name='IDispatch' or @name='IUnknown')]" use="ancestor::TypeLibInfo/@guid|@guid" />
	<!-- (With Member Pages) InterfaceInfo.GUID=>Filtered MemberInfo Collection -->
	<xsl:key name="filterifmembers" match="/TypeLibs/TypeLibInfo/Interfaces/InterfaceInfo/Members/MemberInfo[not(contains(@attributestrings,'hidden') or contains(@attributestrings,'restricted') or substring(@name,1,1)='_' or @name=preceding-sibling::MemberInfo/@name)]" use="ancestor::InterfaceInfo/@guid" />
	
	<!-- TypeLibInfo.GUID=>Filtered DeclarationInfo Collection; DeclarationInfo.name=>Item/Null -->
	<xsl:key name="filterdeclarations" match="/TypeLibs/TypeLibInfo/Declarations/DeclarationInfo[not(contains(@attributestrings,'hidden') or substring(@name,1,1)='_')]" use="ancestor::TypeLibInfo/@guid|@name" />
	<!-- (With Member Pages) DeclarationInfo.name=>Filtered MemberInfo Collection -->
	<xsl:key name="filterdcmembers" match="/TypeLibs/TypeLibInfo/Declarations/DeclarationInfo/Members/MemberInfo[not(contains(@attributestrings,'hidden') or contains(@attributestrings,'restricted') or substring(@name,1,1)='_' or @name=preceding-sibling::MemberInfo/@name)]" use="ancestor::DeclarationInfo/@name" />

	<!-- TypeLibInfo.GUID=>Filtered ConstantInfo Collection; ConstantInfo.name=>Item/Null -->
	<xsl:key name="filterconstants" match="/TypeLibs/TypeLibInfo/Constants/ConstantInfo[not(substring(@name,1,1)='_')]" use="ancestor::TypeLibInfo/@guid|@name" />
	<!-- ConstantInfo.name=>Filtered MemberInfo Collection -->
	<xsl:key name="filterctmembers" match="/TypeLibs/TypeLibInfo/Constants/ConstantInfo/Members/MemberInfo[not(substring(@name,1,1)='_')]" use="ancestor::ConstantInfo/@name" />

	<!-- TypeLibInfo.GUID=>Filtered CoclassInfo Collection; CoclassInfo.GUID=>Item/Null -->
	<xsl:key name="filtercoclasses" match="/TypeLibs/TypeLibInfo/Coclasses/CoclassInfo[not(contains(@attributestrings,'hidden') or substring(@name,1,1)='_')]" use="ancestor::TypeLibInfo/@guid|@guid" />

	<!-- TypeLibInfo.GUID=>Filtered RecordInfo Collection; RecordInfo.name=>Item/Null -->
	<xsl:key name="filterrecords" match="/TypeLibs/TypeLibInfo/Records/RecordInfo[not(substring(@name,1,1)='_')]" use="ancestor::TypeLibInfo/@guid|@name" />
	<!-- RecordInfo.name=>Filtered MemberInfo Collection -->
	<xsl:key name="filterrcmembers" match="/TypeLibs/TypeLibInfo/Records/RecordInfo/Members/MemberInfo[not(substring(@name,1,1)='_')]" use="ancestor::RecordInfo/@name" />

	<!-- * Ignore Unions/UnionInfo here, because vb can not use union. -->

	<xsl:template name="memberkind">
		<xsl:param name="mem" />
		<xsl:choose>
			<xsl:when test="key('eventclass',$mem/ancestor::InterfaceInfo/@guid)">
				<xsl:text>Event</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$mem/@vbtype" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template name="obj">
		<xsl:param name="name" />

		<OBJECT type="text/sitemap">
			<param name="Name">
				<xsl:attribute name="value">
					<xsl:value-of select="$name" />
				</xsl:attribute>
			</param>
		</OBJECT>
	</xsl:template>

	<xsl:template name="link">
		<xsl:param name="htmlname" />
		<xsl:param name="name" />
		<B>
		<A>
			<xsl:attribute name="href">
				<xsl:text>html/</xsl:text>
				<xsl:value-of select="$htmlname" />
				<xsl:text>.htm</xsl:text>
			</xsl:attribute>
			<xsl:value-of select="$name" />
		</A>
		</B>
	</xsl:template>

	<xsl:template name="objlink">
		<xsl:param name="elem" />
		<xsl:param name="image" />
		
		<OBJECT type="text/sitemap">
			<param name="Name">
				<xsl:attribute name="value">
					<xsl:choose>
						<xsl:when test="$elem/@name">
							<xsl:value-of select="$elem/@name" />
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="name($elem)" />
						</xsl:otherwise>
					</xsl:choose>
					<xsl:if test="$elem/@vbtype">
						<xsl:text> </xsl:text>
						<xsl:call-template name="memberkind">
							<xsl:with-param name="mem" select="$elem" />
						</xsl:call-template>
					</xsl:if>
				</xsl:attribute>
			</param>
			<param name="Local">
				<xsl:attribute name="value">
					<xsl:text>html/</xsl:text>
					<xsl:value-of select="$elem/@htmlname" />
					<xsl:text>.htm</xsl:text>
				</xsl:attribute>
			</param>
			<xsl:if test="$image">
				<param name="ImageNumber" value="11" />
			</xsl:if>
		</OBJECT>
	</xsl:template>

	<xsl:template name="liobjlink">
		<xsl:param name="elem" />
		<li>
			<xsl:call-template name="objlink">
				<xsl:with-param name="elem" select="$elem" />
			</xsl:call-template>
		</li>
	</xsl:template>

	<xsl:template name="objmerge">
		<xsl:param name="elem" />
		<xsl:if test="$elem/@helpfile!=''">
			<OBJECT type="text/sitemap">
				<param name="Merge">
					<xsl:attribute name="value">
						<xsl:value-of select="$elem/@helpref" />
						<xsl:text>::\</xsl:text>
						<xsl:value-of select="$elem/@hhcname" />
					</xsl:attribute>
				</param>
			</OBJECT>
		</xsl:if>
	</xsl:template>

</xsl:stylesheet>

