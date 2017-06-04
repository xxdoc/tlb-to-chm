<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" omit-xml-declaration="yes" />
	<xsl:include href="common.xsl" />
	<xsl:template match="TypeLibInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(Type Library Reference)</title>
		</head>
		<body>
		<xsl:call-template name="firstpart">
			<xsl:with-param name="name" select="@name" />
			<xsl:with-param name="type" select="'Type Library Reference'" />
			<xsl:with-param name="majorver" select="@majorversion" />
			<xsl:with-param name="minorver" select="@minorversion" />
			<xsl:with-param name="attributes" select="@attributestrings" />
			<xsl:with-param name="helpstring" select="@helpstring" />
		</xsl:call-template>
		<xsl:call-template name="Coclasses">
			<xsl:with-param name="guid" select="@guid" />
		</xsl:call-template>
		<xsl:call-template name="Interfaces">
			<xsl:with-param name="guid" select="@guid" />
		</xsl:call-template>
		<xsl:call-template name="Constants">
			<xsl:with-param name="guid" select="@guid" />
		</xsl:call-template>
		<xsl:call-template name="Declarations">
			<xsl:with-param name="guid" select="@guid" />
		</xsl:call-template>
		<xsl:call-template name="Records">
			<xsl:with-param name="guid" select="@guid" />
		</xsl:call-template>
		<!-- * Ignore Unions/UnionInfo here, because vb can not use union. -->
		<br />
		<xsl:text>Containing File: </xsl:text><xsl:value-of select="@containingfile" /><br />
		<xsl:text>TypeLib GUID: </xsl:text><xsl:value-of select="@guid" /><br />
		<xsl:call-template name="poweredby" />
		<br />
		</body>
		</html>
	</xsl:template>

	<xsl:template match="Coclasses">
		<html>
		<head>
		<title><xsl:value-of select="ancestor::TypeLibInfo/@name" />.Coclasses</title>
		</head>
		<body>
		<xsl:call-template name="Coclasses">
			<xsl:with-param name="guid" select="ancestor::TypeLibInfo/@guid" />
		</xsl:call-template>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<xsl:template name="Coclasses">
		<xsl:param name="guid" />
		<xsl:if test="key('filtercoclasses',$guid)">
			<H2><xsl:text>Classes</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filtercoclasses',$guid)">
					<tr>
						<td>
							<xsl:call-template name="link">
								<xsl:with-param name="htmlname" select="@htmlname" />
								<xsl:with-param name="name" select="@name" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:value-of select="@shorthelpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
		</xsl:if>
	</xsl:template>

	<xsl:template match="Interfaces">
		<html>
		<head>
		<title><xsl:value-of select="ancestor::TypeLibInfo/@name" />.Interfaces</title>
		</head>
		<body>
		<xsl:call-template name="Interfaces">
			<xsl:with-param name="guid" select="ancestor::TypeLibInfo/@guid" />
		</xsl:call-template>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<xsl:template name="Interfaces">
		<xsl:param name="guid" />
		<xsl:if test="key('filterinterfaces',$guid)">
			<H2><xsl:text>Interfaces</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterinterfaces',$guid)">
					<tr>
						<td>
							<xsl:call-template name="link">
								<xsl:with-param name="htmlname" select="@htmlname" />
								<xsl:with-param name="name" select="@name" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:value-of select="@shorthelpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
		</xsl:if>
	</xsl:template>

	<xsl:template match="Constants">
		<html>
		<head>
		<title><xsl:value-of select="ancestor::TypeLibInfo/@name" />.Constants</title>
		</head>
		<body>
		<xsl:call-template name="Constants">
			<xsl:with-param name="guid" select="ancestor::TypeLibInfo/@guid" />
		</xsl:call-template>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<xsl:template name="Constants">
		<xsl:param name="guid" />
		<xsl:if test="key('filterconstants',$guid)">
			<H2><xsl:text>Constants</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Type</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterconstants',$guid)">
					<tr>
						<td>
							<xsl:call-template name="link">
								<xsl:with-param name="htmlname" select="@htmlname" />
								<xsl:with-param name="name" select="@name" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:value-of select="@typekindstring" />
						</td>
						<td>
							<xsl:value-of select="@shorthelpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
		</xsl:if>
	</xsl:template>

	<xsl:template match="Declarations">
		<html>
		<head>
		<title><xsl:value-of select="ancestor::TypeLibInfo/@name" />.Declarations</title>
		</head>
		<body>
		<xsl:call-template name="Declarations">
			<xsl:with-param name="guid" select="ancestor::TypeLibInfo/@guid" />
		</xsl:call-template>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<xsl:template name="Declarations">
		<xsl:param name="guid" />
			<xsl:if test="key('filterdeclarations',$guid)">
			<H2><xsl:text>Declarations</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterdeclarations',$guid)">
					<tr>
						<td>
							<xsl:call-template name="link">
								<xsl:with-param name="htmlname" select="@htmlname" />
								<xsl:with-param name="name" select="@name" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:value-of select="@shorthelpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
		</xsl:if>
	</xsl:template>

	<xsl:template match="Records">
		<html>
		<head>
		<title><xsl:value-of select="ancestor::TypeLibInfo/@name" />.Records</title>
		</head>
		<body>
		<xsl:call-template name="Records">
			<xsl:with-param name="guid" select="ancestor::TypeLibInfo/@guid" />
		</xsl:call-template>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<xsl:template name="Records">
		<xsl:param name="guid" />
		<xsl:if test="key('filterrecords',$guid)">
			<H2><xsl:text>Records</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterrecords',$guid)">
					<tr>
						<td>
							<xsl:call-template name="link">
								<xsl:with-param name="htmlname" select="@htmlname" />
								<xsl:with-param name="name" select="@name" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:value-of select="@shorthelpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
		</xsl:if>
	</xsl:template>

	<xsl:template match="CoclassInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(Class)</title>
		</head>
		<body>
		<!-- For accelerate -->
		<xsl:if test="not(contains(@attributestrings,'hidden') or substring(@name,1,1)='_')">
			<xsl:call-template name="firstpart">
				<xsl:with-param name="name" select="@name" />
				<xsl:with-param name="type" select="'Class'" />
				<xsl:with-param name="majorver" select="@majorversion" />
				<xsl:with-param name="minorver" select="@minorversion" />
				<xsl:with-param name="attributes" select="@attributestrings" />
				<xsl:with-param name="helpstring" select="@helpstring" />
			</xsl:call-template>
			<xsl:if test="key('filterifmembers',@defaulteventinterfaceguid)">
				<H2><xsl:text>Event List (Interface </xsl:text><xsl:value-of select="key('interface',@defaulteventinterfaceguid)/@name" /><xsl:text>)</xsl:text></H2>
				<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
					<tr>
						<td><b>
							<xsl:text>Name</xsl:text>
						</b></td>
						<td><b>
							<xsl:text>Description</xsl:text>
						</b></td>
					</tr>
					<xsl:for-each select="key('filterifmembers',@defaulteventinterfaceguid)">
						<tr>
							<td>
								<xsl:call-template name="link">
									<xsl:with-param name="htmlname" select="@htmlname" />
									<xsl:with-param name="name" select="@name" />
								</xsl:call-template>
							</td>
							<td>
								<xsl:value-of select="@shorthelpstring" />
							</td>
						</tr>
					</xsl:for-each>
				</table>
				<br />
				<HR />
			</xsl:if>
			<xsl:if test="key('filterifmembers',@defaultinterfaceguid)">
				<H2><xsl:text>Member List (Interface </xsl:text><xsl:value-of select="key('interface',@defaultinterfaceguid)/@name" /><xsl:text>)</xsl:text></H2>
				<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
					<tr>
						<td><b>
							<xsl:text>Name</xsl:text>
						</b></td>
						<td><b>
							<xsl:text>Type</xsl:text>
						</b></td>
						<td><b>
							<xsl:text>Description</xsl:text>
						</b></td>
					</tr>
					<xsl:for-each select="key('filterifmembers',@defaultinterfaceguid)">
						<tr>
							<td>
								<xsl:call-template name="link">
									<xsl:with-param name="htmlname" select="@htmlname" />
									<xsl:with-param name="name" select="@name" />
								</xsl:call-template>
							</td>
							<td>
								<xsl:call-template name="memberkind">
									<xsl:with-param name="mem" select="current()" />
								</xsl:call-template>
							</td>
							<td>
								<xsl:value-of select="@shorthelpstring" />
							</td>
						</tr>
					</xsl:for-each>
				</table>
				<br />
				<HR />
			</xsl:if>
			<xsl:call-template name="lastpart">
				<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
			</xsl:call-template>
		</xsl:if>
		</body>
		</html>
	</xsl:template>

	<xsl:template match="InterfaceInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(Interface)</title>
		</head>
		<body>
		<!-- For accelerate -->
		<xsl:if test="not(contains(@attributestrings,'hidden') or substring(@name,1,1)='_' or @name='IDispatch' or @name='IUnknown')">
			<xsl:call-template name="firstpart">
				<xsl:with-param name="name" select="@name" />
				<xsl:with-param name="type" select="'Interface'" />
				<xsl:with-param name="majorver" select="@majorversion" />
				<xsl:with-param name="minorver" select="@minorversion" />
				<xsl:with-param name="attributes" select="@attributestrings" />
				<xsl:with-param name="helpstring" select="@helpstring" />
			</xsl:call-template>
			<H2><xsl:text>Member List</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Type</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterifmembers',@guid)">
					<tr>
						<td>
							<xsl:call-template name="link">
								<xsl:with-param name="htmlname" select="@htmlname" />
								<xsl:with-param name="name" select="@name" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:call-template name="memberkind">
								<xsl:with-param name="mem" select="current()" />
							</xsl:call-template>
						</td>
						<td>
							<xsl:value-of select="@shorthelpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
			<xsl:call-template name="lastpart">
				<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
			</xsl:call-template>
		</xsl:if>
		</body>
		</html>
	</xsl:template>

	<xsl:template match="DeclarationInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(Declaration)</title>
		</head>
		<body>
		<xsl:call-template name="firstpart">
			<xsl:with-param name="name" select="@name" />
			<xsl:with-param name="type" select="'Declaration'" />
			<xsl:with-param name="majorver" select="@majorversion" />
			<xsl:with-param name="minorver" select="@minorversion" />
			<xsl:with-param name="attributes" select="@attributestrings" />
			<xsl:with-param name="helpstring" select="@helpstring" />
		</xsl:call-template>
		<H2><xsl:text>Member List</xsl:text></H2>
		<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
			<tr>
				<td><b>
					<xsl:text>Name</xsl:text>
				</b></td>
				<td><b>
					<xsl:text>Description</xsl:text>
				</b></td>
			</tr>
			<xsl:for-each select="key('filterdcmembers',@name)">
				<tr>
					<td>
						<xsl:call-template name="link">
							<xsl:with-param name="htmlname" select="@htmlname" />
							<xsl:with-param name="name" select="@name" />
						</xsl:call-template>
					</td>
					<td>
						<xsl:value-of select="@shorthelpstring" />
					</td>
				</tr>
			</xsl:for-each>
		</table>
		<br />
		<HR />
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<xsl:template match="ConstantInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(<xsl:value-of select="@typekindstring" />)</title>
		</head>
		<body>
		<xsl:call-template name="firstpart">
			<xsl:with-param name="name" select="@name" />
			<xsl:with-param name="type" select="@typekindstring" />
			<xsl:with-param name="majorver" select="@majorversion" />
			<xsl:with-param name="minorver" select="@minorversion" />
			<xsl:with-param name="attributes" select="@attributestrings" />
			<xsl:with-param name="helpstring" select="@helpstring" />
		</xsl:call-template>
		<xsl:if test="key('filterctmembers',@name)">
			<H2><xsl:text>Member List</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Value</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterctmembers',@name)">
					<tr>
						<td>
							<xsl:value-of select="@name" />
						</td>
						<td>
							<nobr><xsl:value-of select="@value" /></nobr>
						</td>
						<td>
							<xsl:value-of disable-output-escaping="yes" select="@helpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
		</xsl:if>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>
	
	<xsl:template match="RecordInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(<xsl:value-of select="@typekindstring" />)</title>
		</head>
		<body>
		<xsl:call-template name="firstpart">
			<xsl:with-param name="name" select="@name" />
			<xsl:with-param name="type" select="@typekindstring" />
			<xsl:with-param name="majorver" select="@majorversion" />
			<xsl:with-param name="minorver" select="@minorversion" />
			<xsl:with-param name="attributes" select="@attributestrings" />
			<xsl:with-param name="helpstring" select="@helpstring" />
		</xsl:call-template>
		<xsl:if test="key('filterrcmembers',@name)">
			<H2><xsl:text>Member List</xsl:text></H2>
			<table border="1" cellpadding="5" cols="2" frame="below" rules="rows">
				<tr>
					<td><b>
						<xsl:text>Name</xsl:text>
					</b></td>
					<td><b>
						<xsl:text>Description</xsl:text>
					</b></td>
				</tr>
				<xsl:for-each select="key('filterrcmembers',@name)">
					<tr>
						<td>
							<xsl:value-of select="@name" />
						</td>
						<td>
							<xsl:value-of disable-output-escaping="yes" select="@helpstring" />
						</td>
					</tr>
				</xsl:for-each>
			</table>
			<br />
			<HR />
		</xsl:if>
		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
		</xsl:call-template>
		</body>
		</html>
	</xsl:template>

	<!-- 
	InfoNodes include member(function/sub/event etc.) page:
		InterfaceInfo
		DeclarationInfo
	-->
	<xsl:template match="MemberInfo">
		<html>
		<head>
		<title><xsl:value-of select="@name" />(Member)</title>
		</head>
		<body>
		<xsl:call-template name="firstpart">
			<xsl:with-param name="name" select="@name" />
			<xsl:with-param name="mem" select="current()" />
			<xsl:with-param name="attributes" select="@attributestrings" />
		</xsl:call-template>

		<b><xsl:text>Syntax</xsl:text></b><br />
		<xsl:call-template name="prototype">
			<xsl:with-param name="mem" select="current()" />
		</xsl:call-template>
		<xsl:value-of disable-output-escaping="yes" select="@helpstring" /><br />
		<br />
		<hr />

		<xsl:call-template name="lastpart">
			<xsl:with-param name="libnode" select="ancestor::TypeLibInfo" />
			<xsl:with-param name="ifnode" select="key('class',ancestor::InterfaceInfo/@guid)|ancestor::InterfaceInfo|ancestor::DeclarationInfo" />
		</xsl:call-template>
		<br />
		</body>
		</html>
	</xsl:template>

	<xsl:template name="prototype">
		<xsl:param name="mem" />

		<pre>
			<xsl:call-template name="memberkind">
				<xsl:with-param name="mem" select="$mem" />
			</xsl:call-template>
			<xsl:text> </xsl:text>
			<B><xsl:value-of select="$mem/@name" /></B>
			<xsl:text>(</xsl:text>
			
			<xsl:for-each select="$mem/Parameters/ParameterInfo">
				<xsl:if test="count($mem/Parameters/ParameterInfo)>1">
					<xsl:text> _</xsl:text><br />
					<xsl:text > </xsl:text>
				</xsl:if>

				<xsl:if test="@optional='True' or @default='True'">
					<xsl:text>[</xsl:text>
					<xsl:if test="$mem/Parameters/@optionalcount='-1'">
						<xsl:text>ParamArray </xsl:text>
					</xsl:if>
				</xsl:if>

				<xsl:if test="VarTypeInfo/@vbbyval='True'">
					<xsl:text>ByVal </xsl:text>
				</xsl:if>

				<I><xsl:value-of select="@name" /></I>

				<xsl:if test="VarTypeInfo/@vbarray='True'">
					<xsl:text>()</xsl:text>
				</xsl:if>

				<xsl:call-template name="vartype">
					<xsl:with-param name="var" select="VarTypeInfo" />
				</xsl:call-template>

				<xsl:if test="@optional='True' or @default='True'">
					<xsl:if test="VarTypeInfo/@vbdefval!=''">
						<xsl:text> = </xsl:text>
						<xsl:value-of select="VarTypeInfo/@vbdefval" />
					</xsl:if>
					<xsl:text>]</xsl:text>
				</xsl:if>

				<xsl:if test="following-sibling::ParameterInfo">
					<xsl:text>,</xsl:text>
				</xsl:if>

			</xsl:for-each>
			<xsl:text>)</xsl:text>
			<xsl:call-template name="vartype">
				<xsl:with-param name="var" select="ReturnType" />
			</xsl:call-template>

			<xsl:if test="$mem/ReturnType/@vbarray='True'">
				<xsl:text>()</xsl:text>
			</xsl:if>
		</pre>
	</xsl:template>

	<xsl:template name="vartype">
		<xsl:param name="var" />
		<xsl:if test="$var/@vbtypename!=''">
			<xsl:text> As </xsl:text>
			<xsl:choose>
				<xsl:when test="key('class',$var/@vbtypeguid)">
					<xsl:call-template name="link">
						<xsl:with-param name="htmlname" select="key('class',$var/@vbtypeguid)/@htmlname" />
						<xsl:with-param name="name" select="$var/@vbtypename" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="key('interface',$var/@vbtypeguid)">
					<xsl:call-template name="link">
						<xsl:with-param name="htmlname" select="key('interface',$var/@vbtypeguid)/@htmlname" />
						<xsl:with-param name="name" select="$var/@vbtypename" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="//ConstantInfo[@name=$var/@vbtypename]">
					<xsl:call-template name="link">
						<xsl:with-param name="htmlname" select="//ConstantInfo[@name=$var/@vbtypename]/@htmlname" />
						<xsl:with-param name="name" select="$var/@vbtypename" />
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$var/@vbtypename" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>

	<xsl:template name="propertykinds">
		<xsl:param name="mem" />
		<xsl:value-of select="$mem/@invokekindstring" />

		<xsl:for-each select="following-sibling::MemberInfo[@name=$mem/@name]">
			<xsl:text>,</xsl:text>
			<xsl:value-of select="@invokekindstring" />
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="firstpart">
		<xsl:param name="name" />
		<xsl:param name="type" />
		<xsl:param name="mem" />
		<xsl:param name="majorver" />
		<xsl:param name="minorver" />
		<xsl:param name="attributes" />
		<xsl:param name="helpstring" />
		
		<H1>
			<xsl:value-of select="$name" />
			<xsl:text> (</xsl:text>
			<xsl:choose>
				<xsl:when test="$type!=''">
					<xsl:value-of select="$type" />
				</xsl:when>
				<xsl:when test="$mem">
					<xsl:call-template name="memberkind">
						<xsl:with-param name="mem" select="$mem" />
					</xsl:call-template>
					<xsl:if test="$mem/@vbtype='Property'">
						<xsl:text> </xsl:text>
						<xsl:call-template name="propertykinds">
							<xsl:with-param name="mem" select="$mem" />
						</xsl:call-template>
					</xsl:if>
				</xsl:when>
			</xsl:choose>
			<xsl:text>)</xsl:text>
		</H1>
		<xsl:if test="$majorver!='' and $minorver!=''">
			<xsl:text>Version: </xsl:text><xsl:value-of select="$majorver" /><xsl:text>.</xsl:text><xsl:value-of select="$minorver" /><br />
		</xsl:if>
		<xsl:if test="$attributes!=''">
			<xsl:text>Attributes: </xsl:text><xsl:value-of select="$attributes" /><br />
		</xsl:if>
		<xsl:if test="$helpstring!=''">
			<br />
			<hr />
			<br />
			<xsl:value-of disable-output-escaping="yes" select="$helpstring" /><br />
		</xsl:if>
		<br />
		<HR />
	</xsl:template>

	<xsl:template name="lastpart">
		<xsl:param name="libnode" />
		<xsl:param name="ifnode" />

		<xsl:if test="$libnode">
			<xsl:text>Apply to: </xsl:text>
			<xsl:call-template name="link">
				<xsl:with-param name="htmlname" select="$libnode/@htmlname" />
				<xsl:with-param name="name" select="$libnode/@name" />
			</xsl:call-template>
		</xsl:if>
		<xsl:if test="$ifnode">
			<xsl:text>.</xsl:text>
			<xsl:call-template name="link">
				<xsl:with-param name="htmlname" select="$ifnode/@htmlname" />
				<xsl:with-param name="name" select="$ifnode/@name" />
			</xsl:call-template>
		</xsl:if>

		<xsl:call-template name="poweredby" />
	</xsl:template>

	<xsl:template name="poweredby">
		<br />
		<p align="right"><i><xsl:text>Converted by tlb-to-chm.</xsl:text></i> </p>
	</xsl:template>
</xsl:stylesheet>
