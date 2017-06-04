<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="html" omit-xml-declaration="yes" />
	<xsl:template match="Comment|comment">
		<xsl:if test="Summary|summary">
			<b>Summary:</b><br />
			<pre>
				<xsl:value-of select="Summary|summary" />
			</pre>
		</xsl:if>
		<ul>
			<xsl:for-each select="Parameter|parameter">
				<li>
					<b><xsl:text>Parameter</xsl:text></b><xsl:text> </xsl:text><i><xsl:value-of select="@Name|@name" /></i><xsl:text>:</xsl:text><br />
					<xsl:value-of select="." /><br />
				</li>
			</xsl:for-each>
		</ul>
		<xsl:if test="ReturnValue|returnvalue">
			<b>Return Value:</b><br />
			<xsl:value-of select="ReturnValue|returnvalue" /><br />
		</xsl:if>
		<xsl:if test="Example|example">
			<b>Example:</b><br />
			<pre><xsl:value-of select="Example|example" /></pre>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
