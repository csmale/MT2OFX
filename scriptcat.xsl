<?xml version="1.0"?>
<xsl:stylesheet
	version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
>
	<xsl:output method="html"/>
	<xsl:variable name="lang">@en</xsl:variable>
	<xsl:template match="mt2ofx">
		<html>
			<head>
				<title>XSL Test</title>
			</head>
			<body>
				<table frame='box' rules='groups' cellpadding='4'>
					<colgroup/><colgroup/><colgroup/><colgroup/>
				<thead><tr>
					<th>Bank</th>
					<th>Region</th>
					<th>Format</th>
					<th>Script name</th>
				</tr></thead>
				<tbody>
				<xsl:apply-templates/>
				</tbody>
				</table>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="bankscript">
		<tr>
		<xsl:variable name="region"><xsl:value-of select="region" /></xsl:variable>
		<xsl:variable name="fmt"><xsl:value-of select="@format" /></xsl:variable>
		<td><xsl:value-of select="@name" /></td>
		<td><xsl:value-of select="//region[@id=$region]/@en" /></td>
		<td><xsl:value-of select="//format[@id=$fmt]/@en" /></td>
		<td><xsl:value-of select="script" /></td>
		</tr>
	</xsl:template>
	<xsl:template match="version" />
	<xsl:template match="versiondate" />
	<xsl:template match="language" />
</xsl:stylesheet>
