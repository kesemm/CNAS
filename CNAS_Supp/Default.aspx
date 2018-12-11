<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Default.aspx.vb" Inherits="CNAS_Supp._Default"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<%
'****************************************************************************************
'* Created by:    Kelly T. Walsh (Leidos Canada)
'* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
'* Purpose:       ASP.Net Page - Default.aspx
'*                The Default.aspx page is the first page launched and validates the inter-
'*                CNAS operation with simple security authorisation, and redirects to the
'*                requested functional page with URL parms.
'* CVS File:      Default.aspx,v
'* Commit Date:   2018/01/25 16:43:49 (UTC)
'* Committed by:  saic-ottawa\walshkel
'* CVS Revision:  1.1
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
	<HEAD>
		<title>Default</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body bgColor="#d7c7a4">
		<form id="Form1" method="post" runat="server">
			<P align="center">
				<FONT face="Arial Black" color="maroon" size="4"><STRONG>Close This Window</STRONG></FONT></P>
			<P align="center"><b><FONT face="Arial Black" size="2">You're CNAS Supplementary (.Net) session has ended by request or error.</FONT></b></P>
			<P align="center"><STRONG><FONT face="Arial Black" color="#800000" size="2">If you are 
						having technical difficulty,<br>please contact CNAS System Support</FONT></STRONG>&nbsp;</P>
			<P></P>
		</form>
	</body>
</HTML>
