<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS Supplementary Menu</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: CNAS_Supp_Menu.asp,v $
'* Commit Date:   $Date: 2004/12/03 17:12:21 $ (UTC)
'* Committed by:  $Author: WalshKel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"--></FORM>
<FORM action="xca_NANPCANMenuInt.asp" method="post" id="formP4" name="formP4"></FORM>
<CENTER><FONT face="Arial Black" color="maroon" size="5"><STRONG>CNAS SUPPLEMENTARY MENU</STRONG></FONT></CENTER>
<BR>
<TABLE align="center" border="0" cellpadding="1" cellspacing="1" width="50%">
<TR>
<TD><BR></TD>
<TD><FONT face="Arial Black" color="maroon" size="3">CNAS ASP.Net Applications Launcher</FONT></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A target="blank" href="CNAS_Supp_Launcher.asp?s=1"><FONT face="Arial">Requested Code Validator - EAS Check</FONT></A></TD>
</TR>
<TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A target="blank" href="CNAS_Supp_Launcher.asp?s=2"><FONT face="Arial">CNA Date Calculator</FONT></A></TD>
</TR>
</TABLE>
<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %><BR>
<%
' THIS IS THE VERSION CONTROL INFORMATION BLOCK
' ---------------------------------------------
'
' Subdued input text box, that when clicked will make an alert with CVS Info
'
%>
<br><br>
<INPUT TYPE="TEXT" 
       STYLE="border: none; background-color: #D7C7A4; font: 7pt Arial; color: gray; width: 200px" 
       ONCLICK="VerInfo()" VALUE="CNAS Version Control Information"
       READONLY>
<SCRIPT language="JavaScript">
function VerInfo()
{
var strAlertText
strAlertText="SAIC Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: CNAS_Supp_Menu.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2004/12/03 17:12:21 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
