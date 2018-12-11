<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE></TITLE>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"--></FORM>
<CENTER><FONT face="Arial Black" color="maroon" size="5"><STRONG>ENTRY PAGE</STRONG></FONT><BR>
<BR>
<BR></CENTER>
<TABLE align="center" border="0" cellpadding="1" cellspacing="1" width="50%">
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A href="xca_MenuMainPost.asp"><FONT face="Arial">CNAS Main Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A href="xca_NANPCANMainPost.asp"><FONT face="Arial">CNAS Utilties Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A href="CNAS_Supp_Menu.asp"><FONT face="Arial">CNAS Supplementary Menu (.Net Apps)</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_AdminTrackingPost.asp"><FONT face="Arial">CNA Admin Tracking System</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="MLM_Launcher.asp"><FONT face="Arial">CNA Contacts &amp; Mailing List Manager</FONT></A></TD>
</TR>
<TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_MenuESRD.asp"><FONT face="Arial">ESRD Applications Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_MenuMBI.asp"><FONT face="Arial">MBI Applications Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_PerformanceTrackingPost.asp"><FONT face="Arial">CNA Performance Tracking (Working Days)</FONT></A><SMALL><FONT face="Arial">&nbsp;(Terry &amp; Glenn)</FONT></SMALL></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/undercon.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_CalendarDaysPerformanceTrackingPost.asp"><FONT face="Arial">CNA Performance Tracking (Calendar Days)</FONT></A><SMALL><FONT face="Arial">&nbsp;(Terry &amp; Glenn)</FONT></SMALL></TD>
</TR>

<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_WebSiteStatsPost.asp"><FONT face="Arial">CNA Web Site Stats</FONT></A></TD>
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
+"$RCSfile: xca_MenuMain.asp,v $\n"
+"$Revision: 1.10 $\n"
+"$Date: 2007/04/03 16:50:52 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
