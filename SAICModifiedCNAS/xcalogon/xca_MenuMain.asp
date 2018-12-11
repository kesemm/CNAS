<%@ Language=VBScript %>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: xca_MenuMain.asp,v $
'* Commit Date:   $Date: 2017/06/07 18:00:45 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.6 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
<%
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
<TD colspan="3"><A href="xca_NANPCANMainPost.asp"><FONT face="Arial">CNAS Utilities Menu</FONT></A></TD>
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
<TD colspan="3"><A target="blank" href="OpenCNATasks_some.asp"><FONT face="Arial">Open CNA Tasks that have a due date within the next 14 days </FONT></A></TD>
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
<TD colspan="3"><A target="blank" href="NonGeo_Menu.asp"><FONT face="Arial">Non-Geographic Applications Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/undercon.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_MenuSID.asp"><FONT face="Arial">SID Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_LIR.asp"><FONT face="Arial">LIR Menu</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="http://crtc.gc.ca/8740/eng/"><FONT face="Arial">Tariff Applications</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="http://www.crtc.gc.ca/eng/comm/telecom/eslcclec.htm"><FONT face="Arial">CLEC Files (Select year under Documents from CLECS:)</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="http://www.ic.gc.ca/eic/site/smt-gst.nsf/eng/sf08464.html"><FONT face="Arial">Combined Auctions - 2300 MHz and 3500 MHz</FONT></A></TD>
</TR>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="http://www.ic.gc.ca/eic/site/smt-gst.nsf/eng/sf05471.html"><FONT face="Arial">Additional PCS Spectrum - 2 GHz</FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/ball25.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="http://www.ic.gc.ca/eic/site/smt-gst.nsf/eng/sf02103.html"><FONT face="Arial">24 and 38 GHz Auction </FONT></A></TD>
</TR>
<TR>
<TD width="25"><IMG height="36" src="../images/arrow1rightred_e0.gif" width="35"></TD>
<TD colspan="3"><A target="blank" href="xca_CalendarDaysPerformanceTrackingPost.asp"><FONT face="Arial">CNA Performance Tracking</FONT></A><SMALL><FONT face="Arial">&nbsp;(Calendar Days)</FONT></SMALL></TD>
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
+"$Revision: 1.6 $\n"
+"$Date: 2017/06/07 18:00:45 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
