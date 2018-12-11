<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>Code Applicant Menu</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: xca_MenuNANPCAN.asp,v $
'* Commit Date:   $Date: 2016/02/12 15:52:15 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.4 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"--></FORM>
<FORM action="xca_NANPCANMenuInt.asp" method="post" id="formP4" name="formP4"></FORM>
<CENTER><FONT face="Arial Black" color="maroon" size="5"><STRONG>CNAS UTILITIES MENU</STRONG></FONT></CENTER>
<BR>
<TABLE align="center" border="0" cellpadding="1" cellspacing="1" width="50%">
<TR>
<TD><BR></TD>
<TD><FONT face="Arial Black" color="maroon" size="3">Update CNAS Records</FONT></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="CNAS_SwitchID.asp" target=""><FONT face="Arial">CNAS SwitchID Update</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="CNAS_Remarks.asp" target=""><FONT face="Arial">CNAS Effective Date Update</FONT></A></TD>
</TR>
<TD><BR></TD>
<TD><FONT face="Arial Black" color="maroon" size="3">Query CNAS/LERG Information</FONT></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="LERG_OCN.asp" target=""><FONT face="Arial">Get a list of Canadian OCNs (i.e. LERG 1)</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPA_NXX_Query.asp" target=""><FONT face="Arial">NPA/NXX Lookup</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="Switch_Query.asp" target=""><FONT face="Arial">Switch Lookup (i.e. SRD/SHA)</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="cnas_switch_query.asp" target=""><FONT face="Arial">NPA-NXX for SwitchID Lookup</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="LRN_Query.asp" target=""><FONT face="Arial">LRN Lookup</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="UserLogon.asp" target=""><FONT face="Arial">Authorised CNAS Users</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="Ratecentre_Parms.asp" target=""><FONT face="Arial">Locality Lookup based on RC</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="Actual_Switch_Parms.asp" target=""><FONT face="Arial">Find NXXs using an Actual Switch</FONT></A></TD>
</TR>
<TD><BR></TD>
<TR>
<TD><BR></TD>
<TD><FONT face="Arial Black" color="maroon" size="3">Query CNAS Information</FONT></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="CompanyListQuery.asp" target=""><FONT face="Arial">List Of Companies with CO Codes</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="CompanyNPAQuery.asp" target=""><FONT face="Arial">CNAS Company/NPA Database Lookup</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPANXXStatusList_Parms.asp" target=""><FONT face="Arial">NPA NXX Status List (Availability)</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="Being_Recovered.asp" target=""><FONT face="Arial">Being Recovered Codes</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="Recovered_Aging.asp" target=""><FONT face="Arial">Recovered / Aging Codes</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPACompanyCOCodes_Parms.asp" target=""><FONT face="Arial">NPA CO Codes List By Company</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPASpecialCodes_Parms.asp" target=""><FONT face="Arial">CNA Special Codes in NPA</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="FutureNPAs.asp" target=""><FONT face="Arial">Future NPAs</FONT></A></TD>
</TR>
<TR>
<TD><BR></TD>
<TD><BR></TD>
</TR>
<TR>
<TD><BR></TD>
<TD><FONT face="Arial Black" color="maroon" size="3">Totals &amp; Counts</FONT></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="Total.asp" target=""><FONT face="Arial">Company Total By NPA</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPACountCodesAvailable.asp" target=""><FONT face="Arial">Number of Available CO Codes by NPAs</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPACOCodeMonthlyTotals_Parms.asp" target=""><FONT face="Arial">NPA CO Code Monthly Totals</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="NPACOCodeAnnualTotals_Parms.asp" target=""><FONT face="Arial">NPA CO Code Annual Totals</FONT></A></TD>
</TR>
<TR>
<TD><IMG src="ball25.gif" width="35" height="36"></TD>
<TD><A href="CompanyCountQuery.asp" target=""><FONT face="Arial">CNAS Company Count By NPA</FONT></A></TD>
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
strAlertText="Leidos Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: xca_MenuNANPCAN.asp,v $\n"
+"$Revision: 1.4 $\n"
+"$Date: 2016/02/12 15:52:15 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
