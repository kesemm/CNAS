<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>List of CO Codes By Exchange By NPA</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: CompanyListByExchange.asp,v $
'* Commit Date:   $Date: 2006/05/17 16:01:02 $ (UTC)
'* Committed by:  $Author: SAIC-OTTAWA\browng $
'* CVS Revision:  $Revision: 1.4 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<form action="xca_NANPCANMenuInt.asp" method="post" id="formP4" name="formP4">
</form>
<font face="Arial Black" color="maroon" size="5">

<p align="center">List of CO Codes By Exchange By NPA</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange204.asp" target>NPA 204</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange250.asp" target>NPA 250</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange289.asp" target>NPA 289</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange306.asp" target>NPA 306</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange403.asp" target>NPA 403</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange416.asp" target>NPA 416</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange418.asp" target>NPA 418</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange450.asp" target>NPA 450</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange506.asp" target>NPA 506</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange514.asp" target>NPA 514</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange519.asp" target>NPA 519</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange604.asp" target>NPA 604</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange613.asp" target>NPA 613</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange647.asp" target>NPA 647</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange705.asp" target>NPA 705</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange709.asp" target>NPA 709</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange778.asp" target>NPA 778</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange780.asp" target>NPA 780</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange807.asp" target>NPA 807</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange819.asp" target>NPA 819</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange867.asp" target>NPA 867</a> </font></td>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange902.asp" target>NPA 902</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="20" HEIGHT="21"> </td>
    <td><font face="Arial"><a href="CompanyExchange905.asp" target>NPA 905</a> </font></td>
  </tr>
</table>
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
+"$RCSfile: CompanyListByExchange.asp,v $\n"
+"$Revision: 1.4 $\n"
+"$Date: 2006/05/17 16:01:02 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>

</html>
