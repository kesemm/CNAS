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
<title>Performance Tracking (by Calendar Days) Home Page</title>
<script language="JavaScript">
   function changeScreenSize()
   {self.moveTo(10,10)
	self.resizeTo(700,500)}
</script>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4" onload="changeScreenSize()">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<form action="xca_CalendarDaysPerformanceTracking.asp" method="post" id="formP4" name="formP4">
</form>
<font face="Arial Black" color="maroon" size="5">

<p align="center">Performance Tracking (by Calendar Days) Home Page</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="../images/arrow1rightred_e0.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="Processed.asp" target>Number of Codes Processed by
    Month</a> </font></td>
  </tr>
    <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="Period_Data.asp" target>Number of Codes Processed by Period (for EAC)</a> </font></td>
  </tr>

  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CalendarDaysMonthlyNewCOCodesTiming.asp" target>New CO Codes</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CalendarDaysMonthlyUpdateCOCodesTiming.asp" target>Update CO Codes</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CalendarDaysMonthlyRecoverCOCodesTiming.asp" target>Recover CO Codes</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CalendarDaysMonthlyESRDsTiming.asp" target>ESRDs</a> </font></td>
  </tr>

 </table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
