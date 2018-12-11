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
<title>Code Applicant Menu</title>
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

<form action="xca_PerformanceTracking.asp" method="post" id="formP4" name="formP4">
</form>
<font face="Arial Black" color="maroon" size="5">

<p align="center">Performance Tracking Home Page</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="Processed.asp" target>Number of CO Codes Processed by
    Month</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MonthlyNewCOCodesTiming.asp" target>New CO Codes</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MonthlyUpdateCOCodesTiming.asp" target>Update CO Codes</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MonthlyRecoverCOCodesTiming.asp" target>Recover CO Codes</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MonthlyESRDsTiming.asp" target>ESRDs</a> </font></td>
  </tr>

 </table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
