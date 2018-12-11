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
<title>ESRD Main Menu</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<form action="xca_MenuESRD.asp" method="post" id="formP4" name="formP4">
</form>
<font face="Arial Black" color="maroon" size="5">

<p align="center">ESRD Main Menu</font> </p>

<p align="center">Please Stay Out!!!  Working on 211 Update</font> </p>
<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="ESRD_NPA_Query.asp" target>ESRD NPA Database Lookup (Based on NPA)</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="ESRD_Block_Query.asp" target>ESRD NPA Database Lookup (Based on NPA and Block)</a> </font></td>
  </tr>
 <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="ESRD.asp" target>ESRD Application - Part 1</a> </font></td>
  </tr>
<tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="ESRD_Part3Pre.asp" target>ESRD In-Service - Part 3</a> </font></td>
  </tr>
</table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
