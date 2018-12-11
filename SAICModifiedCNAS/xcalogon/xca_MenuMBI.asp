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
<title>MBI Main Menu</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>


<font face="Arial Black" color="maroon" size="5">

<p align="center">MBI Main Menu</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MBI_Query.asp" target>MBI NPA Database Lookup</a> </font></td>
  </tr>
 <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MBI_Full_NPA_NXX.asp" target>MBI Application - Part 1: Full Block based on NPA-NXX (Select Rate Centre)</a> </font></td>
  </tr>
     <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MBI_Full_NPA_NXX_No_RC.asp" target>MBI Application - Part 1: Full Block based on NPA-NXX (No Rate Centre)</a> </font></td>
  </tr>
   <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MBI_Full_NPA_RC.asp" target>MBI Application - Part 1: Full Block based on NPA-Rate Centre</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MBI_NPA_NXX_Blocks.asp" target>MBI Application - Part 1: Partial Blocks based on NPA-NXX</a> </font></td>
  </tr>

<tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MBI_Part3Pre.asp" target>MBI In-Service - Part 3</a> </font></td>
  </tr>

</table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
