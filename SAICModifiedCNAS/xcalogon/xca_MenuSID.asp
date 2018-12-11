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
<title>SID Main Menu</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>


<font face="Arial Black" color="maroon" size="5">

<p align="center">SID Main Menu</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="SID.asp" target>All Canadian SID Data</a> </font></td>
  </tr>
 <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="SID-Assigned.asp" target>Assigned SIDs</a> </font></td>
  </tr>
     <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="SID-Available.asp" target>Available SIDs</a> </font></td>
  </tr>
 
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
