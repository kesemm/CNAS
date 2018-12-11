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
<title>NANPCAN Home Page</title>
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

<p><img src="undercon.gif" alt="[Under Construction]" border="0" WIDTH="40" HEIGHT="38"></p>
<font face="Arial Black" color="maroon" size="5">

<p align="center">Admin MENU</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="http://www.cnac.ca" target>CNA's Public Home Page</a></font>
    </td>
  </tr>
</table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
