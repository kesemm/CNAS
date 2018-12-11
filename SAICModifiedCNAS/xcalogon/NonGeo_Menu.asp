<%@ Language=VBScript %>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NonGeo_Menu.asp,v $
'* Commit Date:   $Date: 2015/01/19 13:34:14 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
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
<title>Non-Geographic Codes Menu</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<font face="Arial Black" color="maroon" size="5">

<p align="center">Non-Geographic Codes Menu</font> </p>


<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CNAS_Supp_Launcher.asp?s=3">Non-Geographic Code Lookup</a> </font></td>
  </tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CNAS_Supp_Launcher.asp?s=4">Non-Geographic Code Application - Form A</a> </font></td>
  </tr>
<tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="CNAS_Supp_Launcher.asp?s=5">Non-Geographic Code In-Service - Form C</font></td>
  </tr>
</table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
