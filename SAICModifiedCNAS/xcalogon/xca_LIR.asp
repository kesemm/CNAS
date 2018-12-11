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
<title>LIR Main Menu</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>


<font face="Arial Black" color="maroon" size="5">

<p align="center">LIR Main Menu</font> </p>
<p align="center">(Based on LERG 8 Data)</font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="AB_LIR.asp" target>Alberta</a> </font></td>
  </tr>
 <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="BC_LIR.asp" target>British Columbia</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="MB_LIR.asp" target>Manitoba</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="NB_LIR.asp" target>New Brunswick</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="NF_LIR.asp" target>Newfoundland and Labrador</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="NT_LIR.asp" target>Northwest Territories</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="NS_LIR.asp" target>Nova Scotia</a> </font></td>
  </tr>
   <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="VU_LIR.asp" target>Nunavut</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="ON_LIR.asp" target>Ontario</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="PE_LIR.asp" target>Prince Edward Island</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="PQ_LIR.asp" target>Quebec</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="SK_LIR.asp" target>Saskatchewan</a> </font></td>
  </tr>
   <tr>
    <td><img SRC="ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="YT_LIR.asp" target>Yukon</a> </font></td>
  </tr>
  
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>
