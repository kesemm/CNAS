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
<title>CNAS Database Query</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<form ACTION="CNAS_Part4_Done.asp" METHOD="get">
  <p>Enter a NPA:<input TYPE="text" NAME="NPA" SIZE="3" MAXLENGTH="3"><br>
  Enter a NXX:<input TYPE="text" NAME="NXX" SIZE="3" MAXLENGTH="3"><br>
  Enter a the In-Service Date (yyyy-mm-dd):<input NAME="Date" SIZE="10" MAXLENGTH="10"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
</form>
</body>
</html>
