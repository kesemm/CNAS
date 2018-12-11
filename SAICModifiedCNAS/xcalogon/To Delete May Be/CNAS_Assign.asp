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
<title>CNAS Assign Database Query</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body onload="DoOnLoad()" text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<SCRIPT language="javascript">
function DoOnLoad() {
/*
********************************************************************************
* Purpose: Function invoked on every page load.
********************************************************************************
*/
	//Resize Browser Window - specifed in (Width,Height)
	window.resizeTo(800,700)
		
}
</SCRIPT>
<form ACTION="Assign_RateCentre_query.asp" METHOD="get">
  <p>Enter a NPA:<input TYPE="text" NAME="NPA" SIZE="3" MAXLENGTH="3"><br>
  Enter a NXX:<input TYPE="text" NAME="NXX" SIZE="3" MAXLENGTH="3"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
</form>
</body>
</html>
