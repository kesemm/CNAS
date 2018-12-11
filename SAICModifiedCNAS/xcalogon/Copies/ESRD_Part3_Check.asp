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
<title>NPA / NXX Querry</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" -->
</form>
<%
NPA=Request.Form("NPA")
session("NPA")=NPA
uname = session("UserUserName")
UserEntityID=session("UserEntityID")
SET objConnection = server.createobject("ADODB.connection")
SET rst =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sql1 = "Select ESRD from xca_ESRD where Status='A' and NPA="&NPA&" and EntityID="&UserEntityID&""
SET rst = objConnection.execute(sql1)
If Not rst.BOF Then
Response.Redirect "ESRD_Part3.asp"
Else
Response.Redirect "xca_MenuMain.asp"
End if
%>

<%objConnection.close %>

</body>
</html>
