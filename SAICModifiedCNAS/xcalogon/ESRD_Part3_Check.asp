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

'GET BOTH NPA-NXX SELECTIONS & the NXX Chosen
NPA511=Request.Form("NPA511")
NPA211=Request.Form("NPA211")
SelectedNXX=Request.Form("NXXSelect")

' Set this NXX to session
session("SelectedNXX")=SelectedNXX

' Figure out which NXX to use
if SelectedNXX="211" Then
	SelectedNPA=NPA211
Else
	SelectedNPA=NPA511
end if
session("SelectedNPA")=SelectedNPA

SET objConnection = server.createobject("ADODB.connection")
SET rst =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sql1 = "Select ESRD from xca_ESRD where Status='A' and NPA="&SelectedNPA&" and NXX="&SelectedNXX&" and EntityID="&UserEntityID&""
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
