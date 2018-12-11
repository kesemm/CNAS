<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head><script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

EntityID=cint(session("UserEntityID"))
UserID=session("UserUserID")
NPA=session("aNPA")
NXX=session("aNXX")
session("aRC")= "-"
RC=session("aRC")
AuthorizedRep=Replace(Request.Form("AuthorizedRep"),"'","''")
AuthorizedRepTitle=Replace(Request.Form("AuthorizedRepTitle"),"'","''")
SupportingExplanation=Replace(Request.Form("SupportingExplanation"),"'","''")
ApplicationDate=Request.Form("ApplicationDate")
OCN=Request.Form("OCN")
SET objConnection = server.createobject("ADODB.connection")
SET rstQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlQry = "Exec [MBI_Part1s] " & NPA & "," & NXX & ", '" & OCN & "', " & EntityID & "," & UserID & ", '" & ApplicationDate & "'" & ", '" & AuthorizedRep & "'" & ",'" & AuthorizedRepTitle & "'" & ",'" & RC & "'" & ", '" & SupportingExplanation & "'"
SET rstQry = objConnection.execute(sqlQry)
Response.Redirect "MBI_NPA_NXX_Confirm.asp"
</script>

<title></title>
</head>

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">

<form name="thisForm" METHOD="post">
<!--#Include file="xca_CNASlib.inc"-->
</form>

</BODY>

<p>&nbsp;</p>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</body>

</html>
