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

session("MBI_0")=Request.Form("MBI_0")
session("MBI_1")=Request.Form("MBI_1")
session("MBI_2")=Request.Form("MBI_2")
session("MBI_3")=Request.Form("MBI_3")
session("MBI_4")=Request.Form("MBI_4")
session("MBI_5")=Request.Form("MBI_5")
session("MBI_6")=Request.Form("MBI_6")
session("MBI_7")=Request.Form("MBI_7")
session("MBI_8")=Request.Form("MBI_8")
session("MBI_9")=Request.Form("MBI_9")

aMBI_0=session("MBI_0")
aMBI_1=session("MBI_1")
aMBI_2=session("MBI_2")
aMBI_3=session("MBI_3")
aMBI_4=session("MBI_4")
aMBI_5=session("MBI_5")
aMBI_6=session("MBI_6")
aMBI_7=session("MBI_7")
aMBI_8=session("MBI_8")
aMBI_9=session("MBI_9")

If aMBI_0 <> 10 Then aMBI_0=0
If aMBI_1 <> 11 Then aMBI_1=0
If aMBI_2 <> 12 Then aMBI_2=0
If aMBI_3 <> 13 Then aMBI_3=0
If aMBI_4 <> 14 Then aMBI_4=0
If aMBI_5 <> 15 Then aMBI_5=0
If aMBI_6 <> 16 Then aMBI_6=0
If aMBI_7 <> 17 Then aMBI_7=0
If aMBI_8 <> 18 Then aMBI_8=0
If aMBI_9 <> 19 Then aMBI_9=0

AuthorizedRep=Replace(Request.Form("AuthorizedRep"),"'","''")
AuthorizedRepTitle=Replace(Request.Form("AuthorizedRepTitle"),"'","''")
ApplicationDate=Request.Form("ApplicationDate")
SET objConnection = server.createobject("ADODB.connection")
SET rstQry =server.createobject("ADODB.recordset")
ObjConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlQry = "Exec [Partial_In_Service_MBI] " & NPA & "," & NXX & ", " & EntityID & "," & UserID & ", '" & ApplicationDate & "'" & ", '" & AuthorizedRep & "'" & ",'" & AuthorizedRepTitle & "', " &aMBI_0 & ", " & aMBI_1 & ", " & aMBI_2 & ", " & aMBI_3 & ", " & aMBI_4 & ", " & aMBI_5 & ", " & aMBI_6 & ", " & aMBI_7 & ", " & aMBI_8 & ", " & aMBI_9
Response.Write sqlQry
SET rstQry = objConnection.execute(sqlQry)
Response.Redirect "MBI_Partial_In-Service_Confirm.asp"

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
