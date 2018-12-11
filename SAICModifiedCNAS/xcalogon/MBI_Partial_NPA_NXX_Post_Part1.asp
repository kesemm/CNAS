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

session("aRC_0")=Replace(Request.Form("RateCenterAssignLookup_0"),"'","''")
session("aRC_1")=Replace(Request.Form("RateCenterAssignLookup_1"),"'","''")
session("aRC_2")=Replace(Request.Form("RateCenterAssignLookup_2"),"'","''")
session("aRC_3")=Replace(Request.Form("RateCenterAssignLookup_3"),"'","''")
session("aRC_4")=Replace(Request.Form("RateCenterAssignLookup_4"),"'","''")
session("aRC_5")=Replace(Request.Form("RateCenterAssignLookup_5"),"'","''")
session("aRC_6")=Replace(Request.Form("RateCenterAssignLookup_6"),"'","''")
session("aRC_7")=Replace(Request.Form("RateCenterAssignLookup_7"),"'","''")
session("aRC_8")=Replace(Request.Form("RateCenterAssignLookup_8"),"'","''")
session("aRC_9")=Replace(Request.Form("RateCenterAssignLookup_9"),"'","''")

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

If aMBI_0 = 10 Then
	RC_0=session("aRC_0")
Else
	RC_0=""
End If

If aMBI_1 = 11 Then
	RC_1=session("aRC_1")
Else
	RC_1=""
End If

If aMBI_2 = 12 Then
	RC_2=session("aRC_2")
Else
	RC_2=""
End If

If aMBI_3 = 13 Then
	RC_3=session("aRC_3")
Else
	RC_3=""
End If

If aMBI_4 = 14 Then
	RC_4=session("aRC_4")
Else
	RC_4=""
End If

If aMBI_5 = 15 Then
	RC_5=session("aRC_5")
Else
	RC_5=""
End If

If aMBI_6 = 16 Then
	RC_6=session("aRC_6")
Else
	RC_6=""
End If

If aMBI_7 = 17 Then
	RC_7=session("aRC_7")
Else
	RC_7=""
End If

If aMBI_8 = 18 Then
	RC_8=session("aRC_8")
Else
	RC_8=""
End If

If aMBI_9 = 19 Then
	RC_9=session("aRC_9")
Else
	RC_9=""
End If

RC_1=session("aRC_1")
RC_2=session("aRC_2")
RC_3=session("aRC_3")
RC_4=session("aRC_4")
RC_5=session("aRC_5")
RC_6=session("aRC_6")
RC_7=session("aRC_7")
RC_8=session("aRC_8")
RC_9=session("aRC_9")

AuthorizedRep=Replace(Request.Form("AuthorizedRep"),"'","''")
AuthorizedRepTitle=Replace(Request.Form("AuthorizedRepTitle"),"'","''")
SupportingExplanation=Replace(Request.Form("SupportingExplanation"),"'","''")
ApplicationDate=Request.Form("ApplicationDate")
OCN=Request.Form("OCN")
SET objConnection = server.createobject("ADODB.connection")
SET rstQry =server.createobject("ADODB.recordset")
ObjConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlQry = "Exec [Partial_MBI] " & NPA & "," & NXX & ", '" & OCN & "', " & EntityID & "," & UserID & ", '" & ApplicationDate & "'" & ", '" & AuthorizedRep & "'" & ",'" & AuthorizedRepTitle & "'" & ",'" & RC_0 & "'" & ",'" & RC_1 & "'" & ",'" & RC_2 & "'" & ",'" & RC_3 & "'" & ",'" & RC_4 & "'" & ",'" & RC_5 & "'" & ",'" & RC_6 & "'" & ",'" & RC_7 & "'" & ",'" & RC_8 & "'" & ",'" & RC_9 & "'" & ", '" & SupportingExplanation & "'"
Response.Write sqlQry
SET rstQry = objConnection.execute(sqlQry)
Response.Redirect "MBI_Partial_NPA_NXX_Confirm.asp"

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
