<%@ Language=VBScript %>
<%Response.Buffer = true
Response.Expires=0%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>CNAS Database Query</title>
<%UserEntityType=session("UserEntityType")
aEntityName = Request.Form("EntityName") %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<%
SET objConnectionEntity = server.createobject("ADODB.connection")
objConnectionEntity.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET objConnectionNPA = server.createobject("ADODB.connection")
objConnectionNPA.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET objConnectionCount = server.createobject("ADODB.connection")
objConnectionCount.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET RSTNPAQry =server.createobject("ADODB.recordset")
SQLEntityID=" SELECT EntityID FROM xca_Entity WHERE [EntityName]='" & aEntityName &"';"
SET RSTEntityID = objConnectionEntity.execute(SQLEntityID)
SQLNPAQry = "SELECT DISTINCT NPA FROM xca_COCode WHERE (((xca_COCode.EntityID)='" & RSTEntityID("EntityID") &"')) ORDER BY NPA;"
SET RSTNPAQry = objConnectionNPA.execute(SQLNPAQry) %>


<p align="center"><strong>CNAS Company Count By NPA</strong></p>

<p align="center"><strong>for</strong></p>

<p align="center"><strong><%=aEntityName %></strong></p>
<b>
<%If cint(RSTEntityID("EntityID")) = 1 Then %>
<p align="center"><strong>Temporarily Unavailable, Aging and 800/900 Codes not counted towards Admin as they are truely available</strong></p>
<% end if %>
<table align="center" BORDER="1">
  <tr>
    <th align="center">NPA</th>
    <th align="center">Number of CO Codes</th>
  </tr>
<%
Do Until RSTNPAQry.EOF
SET RSTCount =server.createobject("ADODB.recordset")
If cint(RSTEntityID("EntityID")) = 1 Then 
SQLCount = "SELECT Count (NXX) FROM xca_COCode WHERE ((xca_COCode.NPA=" & RSTNPAQry("NPA") &") AND (xca_COCode.EntityID=" & RSTEntityID("EntityID") &") AND (xca_COCode.Status <> 'L' AND xca_COCode.Status <> 'B' AND xca_COCode.Status <> '4'));"
Else
SQLCount = "SELECT Count (NXX) FROM xca_COCode WHERE ((xca_COCode.NPA=" & RSTNPAQry("NPA") &") AND (xca_COCode.EntityID=" & RSTEntityID("EntityID") &"));"
End IF
SET RSTCount = objConnectionCount.execute(SQLCount)%>

  <tr align="center">
    <td><%= RSTNPAQry("NPA") %>
</td>
    <td><%= RSTCount.Fields(0)%>
</td>
  </tr>
<% RSTNPAQry.moveNext 
    loop %>
<%objConnectionCount.close %>
<%objConnectionNPA.close %>
<%objConnectionEntity.close %>
</table>
</b>
</body>
</html>
