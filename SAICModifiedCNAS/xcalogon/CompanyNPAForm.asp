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
aNPA = Request.Form("NPA")
aEntityName = Request.Form("EntityName")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<p align="center"><strong>CNAS Company NPA Query</strong></p>

<p align="center"><strong>for</strong></p>

<p align="center"><strong><%=aEntityName %> in NPA <%= aNPA %> </strong></p>
<b>

<table align="center" BORDER="1">
  <tr>
    <th align="center">Ticket Number</th>
    <th align="center">NXX</th>
    <th align="center">COStatusDescription</th>
    <th align="center">SwitchID</th>
    <th align="center">WireCenter</th>
    <th align="center">RateCenter</th>
    <th align="center">InServiceDateNXX</th>
  </tr>
<%
SET objConnectionEntity = server.createobject("ADODB.connection")
SET objConnectionCompany = server.createobject("ADODB.connection")

SET rstEntityQry =server.createobject("ADODB.recordset")
SET rstCompanyQry=server.createobject("ADODB.recordset")

objConnectionEntity.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
objConnectionCompany.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SQLEntityID=" SELECT EntityID FROM xca_Entity WHERE [EntityName]='" & aEntityName &"';"
SET rstEntityQry = objConnectionEntity.execute(SQLEntityID)

SQLCompanyNPAQry = "SELECT Tix,NXX,COStatusDescription,SwitchID,WireCenter,RateCenter,InServiceDate FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.EntityID)='" & rstEntityQry("EntityID") &"')) ORDER BY NXX;"
SET rstCompanyQry = objConnectionCompany.execute(SQLCompanyNPAQry)%>
<% Do Until rstCompanyQry.EOF %>
  <tr align="center">
    <td><%= rstCompanyQry("Tix") %>
</td>
    <td><%= rstCompanyQry("NXX") %>
</td>
    <td><%= rstCompanyQry("COStatusDescription") %>
</td>
    <td><%= rstCompanyQry("SwitchID") %>
</td>
    <td><%= rstCompanyQry("WireCenter") %>
</td>
    <td><%= rstCompanyQry("RateCenter") %>
</td>
    <td><%= rstCompanyQry("InServiceDate") %>
</td>
  </tr>
<% rstCompanyQry.moveNext 
    loop %>
<%objConnectionEntity.close %>
<%objConnectionCompany.close %>
</table>

<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
</b>
</body>
</html>
