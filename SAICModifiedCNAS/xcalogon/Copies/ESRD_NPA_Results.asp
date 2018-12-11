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

<p><%

aNPA = request.querystring("NPA")
aESRD = request.querystring("ESRD")
SET objConnection = server.createobject("ADODB.connection")
SET rstNPAESRDQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPAESRDQry = "SELECT Tix,NPA,ESRD,COStatusDescription,EntityName,xca_ESRD.OCN,PublicRemarks,ReserveID,CNARemarks FROM xca_ESRD Left Join xca_status_codes ON xca_ESRD.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_ESRD.EntityID=xca_Entity.EntityID WHERE xca_ESRD.NPA='" & aNPA &"' Order By ESRD;"
SET RSTNPAESRDQry = objConnection.execute(SQLNPAESRDQry)
%> </p>

<p align="center"><strong>CNAS NPA-ESRD Query </strong></p>
<b>
<% if RSTNPAESRDQry.EOF then %><b><p>No record found for:  <% = (aNPA) %> 
<%else%></p>

<p align="center"><strong>ESRDs Assigned to <% = (aNPA) %> </strong></p></td>
<b>

<p><br>
<table align="center" BORDER="1">
  <tr>

 <tr>
    <th align="center">&nbsp; Tix &nbsp;</th>
    <th align="center">&nbsp; NPA &nbsp;</th>
    <th align="center">&nbsp; ESRD &nbsp;</th>
    <th align="center">&nbsp; Status &nbsp;</th>
    <th align="center">&nbsp; Company &nbsp;</th>
    <th align="center">&nbsp; OCN &nbsp;</th>
	<th align="center">&nbsp; Public Remarks &nbsp;</th>
	<th align="center">&nbsp; ReserveID &nbsp;</th>
	<th align="center">&nbsp; CNA Remarks &nbsp;</th>
  </tr>

<% Do Until RSTNPAESRDQry.EOF %>

  <tr align="center">
    <td>&nbsp;<%= RSTNPAESRDQry("Tix") %>&nbsp;</td>
    <td>&nbsp;<%= RSTNPAESRDQry("NPA") %>&nbsp;</td>
    <td nowrap>&nbsp;<%= RSTNPAESRDQry("ESRD") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPAESRDQry("COStatusDescription") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPAESRDQry("EntityName") %>&nbsp;</td>
	<td>&nbsp;<%= RSTNPAESRDQry("OCN") %>&nbsp;</td>
	<td>&nbsp;<%= RSTNPAESRDQry("PublicRemarks") %>&nbsp;</td>
	<td>&nbsp;<%= RSTNPAESRDQry("ReserveID") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPAESRDQry("CNARemarks") %>&nbsp;</td>
  </tr>
<% RSTNPAESRDQry.moveNext
 loop %>
</p>
</table>



<p>Note: A ticket number of 999999999 implies a grandfathered MBI.</p>
 <%end if%>
</p>
<%
objConnection.close %>
</b>
</body>
</html>
