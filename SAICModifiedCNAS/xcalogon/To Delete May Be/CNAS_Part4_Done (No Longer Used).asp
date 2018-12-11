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
<title>CNAS Part 4</title>
<%UserEntityType=session("UserEntityType")%>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<p><%
aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
aDate=Request.querystring("Date")
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SET rstNPANXXQry = objConnection.execute(SQLNPANXXQry) %> </p>

<p align="center"><strong>CNAS Part 4</strong></p>

<p><% if rstNPANXXQry("Status")<>"A" then %><b></p>

<p align="center">Sorry that code is not available to place In-Service. </p>

<p><% elseif rstNPANXXQry("Tix")<>"999999999" then %><b></p>

<p align="center">You need to use the 'normal' method for NPA <%= aNPA %> NXX <%= aNXX %>. </p>

<p><br>
</p>

<p align="center">Logon on as a User from <% =rstNPANXXQry("EntityName")%>.</b> </p>
<%ELSE%>
<b><% objConnection.close %>
<%
aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
aDate=Request.querystring("Date")
aDate=CDate(aDate)
SET objConnection = server.createobject("ADODB.connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
aSQL = "UPDATE xca_COCode Set Status='I', InServiceDate='"& aDate & "' WHERE (xca_COCode.NPA='" & aNPA & "' AND xca_COCode.NXX='" &  aNXX & "');" 
objConnection.execute(aSQL)
objConnection.close
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXQry = "SELECT Tix,NPA,NXX,COStatusDescription,EntityName,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SET rstNPANXXQry = objConnection.execute(SQLNPANXXQry) %>


<p align="center"><strong>Done</strong></p>

<table align="center" BORDER="1">
  <tr align="center">
    <td><b>Ticket Number</b></td>
    <td><%= rstNPANXXQry("Tix") %>
</td>
  </tr>
  <tr align="center">
    <td><b>NPA</b></td>
    <td><%= rstNPANXXQry("NPA") %>
</td>
  </tr>
  <tr align="center">
    <td><b>NXX</b></td>
    <td><%= rstNPANXXQry("NXX") %>
</td>
  </tr>
  <tr align="center">
    <td><b>Status</b></td>
    <td><%= rstNPANXXQry("COStatusDescription") %>
</td>
  </tr>
  <tr align="center">
    <td><b>Company</b></td>
    <td><%= rstNPANXXQry("EntityName") %>
</td>
  </tr>
  <tr align="center">
    <td><b>OCN</b></td>
    <td><%= rstNPANXXQry("OCN1") %>
</td>
  </tr>
  <tr align="center">
    <td><b>SwitchID</b></td>
    <td><%= rstNPANXXQry("SwitchID") %>
</td>
  </tr>
  <tr align="center">
    <td><b>WireCentre</b></td>
    <td><%= rstNPANXXQry("WireCenter") %>
</td>
  </tr>
  <tr align="center">
    <td><b>RateCentre</b></td>
    <td><%= rstNPANXXQry("RateCenter") %>
</td>
  </tr>
  <tr align="center">
    <td><b>InService Date</b></td>
    <td><%= rstNPANXXQry("InServiceDate") %>
</td>
  </tr>
  <tr align="center">
    <td><b>Public Remarks</b></td>
    <td><%= rstNPANXXQry("PublicRemarks") %>
</td>
  </tr>
  <tr align="center">
    <td><b>CNA Remarks</b></td>
    <td><%= rstNPANXXQry("CNARemarks") %>
</td>
  </tr>
</table>
<% end if%>
<% objConnection.close %>
</b></b>
</body>
</html>
