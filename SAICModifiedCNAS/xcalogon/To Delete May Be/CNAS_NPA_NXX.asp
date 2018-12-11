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
aNXX = request.querystring("NXX")
SET objConnection1 = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET objConnection2 = server.createobject("ADODB.connection")
SET RSTPart1Qry =server.createobject("ADODB.recordset")
objConnection2.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SQLNPANXXQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks,StrandedCodeComment FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SQLPart1Qry="Select Max(RequestedEffDate) As RequestedEffDate From xca_Part1 WHERE (((xca_Part1.NPA)='" & aNPA &"') AND ((xca_Part1.NXX1Preferred)='" & aNXX & "'));"
SET RSTNPANXXQry = objConnection1.execute(SQLNPANXXQry)
SET RSTPart1Qry=objConnection2.execute(SQLPart1Qry) %> </p>

<p align="center"><strong>CNAS NPA-NXX Query </strong></p>
<b>

<p><br>
<% if rstNPANXXQry.EOF then %><b></p>

<p>No record found for NPA <%= aNPA %> NXX <%= aNXX %>.</b> <% Elseif rstNPANXXQry("Status")="I" then %> </p>

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
  <tr align="center">
    <td><b>Stranded Code Comment</b></td>
    <td><%= rstNPANXXQry("StrandedCodeComment") %>
</td>
  </tr>

</table>

<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<% Elseif rstNPANXXQry("Status")="A" then %>

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
    <td><b>Requested Effective Date</b></td>
    <td><%= rstPart1Qry("RequestedEffDate") %>
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
  <tr align="center">
    <td><b>Stranded Code Comment</b></td>
    <td><%= rstNPANXXQry("StrandedCodeComment") %>
</td>
  </tr>

</table>

<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<%else%>

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
    <td><b>Public Remarks</b></td>
    <td><%= rstNPANXXQry("PublicRemarks") %>
</td>
  </tr>
  <tr align="center">
    <td><b>CNA Remarks</b></td>
    <td><%= rstNPANXXQry("CNARemarks") %>
</td>
  </tr>
  <tr align="center">
    <td><b>Stranded Code Comment</b></td>
    <td><%= rstNPANXXQry("StrandedCodeComment") %>
</td>
  </tr>

</table>

<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<% end if
objConnection1.close
objConnection2.close%>
</b>
</body>
</html>
