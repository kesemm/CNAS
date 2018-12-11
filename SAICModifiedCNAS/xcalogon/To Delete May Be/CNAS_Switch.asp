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

aSwitch = request.querystring("Switch")
SET objConnection1 = server.createobject("ADODB.connection")
SET rstSwitchQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLSwitchQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.OCN as OCN,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks,StrandedCodeComment FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.SwitchID)='" & aSwitch & "')) ORDER by NPA,NXX;"
SET rstSwitchQry = objConnection1.execute(SQLSwitchQry)

%> </p>

<p align="center"><strong>CNAS Switch Query for: <% = UCASE(aSwitch) %> </strong></p>
<b>

<p><br>
<table align="center" BORDER="1">
  <tr>

 <tr>
<th align="center">Tix</th>
    <th align="center">NPA</th>
    <th align="center">NXX</th>
    <th align="center">Status</th>
    <th align="center">Company</th>
    <th align="center">OCN</th>
    <th align="center">Rate Centre</th>
    <th align="center">Public Remarks</th>
    <th align="center">CNA Remarks</th>
    <th align="center">Stranded Code Comment</th>
    <td><br>

<% Do Until rstSwitchQry.EOF %>    </td>
  </tr>
  <tr align="center">
    <td nowrap><%= rstSwitchQry("Tix") %>
</td>

      <td nowrap><%= rstSwitchQry("NPA") %>
</td>
    <td nowrap><%= rstSwitchQry("NXX") %>
</td>
    <td nowrap><%= rstSwitchQry("COStatusDescription") %>
</td>
    <td nowrap><%= rstSwitchQry("EntityName") %>
</td>
    <td><%= rstSwitchQry("OCN") %>
</td>
    <td nowrap><%= rstSwitchQry("RateCenter") %>
</td>
    <td nowrap><%= rstSwitchQry("PublicRemarks") %>
</td>
    <td nowrap><%= rstSwitchQry("CNARemarks") %>
</td>
    <td nowrap><%= rstSwitchQry("StrandedCodeComment") %>
</td>

  </tr>
<% rstSwitchQry.moveNext
 loop %>
</table>


</table>
<% 
objConnection1.close
%>
</b>
</body>
</html>
