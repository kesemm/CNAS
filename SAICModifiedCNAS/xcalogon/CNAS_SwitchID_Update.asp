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
<title>CNAS SwitchID Update</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: CNAS_SwitchID_Update.asp,v $
'* Commit Date:   $Date: 2014/04/17 16:44:14 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
<%UserEntityType=session("UserEntityType")%>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<p><%
aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
aCLLI=Request.querystring("CLLI")
aRemarks=Request.querystring("Remarks")
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXQry = "SELECT Tix,NPA,NXX,Status,CNARemarks FROM xca_COCode WHERE (xca_COCode.NPA='" & aNPA &"' AND xca_COCode.NXX='" & aNXX & "');"
SET rstNPANXXQry = objConnection.execute(SQLNPANXXQry)
aCNARemarks=trim(rstNPANXXQry("CNARemarks"))%> </p>

<p align="center"><strong>CNAS SwitchID Update </strong></p>

<p><% if rstNPANXXQry("Status")="S" then %><b></p>

<p align="center">Sorry that code is not available to change the SwitchID. </p>

<p><% elseif rstNPANXXQry("Tix")="999999999" then %><b></p>

<p><%
aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
aCLLI=Request.querystring("CLLI")
' 2014-04-17 Changed remarks order on insert and add semicolon. /ktwalsh
IF LEN(aCNARemarks) > 0 THEN
	' Existing remarks, put new remark then separator then old remarks
	aRemarks=Request.querystring("Remarks") & "; " & aCNARemarks
ELSE
	' Just use new remark
	aRemarks=Request.querystring("Remarks")
END IF
SET objConnection = server.createobject("ADODB.connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
aSQL = "UPDATE xca_COCode Set SwitchID='"& aCLLI & "', CNARemarks='"& aRemarks &"' WHERE (xca_COCode.NPA='" & aNPA & "' AND xca_COCode.NXX='" &  aNXX & "');" 
objConnection.execute(aSQL)
objConnection.close
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXQry = "SELECT Tix,NPA,NXX,COStatusDescription,EntityName,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,PublicRemarks,CNARemarks FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SET rstNPANXXQry = objConnection.execute(SQLNPANXXQry) %> </p>

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

<p><% elseif rstNPANXXQry("Tix") <>"999999999" then %><b></p>

<p><%
aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
aCLLI=Request.querystring("CLLI")
' 2014-04-17 Changed remarks order on insert and add semicolon. /ktwalsh
IF LEN(aCNARemarks) > 0 THEN
	' Existing remarks, put new remark then separator then old remarks
	aRemarks=Request.querystring("Remarks") & "; " & aCNARemarks
ELSE
	' Just use new remark
	aRemarks=Request.querystring("Remarks")
END IF
SET objConnection = server.createobject("ADODB.connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
aSQL = "UPDATE xca_COCode Set SwitchID='"& aCLLI & "', CNARemarks='"& aRemarks &"' WHERE (xca_COCode.NPA='" & aNPA & "' AND xca_COCode.NXX='" &  aNXX & "');" 
objConnection.execute(aSQL)
objConnection.close
aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
aCLLI=Request.querystring("CLLI")
SET objConnection = server.createobject("ADODB.connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
aSQL = "UPDATE xca_Part1 Set SwitchID='"& aCLLI & "' WHERE (xca_Part1.NPA='" & aNPA & "' AND xca_Part1.NXX1preferred='" &  aNXX & "');" 
objConnection.execute(aSQL)
objConnection.close
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXQry = "SELECT Tix,NPA,NXX,COStatusDescription,EntityName,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,PublicRemarks,CNARemarks FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SET rstNPANXXQry = objConnection.execute(SQLNPANXXQry) %> </p>

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
</b></b></b>
<%
' THIS IS THE VERSION CONTROL INFORMATION BLOCK
' ---------------------------------------------
'
' Subdued input text box, that when clicked will make an alert with CVS Info
'
%>
<br><br>
<INPUT TYPE="TEXT" 
       STYLE="border: none; background-color: #D7C7A4; font: 7pt Arial; color: gray; width: 200px" 
       ONCLICK="VerInfo()" VALUE="CNAS Version Control Information"
       READONLY>
<SCRIPT language="JavaScript">
function VerInfo()
{
var strAlertText
strAlertText="Leidos Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: CNAS_SwitchID_Update.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2014/04/17 16:44:14 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
