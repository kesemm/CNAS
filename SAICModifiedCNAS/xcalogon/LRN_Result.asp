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
<title>NPA / NXX Querry</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: LRN_Result.asp,v $
'* Commit Date:   $Date: 2014/04/21 16:49:14 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 

' 2014-04-21  Changed date format throughout when updating SQL code to return latest LERG record /ktwalsh
%>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" -->
</form>
<%

aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")

SET objConnectionLRN = server.createobject("ADODB.connection")
SET rstLRN = server.createobject("ADODB.recordset")
objConnectionLRN.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLRNQry = "Select NPA,NXX,LRN,LERGSTATUS.Description as [STATUS],CONVERT(CHAR(10),EFF_DATE,120) AS EFF_DATE,LRN_TYPE,SWITCH,LERG12.OCN,OCN_NAME,RC_NAME10" &_
" From LERG12" &_
" INNER JOIN LERGSTATUS ON LERGSTATUS.STATUS=LERG12.STATUS" &_
" Left Join LERG1 On LERG1.OCN=LERG12.OCN" &_
" WHERE [LERG12].NPA='" & aNPA & "' AND [LERG12].NXX='" & aNXX & "' AND EFF_DATE=(SELECT MAX(EFF_DATE) FROM LERG12 WHERE [LERG12].NPA='" & aNPA & "' AND [LERG12].NXX='" & aNXX & "')"
SET rstLRN = objConnectionLRN.execute(SQLLRNQry)


SET objConnectionLERG12 = server.createobject("ADODB.connection")
SET rstLergDateLERG12 = server.createobject("ADODB.recordset")
objConnectionLERG12.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLergDateLERG12 = "SELECT CONVERT(CHAR(19),LERG12DATE,120) AS LERG12DATE FROM LERG12DATE"
SET rstLergDateLERG12 = objConnectionLERG12.execute(SQLLergDateLERG12)
%>
<p align="center"><strong>LRN Query </strong></p>
<p align="center">Listing for NPA <strong><%= aNPA %> </strong>NXX <strong><%= aNXX %></strong> based on <%=rstlergDateLERG12("LERG12DATE") %> data</p>
<b>

<p><br>
<% if (rstLRN.EOF) then %><b></p>

<p>No LRN record found for NPA <%= aNPA %> NXX <%= aNXX %> in the LERG.</b> </p>

<% Else %>
<% Do Until rstLRN.EOF %>
<table align="center" BORDER="1">

  <tr align="left">
    <td>&nbsp;<b>LRN</b>&nbsp;</td>
    <td>&nbsp;<%= rstLRN("LRN") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>LRN Type</b>&nbsp;</td>
    <td>&nbsp;<%= rstLRN("LRN_TYPE") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Status</b>&nbsp;</td>
    <td>&nbsp;<%= rstLRN("STATUS") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Effective Date &nbsp; <br>&nbsp;(yyyy-mm-dd)</b>&nbsp;</td>
    <td>&nbsp;<%= rstLRN("EFF_DATE") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Company</b>&nbsp;</td>
    <td>&nbsp;<%= rstLRN("OCN_NAME") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>OCN</b>&nbsp;</td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstLRN("OCN") %> "><%=rstLRN("OCN") %> </a>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (Details)</b> &nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstLRN("SWITCH") %> "><%= rstLRN("SWITCH") %> </a> &nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>RateCentre</b>&nbsp;</td>
    <td>&nbsp;<%= rstLRN("RC_NAME10") %>&nbsp;</td>
  </tr>

<% rstLRN.moveNext
loop %>
</table>



<% end if
objConnectionLRN.close
objConnectionLERG12.close  %>
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
+"$RCSfile: LRN_Result.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2014/04/21 16:49:14 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
