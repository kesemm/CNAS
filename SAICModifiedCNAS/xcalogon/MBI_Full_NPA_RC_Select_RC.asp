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
<title>MBI - Rate Centre Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: MBI_Full_NPA_RC_Select_RC.asp,v $
'* Commit Date:   $Date: 2017/07/31 15:46:21 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%>
<%


UserEntityType=session("UserEntityType")
session("aNPA")=Request.Form("NPA")
aNPA=session("aNPA")

 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<p><%

If aNPA=343 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=613 Order by RateCenter ASC"
ElseIf aNPA=581 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=418 Order by RateCenter ASC"
ElseIf aNPA=587 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA In (403, 780) Order by RateCenter ASC"
ElseIf aNPA=778 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA In (250, 604) Order by RateCenter ASC"
ElseIf aNPA=236 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA In (250, 604,778) Order by RateCenter ASC"
ElseIf aNPA=289 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=905 Order by RateCenter ASC"
ElseIf aNPA=226 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=519 Order by RateCenter ASC"
ElseIf aNPA=548 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=519 Order by RateCenter ASC"
ElseIf aNPA=438 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=514 Order by RateCenter ASC"
ElseIf aNPA=579 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=450 Order by RateCenter ASC"
ElseIf aNPA=249 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=705 Order by RateCenter ASC"
ElseIf aNPA=365 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=905 Order by RateCenter ASC"
ElseIf aNPA=437 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=416 Order by RateCenter ASC"
ElseIf aNPA=819 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA In (819,873) Order by RateCenter ASC"
ElseIf aNPA=825 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA In (403,780) Order by RateCenter ASC"
ElseIf aNPA=873 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA In (819,873) Order by RateCenter ASC"
ElseIf aNPA=431 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=204 Order by RateCenter ASC"
ElseIf aNPA=639 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=306 Order by RateCenter ASC"
ElseIf aNPA=782 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=902 Order by RateCenter ASC"
Else
sqlRC="Select Distinct RateCenter From xca_COCode where NPA='"&aNPA&"' Order by RateCenter ASC"
end if

SET objConnection1 = server.createobject("ADODB.connection")
SET rstRCQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstRCQry = objConnection1.execute(sqlRC)

%> </p>
<p align="center"><strong>Details for NPA : <% = UCASE(aNPA) %> </strong></p>
<p align="center">Select the <strong>Rate Centre</strong></p>
<b>
<% if (rstRCQry.EOF) then %><b></p>

<p>No records found for the NPA.</b> </p>
<% Else %>
<p><br>
<table align="center" BORDER="1">
  <tr>

<tr>
    <th align="center">&nbsp; Rate Centre &nbsp;</th>
</tr>

<% Do Until rstRCQry.EOF %>
  <tr align="left">
    <td><a HREF="MBI_Full_NPA_RC_Select_NXX.asp?RC=<%= rstRCQry("RateCenter") %>&NPA=<%= aNPA %>"><%= rstRCQry("RateCenter") %></a></td>
  </tr>
<% rstRCQry.moveNext
 loop %>
</table>
<%End If%>
<% 
objConnection1.close
%>
</b>
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
strAlertText="SAIC Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: MBI_Full_NPA_RC_Select_RC.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2017/07/31 15:46:21 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
