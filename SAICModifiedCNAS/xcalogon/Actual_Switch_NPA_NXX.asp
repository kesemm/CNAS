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
<title>Actual Switch NPA NXX Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Actual_Switch_NPA_NXX.asp,v $
'* Commit Date:   $Date: 2006/05/17 15:46:35 $ (UTC)
'* Committed by:  $Author: SAIC-OTTAWA\browng $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%>
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
sqlSwitch="Select NPA,NXX,RC_NAME10,RC_NAME,Switch,LERG6.OCN,OCN_NAME " &_
"From LERG7 " &_
"Inner Join LERG6 " &_
"On LERG6.Switch=LERG7.Switch_ID " &_
"Inner Join LERG1 " &_
"On LERG6.OCN=LERG1.OCN " &_
"Where Actual_ID='" & aSwitch & "' " &_
"Order By Switch,NPA,NXX"

SET objConnection1 = server.createobject("ADODB.connection")
SET rstSwitchQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstSwitchQry = objConnection1.execute(sqlSwitch)

%> </p>

<p align="center"><strong>CO Codes using Actual Switch : <% = UCASE(aSwitch) %> </strong></p>
<b>
<% if (rstSwitchQry.EOF) then %><b></p>

<p>No records found for the Switch.</b> </p>
<% Else %>
<p><br>
<table align="center" BORDER="1">
  <tr>
<tr>
    <th align="center">&nbsp; NPA &nbsp;</th>
    <th align="center">&nbsp; NXX &nbsp;</th>
	<th align="center">&nbsp; POI &nbsp;</th>
    <th align="center">&nbsp; RC ABBR &nbsp;</th>
    <th align="center">&nbsp; RC Full &nbsp;</th>
    <th align="center">&nbsp; OCN &nbsp;</th>
    <th align="center">&nbsp; Company &nbsp;</th>
</tr>

<% Do Until rstSwitchQry.EOF %>
  <tr align="left">
    <td>&nbsp;<%= rstSwitchQry("NPA") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("NXX") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("Switch") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("RC_NAME10") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("RC_NAME") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("OCN") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("OCN_NAME") %>&nbsp;</td>
  </tr>
<% rstSwitchQry.moveNext
 loop %>
</table>
<%End If%>
<% 
objConnection1.close
%>
</b>
<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %><BR>
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
+"$RCSfile: Actual_Switch_NPA_NXX.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2006/05/17 15:46:35 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
