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
<title>Localities Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Localities.asp,v $
'* Commit Date:   $Date: 2006/05/17 15:54:01 $ (UTC)
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

aRC = request.querystring("RC")
aPROV = request.querystring("PROV")
aFULLRC = request.querystring("FULLRC")

sqlLOC="Select LOC_ABBR,LOC_PROV,LOC_FULL " &_
"From BIRRDS_Locality " &_
"Where RC_ABBR='" & aRC & "' " &_
"And LOC_PROV='" & aPROV & "' " &_
"Order By LOC_FULL"

SET objConnection1 = server.createobject("ADODB.connection")
SET rstLOCQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstLOCQry = objConnection1.execute(sqlLOC)

%> </p>

<p align="center"><strong>Locality details for the Rate Centre of: <% = UCASE(aFULLRC) %></strong></p>
<p align="center">with an abbreviation of: (<% = UCASE(aRC) %>)</p>
<% if (rstLOCQry.EOF) then %><b></p>

<p>No records found for the NPA.</b> </p>
<% Else %>
<p><br>
<table align="center" BORDER="1">
  <tr>

<tr>
    <th align="center">&nbsp; Full Name &nbsp;</th>
    <th align="center">&nbsp; ABBR &nbsp;</th>
    <th align="center">&nbsp; Prov &nbsp;</th>
</tr>

<% Do Until rstLOCQry.EOF %>
  <tr align="left">
    <td>&nbsp;<%= rstLOCQry("LOC_FULL") %>&nbsp;</td>
	<td>&nbsp;<%= rstLOCQry("LOC_ABBR") %>&nbsp;</td>
 	<td>&nbsp;<%= rstLOCQry("LOC_PROV") %>&nbsp;</td>
  </tr>
<% rstLOCQry.moveNext
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
+"$RCSfile: Localities.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2006/05/17 15:54:01 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
