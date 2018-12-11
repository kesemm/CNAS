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
<title>Rate Centre Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Ratecentre_Select.asp,v $
'* Commit Date:   $Date: 2006/05/17 15:54:01 $ (UTC)
'* Committed by:  $Author: SAIC-OTTAWA\browng $
'* CVS Revision:  $Revision: 1.3 $
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

aNPA = request.querystring("NPA")
sqlRC="Select CNA_Rate_Centre,BIRRDS_Rate_Centre,RC_ABBR,RC_PROV,Major_V,Major_H " &_
"From CNA_RC_TO_BIRRDS_RC " &_
"Left Join BIRRDS_Rate_Centre " &_
"On CNA_RC_TO_BIRRDS_RC.BIRRDS_Rate_Centre=BIRRDS_Rate_Centre.RC_FULL And " &_
"CNA_RC_TO_BIRRDS_RC.NPA=BIRRDS_Rate_Centre.NPA " &_
"Where CNA_RC_TO_BIRRDS_RC.NPA=" & aNPA & " " &_
"Order By CNA_Rate_Centre"

SET objConnection1 = server.createobject("ADODB.connection")
SET rstRCQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstRCQry = objConnection1.execute(sqlRC)

%> </p>

<p align="center"><strong>Details for NPA : <% = UCASE(aNPA) %> </strong></p>
<p align="center">Select the <strong>BIRRDS Abbreviation</strong> for a list of Localities</p>
<b>
<% if (rstRCQry.EOF) then %><b></p>

<p>No records found for the NPA.</b> </p>
<% Else %>
<p><br>
<table align="center" BORDER="1">
  <tr>

<tr>
    <th align="center">&nbsp; CNA Rate Centre &nbsp;</th>
    <th align="center">&nbsp; BIRRDS Full Name &nbsp;</th>
    <th align="center">&nbsp; BIRRDS ABBR &nbsp;</th>
    <th align="center">&nbsp; BIRRDS Prov &nbsp;</th>
    <th align="center">&nbsp; Major V &nbsp;</th>
    <th align="center">&nbsp; Major H &nbsp;</th>
</tr>

<% Do Until rstRCQry.EOF %>
  <tr align="left">
    <td>&nbsp;<%= rstRCQry("CNA_Rate_Centre") %>&nbsp;</td>
	<td>&nbsp;<%= rstRCQry("BIRRDS_Rate_Centre") %>&nbsp;</td>
    <td>&nbsp;<a HREF="Localities.asp?RC=<%= rstRCQry("RC_ABBR") %>&PROV=<%= rstRCQry("RC_PROV") %>&FULLRC=<%= rstRCQry("BIRRDS_Rate_Centre") %> "><%= rstRCQry("RC_ABBR") %> </a>&nbsp;</td>
	<td>&nbsp;<%= rstRCQry("RC_PROV") %>&nbsp;</td>
	<td>&nbsp;<%= rstRCQry("Major_V") %>&nbsp;</td>
	<td>&nbsp;<%= rstRCQry("Major_H") %>&nbsp;</td>
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
+"$RCSfile: Ratecentre_Select.asp,v $\n"
+"$Revision: 1.3 $\n"
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
