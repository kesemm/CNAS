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
<title>Alberta LIR Data</title>
<p align="center"><b>Alberta LIR Data</p></b>

<%
'****************************************************************************************
'* CVS File:      $RCSfile: UserLogon.asp,v $
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
<!-- #Include file="ADOVBS.INC" -->
</form>

<%
' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT
sqlQry = "Select * From AB_LIR Order By LIR,RateCenter"
 %>



<%
SET objConnection = server.createobject("ADODB.connection")
SET rstQry = server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstQry = objConnection.execute(sqlQry)
%>


<TABLE align="center" border="1">
<TD align="center"><B>LIR</B></TD>
<TD align="center"><B>Exchange</B></TD>

<p><br>
<br>
<% if rstQry.EOF then %><b>No records found.</b> <% ELSE %> </p>
<% Do Until rstQry.EOF %>
<tr align="left">
<td>&nbsp;<%=rstQry("LIR") %> </a> &nbsp;</td>
</td>
<td>&nbsp;<%= rstQry("RateCenter") %>&nbsp;</td>
</tr>
<% rstQry.moveNext
loop %>
</table>
<% end if
objConnection.close
%>
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
+"$RCSfile: UserLogon.asp,v $\n"
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
