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
'****************************************************************************************
'* CVS File:      $RCSfile: Switch_Query.asp,v $
'* Commit Date:   $Date: 2006/05/17 16:01:03 $ (UTC)
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

<form ACTION="Switch_Result.asp" METHOD="get">
<table ALIGN="CENTER" BORDER="1" CELLPADING="3" CELLSPACING="3" WIDTH="50%">
<tr>
<td>Ratecentre</td>
<td><input TYPE="text" NAME="Ratecentre" SIZE="4" MAXLENGTH="4"></td>
</tr>
<tr>
<td>Province</td>
<td><input TYPE="text" NAME="Province" SIZE="2" MAXLENGTH="2"></td>
</tr>
<tr>
<td>Building</td>
<td><input TYPE="text" NAME="Building" SIZE="2" MAXLENGTH="2"></td>
</tr>
<tr>
<td>Equipment</td>
<td><input TYPE="text" NAME="Equipment" SIZE="3" MAXLENGTH="3"></td>
</tr>
</table>
<center>Enter one or more fields to search.  Use _ as a wildcard to match any string of one character.<br>
<br>
<center>(Example: DS_ for any switch like DS0, DS1, DS2 etc.)<br>
<br>
  <input TYPE="submit"><input TYPE="reset"><br>

</form>
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
+"$RCSfile: Switch_Query.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2006/05/17 16:01:03 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
