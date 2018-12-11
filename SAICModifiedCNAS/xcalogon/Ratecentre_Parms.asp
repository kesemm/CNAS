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
<title>Ratecentre Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Ratecentre_Parms.asp,v $
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


<form ACTION="Ratecentre_Select.asp" METHOD="get">
  <p>Enter a NPA:<input TYPE="text" NAME="NPA" SIZE="3" MAXLENGTH="3"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
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
+"$RCSfile: Ratecentre_Parms.asp,v $\n"
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
