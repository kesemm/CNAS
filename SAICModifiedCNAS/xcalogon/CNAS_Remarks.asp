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
'* CVS File:      $RCSfile: CNAS_Remarks.asp,v $
'* Commit Date:   $Date: 2014/04/17 16:44:14 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name$ (Version/Build)
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

<form ACTION="CNAS_Remarks_Update.asp" METHOD="get">
  <p>Enter a NPA:<input TYPE="text" NAME="NPA" SIZE="3" MAXLENGTH="3"><br>
  Enter a NXX:<input TYPE="text" NAME="NXX" SIZE="3" MAXLENGTH="3"><br>
<%
' 2014-04-17 Updated the recommended date format to the superior ISO machine readable format!. /ktwalsh
%>
  Enter the CNA Remarks (eg. Updated Eff Date in BIRRDS to 2014-12-31 on 2014-08-15):<input
  TYPE="text" NAME="Remarks" SIZE="60" MAXLENGTH="60"><br>
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
strAlertText="Leidos Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: CNAS_Remarks.asp,v $\n"
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
