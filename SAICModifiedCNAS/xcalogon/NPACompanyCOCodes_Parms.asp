<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA Company CO Codes - Enter Parameters</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPACompanyCOCodes_Parms.asp,v $
'* Commit Date:   $Date: 2004/12/03 17:12:21 $ (UTC)
'* Committed by:  $Author: WalshKel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"--></FORM>
<FORM action="NPACompanyCOCodes_Result.asp" method="get"><BIG><B>&nbsp;&nbsp;Enter Parameters to Generate Company CO Codes List</B></BIG><BR>
<BR>
&nbsp;&nbsp;<B>Enter an NPA:</B> <INPUT type="text" name="NPA" size="3" maxlength="3"><BR>
<BR>
&nbsp;&nbsp;<B>Select the Initial Primary Sort Order:</B><BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="Company" checked="true">&nbsp;&nbsp;Company<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="OCN">&nbsp;&nbsp;OCN<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="Exchange">&nbsp;&nbsp;Exchange<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="CLLI">&nbsp;&nbsp;CLLI<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="Status">&nbsp;&nbsp;Status<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="NXX">&nbsp;&nbsp;NXX<BR>
<BR>
&nbsp;&nbsp;&nbsp;<INPUT type="submit"> <INPUT type="reset"><BR></FORM>
<BR>
<BR>
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
+"$RCSfile: NPACompanyCOCodes_Parms.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2004/12/03 17:12:21 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
