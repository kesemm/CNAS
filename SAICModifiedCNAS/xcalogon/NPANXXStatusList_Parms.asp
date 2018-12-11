<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA NXX Status List - Enter Parameters</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile$
'* Commit Date:   $Date$ (UTC)
'* Committed by:  $Author$
'* CVS Revision:  $Revision$
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"--></FORM>
<FORM action="NPANXXStatusList_Result.asp" method="get"><BIG><B>&nbsp;&nbsp;Enter Parameters to Generate NXX Status List for NPA</B></BIG><BR>
<BR>
&nbsp;&nbsp;<B>Enter an NPA:</B> <INPUT type="text" name="NPA" size="3" maxlength="3"><BR>
<BR>
&nbsp;&nbsp;<B>Select the Initial Primary Sort Order:</B><BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="NXX" checked="true">&nbsp;&nbsp;NXX<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="RateCenter">&nbsp;&nbsp;RateCenter<BR>
&nbsp;&nbsp;<INPUT type="radio" name="SortOrder" value="Company">&nbsp;&nbsp;Company<BR>
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
+"$RCSfile$\n"
+"$Revision$\n"
+"$Date$ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
