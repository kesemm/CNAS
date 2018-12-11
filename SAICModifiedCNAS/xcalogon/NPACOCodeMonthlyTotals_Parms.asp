<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA CO Code Totals - Enter Parameters</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPACOCodeMonthlyTotals_Parms.asp,v $
'* Commit Date:   $Date: 2011/01/04 19:41:18 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%


UserEntityType=session("UserEntityType")


 %>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"--></FORM>
<FORM action="NPACOCodeMonthlyTotals_Result.asp" method="get"><BIG><B>Enter Parameters to Generate CO Code Totals</B></BIG>
<P>NPA(s):
<INPUT type="text" name="NPA1" size="3" maxlength="3">
<INPUT type="text" name="NPA2" size="3" maxlength="3">
<INPUT type="text" name="NPA3" size="3" maxlength="3">
<INPUT type="text" name="NPA4" size="3" maxlength="3">
 (Enter 1, 2, 3 or 4 NPAs)<BR>
<BR>
List by Exchange <INPUT type="radio" name="ListBy" value="Exchange" checked="true"><BR>
List by Company <INPUT type="radio" name="ListBy" value="Company"><BR>
<BR>
<INPUT type="submit"> <INPUT type="reset"><BR>
<BR></P>
</FORM>
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
+"$RCSfile: NPACOCodeMonthlyTotals_Parms.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2011/01/04 19:41:18 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
