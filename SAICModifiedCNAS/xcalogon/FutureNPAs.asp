<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS 'Being Recovered' Codes</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Being_Recovered.asp,v $
'* Commit Date:   $Date: 2004/12/03 17:12:21 $ (UTC)
'* Committed by:  $Author: WalshKel $
'* CVS Revision:  $Revision: 1.4 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>List of Future NPAs</BIG></B><BR></P>
<BR>
<%
' SET UP THE CONNECTION AND RECORDSET

SQLQueryText = "SELECT Distinct NXX,Publicremarks From xca_COCode Where Status='F' order By NXX"
SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstFutureNPAs = server.createObject("ADODB.recordset")
rstFutureNPAs.activeConnection = objConnection
rstFutureNPAs.CursorLocation = adUseServer
rstFutureNPAs.CursorType = adOpenStatic
rstFutureNPAs.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS

if rstFutureNPAs.RecordCount>0 then

%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>NXX</B></TD>
<TD align="center"><B>PublicRemarks</B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstFutureNPAs.EOF
%>
<TD align="center">&nbsp; <%=rstFutureNPAs("NXX")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstFutureNPAs("PublicRemarks")%> &nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstFutureNPAs.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>No Future NPA Codes, or Database Error</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstFutureNPAs.Close
objConnection.close
%><BR>
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
+"$RCSfile: Being_Recovered.asp,v $\n"
+"$Revision: 1.4 $\n"
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
