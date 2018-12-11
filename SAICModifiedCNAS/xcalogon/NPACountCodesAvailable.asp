<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA Number of Available Codes</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPACountCodesAvailable.asp,v $
'* Commit Date:   $Date: 2006/10/04 12:27:10 $ (UTC)
'* Committed by:  $Author: browng $
'* CVS Revision:  $Revision: 1.5 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>NPA Count of Available Codes</BIG></B><BR></P>
<BR>
<p align="center"><strong>Temporarily Unavailable, Aging and 800/900 Codes are counted towards available CO Codes</strong></p>

<%
' SET UP THE CONNECTION AND RECORDSET

if request.querystring("SortOrder")="NPACount" then
SQLQueryText = "SELECT NPA, COUNT(*) AS [NPA-COUNT] FROM xca_COCode WHERE Status in ('s','B','4','L','2') GROUP BY NPA ORDER BY [NPA-COUNT]"
else
SQLQueryText = "SELECT NPA, COUNT(*) AS [NPA-COUNT] FROM xca_COCode WHERE Status in ('s','B','4','L','2') GROUP BY NPA ORDER BY NPA"
end if

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstNPACountAvailable = server.createObject("ADODB.recordset")
rstNPACountAvailable.activeConnection = objConnection
rstNPACountAvailable.CursorLocation = adUseServer
rstNPACountAvailable.CursorType = adOpenStatic
rstNPACountAvailable.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS

if rstNPACountAvailable.RecordCount>0 then

%>
<TABLE align="center" border="0" width="30%">
<TR>
<TD>
<font face="Arial" size="2">
<b>Instructions:</b><br>
- Click on the <image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/> to sort by that column<br>
</font>
</TD>
</TR>
</TABLE>
<br>

<TABLE align="center" border="1">
<TR>
<TD align="center"><B>&nbsp;&nbsp;NPA&nbsp; </B><a href="/xcalogon/NPACountCodesAvailable.asp">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a></TD>
<TD align="center"><B>&nbsp;&nbsp;Number of Available CO Codes&nbsp; </B><a href="/xcalogon/NPACountCodesAvailable.asp?SortOrder=NPACount">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstNPACountAvailable.EOF
%>
<TR>
<TD align="center">&nbsp; <%=rstNPACountAvailable("NPA")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPACountAvailable("NPA-COUNT")%> &nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstNPACountAvailable.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>Database Error</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstNPACountAvailable.Close
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
+"$RCSfile: NPACountCodesAvailable.asp,v $\n"
+"$Revision: 1.5 $\n"
+"$Date: 2006/10/04 12:27:10 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
