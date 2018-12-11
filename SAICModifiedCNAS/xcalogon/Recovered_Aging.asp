<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS Recovered / Aging Codes</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Recovered_Aging.asp,v $
'* Commit Date:   $Date: 2004/12/03 17:12:21 $ (UTC)
'* Committed by:  $Author: WalshKel $
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
<P align="center"><B><BIG>List of Recovered/Aging Codes</BIG></B><BR></P>
<BR>
<%
' SET UP THE CONNECTION AND RECORDSET

SQLQueryText = "select NPA, NXX, " & _
                "convert(char(10), EarliestInServiceDate, 103) as [EISD], " & _
                "CNARemarks  " & _
                "from xca_COCode  where Status='4' order by EarliestInServiceDate"

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstRecovered_Aging = server.createObject("ADODB.recordset")
rstRecovered_Aging.activeConnection = objConnection
rstRecovered_Aging.CursorLocation = adUseServer
rstRecovered_Aging.CursorType = adOpenStatic
rstRecovered_Aging.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS

if rstRecovered_Aging.RecordCount>0 then

%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>NPA</B></TD>
<TD align="center"><B>NXX</B></TD>
<TD align="center"><B><SMALL>&nbsp;&nbsp;Earliest In Service Date&nbsp;&nbsp;<BR>
dd/mm/yyyy</SMALL></B></TD>
<TD align="center"><B>CNA Remarks</B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstRecovered_Aging.EOF
%>
<TR>
<TD align="center">&nbsp; <%=rstRecovered_Aging("NPA")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstRecovered_Aging("NXX")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstRecovered_Aging("EISD")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstRecovered_Aging("CNARemarks")%> &nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstRecovered_Aging.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>No Recovered/Aging Codes, or Database Error</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstRecovered_Aging.Close
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
+"$RCSfile: Recovered_Aging.asp,v $\n"
+"$Revision: 1.5 $\n"
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
