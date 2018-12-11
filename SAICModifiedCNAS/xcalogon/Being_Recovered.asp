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
<P align="center"><B><BIG>List of Codes Being Recovered</BIG></B><BR></P>
<BR>
<%
' SET UP THE CONNECTION AND RECORDSET

SQLQueryText = "SELECT Tix, NPA ,NXX, EntityName From xca_COCode " & _
                "JOIN xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID " & _
                "WHERE Status='5' " & _
                "ORDER BY EntityName, NPA, NXX"

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstBeing_Recovered = server.createObject("ADODB.recordset")
rstBeing_Recovered.activeConnection = objConnection
rstBeing_Recovered.CursorLocation = adUseServer
rstBeing_Recovered.CursorType = adOpenStatic
rstBeing_Recovered.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS

if rstBeing_Recovered.RecordCount>0 then

%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>Tix</B></TD>
<TD align="center"><B>NPA</B></TD>
<TD align="center"><B>NXX</B></TD>
<TD align="center"><B>Company</B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstBeing_Recovered.EOF
%>
<TR><%
if rstBeing_Recovered("Tix")="999999999" then
%>
<TD align="center"><SMALL><SMALL>&nbsp; From Data Load &nbsp;</SMALL></SMALL></TD>
<%
else
%>
<TD align="center">&nbsp; <%=rstBeing_Recovered("Tix")%> &nbsp;</TD>
<%
end if
%>
<TD align="center">&nbsp; <%=rstBeing_Recovered("NPA")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstBeing_Recovered("NXX")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstBeing_Recovered("EntityName")%> &nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstBeing_Recovered.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>No 'Being Recovered' Codes, or Database Error</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstBeing_Recovered.Close
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
