<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>All Open CNA Tasks</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: OpenCNATasks_all.asp,v $
'* Commit Date:   $Date: 2010/08/11 13:39:20 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>All Open CNA Tasks</BIG></B><BR></P>

<%
' SET UP THE CONNECTION AND RECORDSET

SQLQueryText ="Select [id],fname As [Assigned To],cname As [Category],title,Convert(varchar(10),due_date,120) as DueDate " & _
"from problems Inner join status On problems.status=status.status_id " & _
"Inner Join tblUsers " & _
"On problems.assigned_to=tblUsers.sid " & _
"Inner Join categories " & _
"On problems.category=categories.category_id " & _
"Where sname='Open'" & _
"Order By due_date" 

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=CNATracking;UID=CNATracking;PWD=ticket"
SET rstCNATrack = server.createObject("ADODB.recordset")
rstCNATrack.activeConnection = objConnection
rstCNATrack.CursorLocation = adUseServer
rstCNATrack.CursorType = adOpenStatic
rstCNATrack.open SQLQueryText, objConnection
%>

<TABLE align="center" border="0" width="95%">
<TR>
<TD align="left"><B><%=rstCNATrack.RecordCount%> Tasks Total</B></TD>
<TD align="right"><a href="/xcalogon/OpenCNATasks_some.asp">Show tasks due within 14 days only</a></B></TD>
</TR>

<%
if rstCNATrack.RecordCount>0 then
%>

<TABLE align="center" border="1" width="97%">
<TR>
<TD align="center"><B>ID</B></TD>
<TD align="center"><B>Assigned To</B></TD>
<TD align="center"><B>Category</B></TD>
<TD align="center"><B>Title</B></TD>
<TD align="center"><B>Due Date<BR>(yyyy-mm-dd)</B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstCNATrack.EOF
%>
<%
If UserEntityType = "a" and UserUserType= "a" then
Session("lhd_ext_uid")=Session("UserUserLogon")
End If
%>
<TR>
 <TD align="center"><a href="xca_DetailAdminTrackingPost.asp?aid=<%=rstCNATrack("ID")%>" target= "blank"><%=rstCNATrack("ID")%></TD>
<TD align="center"><%=rstCNATrack("Assigned To")%></TD>
<TD align="center"><%=rstCNATrack("Category")%>&nbsp;</TD>
<TD align="center"><%=rstCNATrack("Title")%>&nbsp;</TD>
<TD align="center"><%=rstCNATrack("DueDate")%>&nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstCNATrack.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>&nbsp;There are no open CNA Tasks&nbsp;</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstCNATrack.Close
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
+"$RCSfile: OpenCNATasks_all.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2010/08/11 13:39:20 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
