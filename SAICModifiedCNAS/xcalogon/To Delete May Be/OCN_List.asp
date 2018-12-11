<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS OCN List</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: OCN_List.asp,v $
'* Commit Date:   $Date: 2004/12/03 17:39:17 $ (UTC)
'* Committed by:  $Author: WalshKel $
'* CVS Revision:  $Revision: 1.1 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>List of OCNs</BIG></B><BR></P>
<%
' SET UP THE CONNECTION AND RECORDSET

' SET THE FIRST PART OF THE QUERY TEXT

SQLQueryText = "SELECT [OCN], [Company] From xca_Entity_OCN_Web "

' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT

Select Case request.querystring("SortOrder")
Case "Company"
SQLQueryText = SQLQueryText & _
			"Order By Company"
Case "OCN"
SQLQueryText = SQLQueryText & _
			"Order By OCN"
Case Else
SQLQueryText = SQLQueryText & _
			"Order By Company"
End Select

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstOCN = server.createObject("ADODB.recordset")
rstOCN.activeConnection = objConnection
rstOCN.CursorLocation = adUseServer
rstOCN.CursorType = adOpenStatic
rstOCN.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS TO DETERMINE IF NPA ENTERED IS VALID

if rstOCN.RecordCount>0 then

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

<TD align="center"><B>&nbsp; OCN</B>
<a href="/xcalogon/OCN_List.asp?SortOrder=OCN">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Company</B>
<a href="/xcalogon/OCN_List.asp?SortOrder=Company">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstOCN.EOF
%>
<TR>
<TD align="center"> &nbsp;<%=rstOCN("OCN")%></a>&nbsp; </TD>
<TD align="center">&nbsp; <%=rstOCN("Company")%> &nbsp;</TD>
<%
' GET THE NEXT RECORD IN THE SET
rstOCN.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>A Database Error Has Occured</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstOCN.Close
objConnection.close
%>
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
+"$RCSFile$\n"
+"$Revision: 1.1 $\n"
+"$Date: 2004/12/03 17:39:17 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
