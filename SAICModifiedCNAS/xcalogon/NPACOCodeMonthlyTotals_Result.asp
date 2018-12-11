<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA CO Code Totals</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPACOCodeMonthlyTotals_Result.asp,v $
'* Commit Date:   $Date: 2011/01/04 19:31:09 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>NPA CO Code Monthly Totals <%
' PUT THE CORRECT LIST BY WORDS IN THE TITLE (COMPANY OR EXCHANGE)

if request.querystring("ListBy") = "Company" then
%> (By Company) <%
Else
%> (By Exchange) <%
End If

' PUT THE NPAs TOGETHER IN A STRING FOR DISPLAY
NPADisplayString = request.querystring("NPA1") & "  " &  request.querystring("NPA2") & "  " &  request.querystring("NPA3") & "  " &  request.querystring("NPA4") 

%> <BR> NPA selection:&nbsp;&nbsp; <% = NPADisplayString %></BIG></B><BR></P>
<BR>
<%

' SET THE QUERY TEXT TO RUN THE STORED PROCEDURE FOR COMPANY OR EXCHANGE

if request.querystring("ListBy") = "Company" then
SQLQueryText = "EXEC GetNPACOCodeMonthlyTotals_Company '" & request.querystring("NPA1") & "','" & request.querystring("NPA2") & "','" & request.querystring("NPA3")& "','" & request.querystring("NPA4") & "';"
end if

if request.querystring("ListBy") = "Exchange" then
SQLQueryText = "EXEC GetNPACOCodeMonthlyTotals_Exchange '" & request.querystring("NPA1") & "','" & request.querystring("NPA2") & "','" & request.querystring("NPA3")& "','" & request.querystring("NPA4") & "';"
end if
%><%
' SET UP THE CONNECTION AND RECORDSET

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstCodeSnapshotbyCompany = server.createObject("ADODB.recordset")
rstCodeSnapshotbyCompany.activeConnection = objConnection
rstCodeSnapshotbyCompany.CursorLocation = adUseServer
rstCodeSnapshotbyCompany.CursorType = adOpenStatic
rstCodeSnapshotbyCompany.open SQLQueryText, objConnection
%><%
' DERIVE THE 1ST DAY OF THIS MONTH

dateM1 = DateAdd("d", ((Day(date) -1) * -1), date)

' SUBTRACT MONTHS TO GET BACK MONTHS
dateM2 = DateAdd("m", -1, dateM1)
dateM3 = DateAdd("m", -2, dateM1)
dateM4 = DateAdd("m", -3, dateM1)
dateM5 = DateAdd("m", -4, dateM1)
dateM6 = DateAdd("m", -5, dateM1)

' GET A FORMATTED TITLE FOR EACH MONTH COLUMN USING SHORT MONTH NAME
M1 = Left(MonthName(Month(dateM1)), 3) & " " & Day(dateM1) & ", " & Year(dateM1)
M2 = Left(MonthName(Month(dateM2)), 3) & " " & Day(dateM2) & ", " & Year(dateM2)
M3 = Left(MonthName(Month(dateM3)), 3) & " " & Day(dateM3) & ", " & Year(dateM3)
M4 = Left(MonthName(Month(dateM4)), 3) & " " & Day(dateM4) & ", " & Year(dateM4)
M5 = Left(MonthName(Month(dateM5)), 3) & " " & Day(dateM5) & ", " & Year(dateM5)
M6 = Left(MonthName(Month(dateM6)), 3) & " " & Day(dateM6) & ", " & Year(dateM6)
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>&nbsp; <%
' PUT IN THE CORRECT COLUMN NAME FOR COMPANY OR EXCHANGE
if request.querystring("ListBy") = "Company" then
%> Company <%
Else
%> Exchange <%
End If
%></B></TD>
<TD align="center"><B>Current</B></TD>
<TD align="center"><B><% = M1 %></B></TD>
<TD align="center"><B><% = M2 %></B></TD>
<TD align="center"><B><% = M3 %></B></TD>
<TD align="center"><B><% = M4 %></B></TD>
<TD align="center"><B><% = M5 %></B></TD>
<TD align="center"><B><% = M6 %></B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstCodeSnapshotbyCompany.EOF
%><%
' USE THIS AS THE TABLE ROW SO THE LINE WILL BE BOLD
' IT IS FOR SUB TOTALS, TOTALS AND AVAILABLE ROWS

if rstCodeSnapshotbyCompany("SortOrder")=20 OR rstCodeSnapshotbyCompany("SortOrder")=60 OR rstCodeSnapshotbyCompany("SortOrder")=70 Then
%>
<TR>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Name")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Current")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("M1")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("M2")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("M3")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("M4")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("M5")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("M6")%> &nbsp;</B></TD>
</TR>
<%
Else

' USE THIS AS THE TABLE ROW FOR NON-BOLDED ROWS
' IT IS FOR THE COMPANY OR EXCHANGE ROWS
%>
<TR>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Name")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Current")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("M1")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("M2")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("M3")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("M4")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("M5")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("M6")%> &nbsp;</TD>
</TR>
<%
End If
%><%
' GET THE NEXT RECORD IN THE SET
rstCodeSnapshotbyCompany.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
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
+"$RCSfile: NPACOCodeMonthlyTotals_Result.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2011/01/04 19:31:09 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
