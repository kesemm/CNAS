<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA CO Code Annual Totals</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPACOCodeAnnualTotals_Result.asp,v $
'* Commit Date:   $Date: 2016/02/12 15:52:16 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.1 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>NPA CO Code Annual Totals <%
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
SQLQueryText = "EXEC GetNPACOCodeAnnualTotals_Company '" & request.querystring("NPA1") & "','" & request.querystring("NPA2") & "','" & request.querystring("NPA3")& "','" & request.querystring("NPA4") & "';"
end if

if request.querystring("ListBy") = "Exchange" then
SQLQueryText = "EXEC GetNPACOCodeAnnualTotals_Exchange '" & request.querystring("NPA1") & "','" & request.querystring("NPA2") & "','" & request.querystring("NPA3")& "','" & request.querystring("NPA4") & "';"
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

' DERIVE THE JANUARY 1ST OF THIS YEAR
dateY1 = DateAdd("yyyy",(DateDiff("yyyy", 0, date) - 1), 2)

' SUBTRACT MONTHS TO GET BACK MONTHS
dateY2 = DateAdd("yyyy", -1, dateY1)
dateY3 = DateAdd("yyyy", -2, dateY1)
dateY4 = DateAdd("yyyy", -3, dateY1)
dateY5 = DateAdd("yyyy", -4, dateY1)
dateY6 = DateAdd("yyyy", -5, dateY1)

' GET A FORMATTED TITLE FOR EACH MONTH COLUMN USING SHORT MONTH NAME
Y1 = Left(MonthName(Month(dateY1)), 3) & " " & Day(dateY1) & ", " & Year(dateY1)
Y2 = Left(MonthName(Month(dateY2)), 3) & " " & Day(dateY2) & ", " & Year(dateY2)
Y3 = Left(MonthName(Month(dateY3)), 3) & " " & Day(dateY3) & ", " & Year(dateY3)
Y4 = Left(MonthName(Month(dateY4)), 3) & " " & Day(dateY4) & ", " & Year(dateY4)
Y5 = Left(MonthName(Month(dateY5)), 3) & " " & Day(dateY5) & ", " & Year(dateY5)
Y6 = Left(MonthName(Month(dateY6)), 3) & " " & Day(dateY6) & ", " & Year(dateY6)
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
<TD align="center"><B><% = Y1 %></B></TD>
<TD align="center"><B><% = Y2 %></B></TD>
<TD align="center"><B><% = Y3 %></B></TD>
<TD align="center"><B><% = Y4 %></B></TD>
<TD align="center"><B><% = Y5 %></B></TD>
<TD align="center"><B><% = Y6 %></B></TD>
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
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Y1")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Y2")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Y3")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Y4")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Y5")%> &nbsp;</B></TD>
<TD align="center"><B>&nbsp; <%=rstCodeSnapshotbyCompany("Y6")%> &nbsp;</B></TD>
</TR>
<%
Else

' USE THIS AS THE TABLE ROW FOR NON-BOLDED ROWS
' IT IS FOR THE COMPANY OR EXCHANGE ROWS
%>
<TR>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Name")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Current")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Y1")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Y2")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Y3")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Y4")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Y5")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstCodeSnapshotbyCompany("Y6")%> &nbsp;</TD>
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
strAlertText="Leidos Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: NPACOCodeAnnualTotals_Result.asp,v $\n"
+"$Revision: 1.1 $\n"
+"$Date: 2016/02/12 15:52:16 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
