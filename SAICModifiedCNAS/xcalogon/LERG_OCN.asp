<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<%SET objConnectionLERG1 = server.createobject("ADODB.connection")
SET rstLergDateLERG1 = server.createobject("ADODB.recordset")
'objConnectionLERG1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
objConnectionLERG1.open "DSN=cnasadmin;SERVER=cnac-db.database.windows.net;UID=SysAdmin;PWD=DbAccess460;APP=Microsoft Development Environment;WSID=CNAS.DOMIAN.CA;DATABASE=XCA_DB1;QueryLogFile=Yes"
SQLLergDateLERG1 = "SELECT * FROM LERG1DATE"
SET rstLergDateLERG1 = objConnectionLERG1.execute(SQLLergDateLERG1)%>

<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>Canadian OCNs</title>
<p align="center"><b>Canadian OCNs based on <%=rstlergDateLERG1("LERG1DATE") %>  BIRRDS data </p></b>
<p align="center">Click on the OCN for a complete listing of contact information</p>

<%
'****************************************************************************************
'* CVS File:      $RCSfile: LERG_OCN.asp,v $
'* Commit Date:   $Date: 2006/05/17 16:01:03 $ (UTC)
'* Committed by:  $Author: SAIC-OTTAWA\browng $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" -->
</form>

<%
' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT
sqlOCNQry = "SELECT Distinct [LERG1].OCN, BILL_RAO As RAO, OCN_NAME, COMPANY" &_
" FROM [LERG1]" &_
" Left Join [LERG6]" &_
" On [LERG1].OCN=[LERG6].OCN "

Select Case request.querystring("SortOrder")

Case "Company"
SQLOCNQry = SQLOCNQry & _
			"Order by Company,[LERG1].OCN"
Case "RAO"
SQLOCNQry = SQLOCNQry & _
			"Order by BILL_RAO,[LERG1].OCN"
Case "OCN_Name"
SQLOCNQry = SQLOCNQry & _
	"Order by OCN_Name,[LERG1].OCN"

Case Else
SQLOCNQry = SQLOCNQry & _
			"Order by [LERG1].OCN"

End Select %>



<%
SET objConnection = server.createobject("ADODB.connection")
SET rstOCNQry = server.createobject("ADODB.recordset")
'objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
objConnection.open "DSN=cnasadmin;SERVER=cnac-db.database.windows.net;UID=SysAdmin;PWD=DbAccess460;APP=Microsoft Development Environment;WSID=CNAS.DOMIAN.CA;DATABASE=XCA_DB1;QueryLogFile=Yes"
SET rstOCNQry = objConnection.execute(sqlOCNQry)
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


<TABLE align="center" border="1">

<TD align="center"><B>OCN</B>
<a href="/xcalogon/LERG_OCN.asp?SortOrder=OCN">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>

<TD align="center"><B>RAO</B>
<a href="/xcalogon/LERG_OCN.asp?SortOrder=RAO">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>

<TD align="center"><B>OCN Name</B>
<a href="/xcalogon/LERG_OCN.asp?SortOrder=OCN_Name">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>

<TD align="center"><B>Company</B>
<a href="/xcalogon/LERG_OCN.asp?SortOrder=Company">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>


<p><br>
<br>
<% if rstOCNQry.EOF then %><b>No records found.</b> <% ELSE %> </p>
<% Do Until rstOCNQry.EOF %>
<tr align="left">
<td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstOCNQry("OCN") %> "><%=rstOCNQry("OCN") %> </a> &nbsp;</td>
</td>
<td>&nbsp;<%= rstOCNQry("RAO") %>&nbsp;</td>
<td>&nbsp;<%= rstOCNQry("OCN_NAME") %>&nbsp;</td>
<td>&nbsp;<%= rstOCNQry("COMPANY") %>&nbsp;</td>
</tr>
<% rstOCNQry.moveNext
loop %>
</table>
<% end if
objConnection.close
objConnectionLERG1.close %>
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
+"$RCSfile: LERG_OCN.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2006/05/17 16:01:03 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
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
+"$RCSfile: LERG_OCN.asp,v $\n"
+"$Revision: 1.2 $\n"
+"$Date: 2006/05/17 16:01:03 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
