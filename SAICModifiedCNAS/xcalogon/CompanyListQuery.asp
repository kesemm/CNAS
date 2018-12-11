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

<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>Logon Data</title>
<p align="center"><b>List of Companies with CO Codes </p></b>

<%
'****************************************************************************************
'* CVS File:      $RCSfile: CompanyListQuery.asp,v $
'* Commit Date:   $Date: 2010/12/21 18:35:42 $ (UTC)
'* Committed by:  $Author: browng $
'* CVS Revision:  $Revision: 1.1 $
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
sqlQry = "Select distinct xca_Entity_OCN_Web.Company,OCN_NAME As NECA_Name,xca_COCode.OCN As OCN " &_
" From xca_COCode" &_ 
" Inner Join xca_Entity_OCN_Web" &_
" On xca_COCode.OCN=xca_Entity_OCN_Web.OCN" &_
" Inner Join LERG1" &_
" On xca_COCode.OCN=LERG1.OCN" &_
" Where Status In ('I','A','Q','R','4')" &_
" Order By xca_Entity_OCN_Web.Company,OCN"
 %>



<%
SET objConnection = server.createobject("ADODB.connection")
SET rstQry = server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstQry = objConnection.execute(sqlQry)
%>


<TABLE align="center" border="1">

<TD align="center"><B>Company</B></TD>

<TD align="center"><B>NECA Name</B></TD>

<TD align="center"><B>OCN</B></TD>

<p><br>
<br>
<% if rstQry.EOF then %><b>No records found.</b> <% ELSE %> </p>
<% Do Until rstQry.EOF %>
<tr align="left">
<td>&nbsp;<%=rstQry("Company") %> </a> &nbsp;</td>
</td>
<td>&nbsp;<%= rstQry("NECA_Name") %>&nbsp;</td>
<td>&nbsp;<%= rstQry("OCN") %>&nbsp;</td>
</tr>
<% rstQry.moveNext
loop %>
</table>
<% end if
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
+"$RCSfile: CompanyListQuery.asp,v $\n"
+"$Revision: 1.1 $\n"
+"$Date: 2010/12/21 18:35:42 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
