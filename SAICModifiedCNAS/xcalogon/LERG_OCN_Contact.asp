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
<title>OCN Contact Listing</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: LERG_OCN_Contact.asp,v $
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

<% aOCN=request.querystring("OCN")%>
<%
  SET objConnection = server.createobject("ADODB.connection")
  SET rstOCNContactQry = server.createobject("ADODB.recordset")
  objConnection.open "DSN=cnasadmin;SERVER=cnac-db.database.windows.net;UID=SysAdmin;PWD=DbAccess460"
  sqlOCNContactQry = "SELECT [OCN],[OCN_NAME],[OCN_ST],[OCN_CODE],[TITLE],[LAST_NAME],[FIRST_NAME],[COMPANY],[ADDRESS],[CITY],[ZIP],[PHONE] FROM [LERG1]WHERE OCN='" & aOCN& "';"
  SET rstOCNContactQry = objConnection.execute(sqlOCNContactQry) %>

<% if rstOCNContactQry.EOF then %><b></p>

<p>No record found for OCN <%= aOCN %>.</b> <% Else%> </p>

<table align="center" BORDER="1">
  <tr align="left">
    <td>&nbsp;<b>OCN</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("OCN") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>OCN Name</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("OCN_NAME") %>&nbsp;</td>
  </tr>
    <tr align="left">
    <td>&nbsp;<b>OCN Type</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("OCN_CODE") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>First Name</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("FIRST_NAME") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>Last Name</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("LAST_NAME") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>Company</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("COMPANY") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>Address</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("ADDRESS") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>City</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("CITY") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>Province</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("OCN_ST") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>Phone</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("PHONE") %>&nbsp;</td>
  </tr>
  <tr align="left">
    <td>&nbsp;<b>Postal Code</b>&nbsp;</td>
    <td>&nbsp;<%= rstOCNContactQry("ZIP") %>&nbsp;</td>
  </tr>

</table>
<% end if
objConnection.close %>
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
+"$RCSfile: LERG_OCN_Contact.asp,v $\n"
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
