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
<title>RateCentre Query</title>
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
aNPA = request.querystring("NPA")
SQLNPARateCentre = "SELECT [LERG6].[RC_NAME10], Count([LERG6].NXX) AS CountOfNXX FROM [LERG6] GROUP BY [LERG6].[RC_NAME10], [LERG6].NPA HAVING ((([LERG6].NPA)= '" & aNPA & "')) ORDER BY [LERG6].[RC_NAME10];"
SET objConnection = server.createobject("ADODB.connection")
SET rstNPARateCentre = server.createObject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
rstNPARateCentre.activeConnection = objConnection
rstNPARateCentre.CursorLocation = adUseServer
rstNPARateCentre.CursorType = adOpenStatic
rstNPARateCentre.open SQLNPARateCentre
'SET rstNPARateCentre = objConnection.execute(SQLNPARateCentre)
%>
<b>

<p align="center">LERG NPA-Rate Centre query</b><br></p>
<p align="center">Click on the Rate Centre for a listing of NXXs</p>
<br>
</p>

<table align="center" BORDER="1">
  <tr>
    <td>Number of Rate Centres in NPA: <%= aNPA %></td>
    <td><%= rstNPARateCentre.recordcount %>
</td>
  </tr>
<%
do until rstNPARateCentre.EOF
%>
  <tr>
    <td align="center"><a HREF="LERG_RateCentreNXXQry.asp?NPA=<% =request.querystring("NPA") %>&RC=<%=rstNPARateCentre("RC_NAME10") %> "><%=rstNPARateCentre("RC_NAME10") %> </a> </td>
    <td><%= rstNPARateCentre("CountOfNXX") %>
</td>
  </tr>
<%
  rstNPARateCentre.movenext
loop
%>
</table>
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
+"$RCSfile: LERG_RateCentre_query.asp,v $\n"
+"$Revision: 1.4 $\n"
+"$Date: 2004/12/03 17:12:21 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
