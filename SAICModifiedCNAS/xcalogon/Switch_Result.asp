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
<title>NPA / NXX Querry</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Switch_Result.asp,v $
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
<p><%

aRatecentre = request.querystring("Ratecentre")
aProvince = request.querystring("Province")
aBuilding = request.querystring("Building")
aEquipment = request.querystring("Equipment")

If aRatecentre = "" Then aRatecentre = "Null"
If aProvince = "" Then aProvince = "Null"
If aBuilding = "" Then aBuilding = "Null"
If aEquipment = "" Then aEquipment = "Null"

SET objConnectionLERG = server.createobject("ADODB.connection")
SET rstLERG = server.createobject("ADODB.recordset")
objConnectionLERG.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlLERG = "exec Switch_Lookup '" & aRatecentre & "','" &  aProvince & "','" & aBuilding & "','" & aEquipment & "'"
SET rstLERG = objConnectionLERG.execute(sqlLERG) %> </p>

<p align="center"><strong>Switch Query</strong></p>
<p align="center">Listing for <strong><%=request.querystring("Ratecentre")%><%=request.querystring("Province")%><%=request.querystring("Building")%><%=request.querystring("Equipment")%>
</strong></p>

<% if rstLERG.EOF then %><b><p>No record found for: <%=request.querystring("Ratecentre")%><%=request.querystring("Province")%><%=request.querystring("Building")%><%=request.querystring("Equipment")%>
<%end if%></p>
<p><br>
<table align="center" BORDER="1">
<tr align="left">
<td>Switch</td>
<td>OCN</td>
<td>Company</td>
<td>Address</td>
<td>City</td>
<td>Postal Code</td>
<% Do Until rstLERG.EOF %>
<tr align="left">
<td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstLERG("Switch_ID") %> "><%= rstLERG("Switch_ID") %> </a> &nbsp;</td>
<td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstLERG("OCN") %> "><%=rstLERG("OCN") %> </a> &nbsp;</td>
<td>&nbsp;<%=rstLERG("OCN_NAME")%>&nbsp;</td>
<td>&nbsp;<%=rstLERG("STREET")%>&nbsp;</td>
<td>&nbsp;<%=rstLERG("CITY")%>&nbsp;</td>
<td>&nbsp;<%=rstLERG("ZIP")%>&nbsp;</td>
</tr>
<% rstLERG.moveNext
loop %>
</table>
<%objConnectionLERG.close %>
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
+"$RCSfile: Switch_Result.asp,v $\n"
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
+"$RCSfile: Switch_Result.asp,v $\n"
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
