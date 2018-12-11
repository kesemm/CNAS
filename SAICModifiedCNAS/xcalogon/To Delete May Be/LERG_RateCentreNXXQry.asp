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
<title>Rate Centre NXX Listing Query</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" -->
</form>

<% aRC=request.querystring("RC")
aNPA=request.querystring("NPA") %>
<%
  SET objConnection = server.createobject("ADODB.connection")
  SET rstNPANXXQry = server.createobject("ADODB.recordset")
  objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
  SQLNPANXXQry = "SELECT [LERG6].NPA, [LERG6].NXX, [LERG6].OCN, [LERG1].OCN_NAME, [LERG6].SWITCH, [LERG6].[LOC_ST], [LERG6].[RC_NAME10], [LERG6].Status, CONVERT(CHAR(10),[LERG6].[Eff_DATE],103) AS [EffDate] FROM [LERG1] INNER JOIN [LERG6] ON [LERG1].[OCN] = [LERG6].OCN WHERE ((([LERG6].NPA)='" & aNPA & "') AND (([LERG6].[RC_NAME10])='" & aRC & "')) Order By NXX;"
  SET RSTNPANXXQry = objConnection.execute(SQLNPANXXQry) %>

<p><br>
<p align="center"><b>Listing of NXXs for <%=aRC%> </b></p>
<table align="center" BORDER="1">
<tr>
<th align="center">NPA</th>
<th align="center">NXX</th>
<th align="center">Status</th>
<th align="center">Eff Date <br> (dd/mm/yyyy)</th>
<th align="center">Switch</th>
<th align="center">Rate Centre</th>
<th align="center">Province</th>
<th align="center">OCN</th>
<th align="center">Company</th>
<br>
<% if rstNPANXXQry.EOF then %><b>No record found for NPA <%= aNPA %> RC <%= aRC %>.</b> <% ELSE %> </p>
<% Do Until rstNPANXXQry.EOF %>
<tr align="center">
<td><%=rstNPANXXQry("NPA") %>
</td>
<td><a HREF="LERG_NPA_NXX.asp?NPA=<%=rstNPANXXQry("NPA")%>&NXX=<%=rstNPANXXQry("NXX")%> "><%= rstNPANXXQry("NXX") %> </a>
</td>
<td><%= rstNPANXXQry("Status") %>
</td>
<td><%= rstNPANXXQry("EffDate") %>
</td>
<td><%= rstNPANXXQry("Switch") %>
</td>
<td><%= rstNPANXXQry("RC_NAME10") %>
</td>
<td><%= rstNPANXXQry("LOC_ST") %>
</td>
<td><a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXQry("OCN") %> "><%=rstNPANXXQry("OCN") %> </a>
</td>
<td><%= rstNPANXXQry("OCN_NAME") %>
</td>
</tr>
<% rstNPANXXQry.moveNext
loop %>
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
+"$RCSfile: LERG_RateCentreNXXQry.asp,v $\n"
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
