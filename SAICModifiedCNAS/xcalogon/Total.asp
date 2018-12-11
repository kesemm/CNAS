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
<title>Company NPA Table</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Total.asp,v $
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
</form>
<p align="center"><strong>Total CO Codes Assigned By Company By NPA:
</p>
</strong>
<p>  <br>
</p>
<%
  SET objConnectionCompany = server.createobject("ADODB.connection")
  SET rstTotalCompany = server.createobject("ADODB.recordset")
  objConnectionCompany.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
  sqlTotalCompany = "Exec Company_Totals"
  SET rstTotalCompany = objConnectionCompany.execute(sqlTotalCompany)

  SET objConnectionCNA = server.createobject("ADODB.connection")
  SET rstTotalCNA = server.createobject("ADODB.recordset")
  objConnectionCNA.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
  sqlTotalCNA = "Exec CNA_Totals"
  SET rstTotalCNA = objConnectionCNA.execute(sqlTotalCNA)

  SET objConnectionMisc = server.createobject("ADODB.connection")
  SET rstTotalMisc = server.createobject("ADODB.recordset")
  objConnectionMisc.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
  sqlTotalMisc = "Exec Misc_Totals"
  SET rstTotalMisc = objConnectionMisc.execute(sqlTotalMisc)


  numColumnsCompany=rstTotalCompany.Fields.Count
  numColumnsCNA=rstTotalCNA.Fields.Count
  numColumnsMisc=rstTotalMisc.Fields.Count%>
<%Dim CNASum()
ReDim CNASum(numColumnsCNA)
Dim CompanySum()
ReDim CompanySum(Cint(numColumnsCompany))
Dim MiscSum()
ReDim MiscSum(Cint(numColumnsMisc))%>
</p>

<%rstTotalCompany.MoveFirst%>

<table Border="1">
  <tr>
<%for i=0 to numColumnsCompany-1%>
 <%if i=0 Then %>
    <th width=900>
 <%Else%>
    <th width=30>
<%End if%>
<%=rstTotalCompany.Fields(i).Name%>

</th>
<%Next%>
  </tr>
</tbody>
<% Do While Not rstTotalCompany.EOF%>
  <tr>
<% for i=0 to numColumnsCompany-1%>
    <td><%=rstTotalCompany.Fields(i)%>
    <%Companysum(i) = Companysum(i) + rstTotalCompany.Fields(i).Value%>
</td>
<%Next
rstTotalCompany.MoveNext
Loop%>
  </tr>
<tr>
<td>Sub Total Company Codes</td>
<% for i=1 to numColumnsCompany-1%>
    <td><%=Companysum(i)%>
</td>
<%Next %>
</table>
<br>

<%rstTotalCNA.MoveFirst%>

<table Border="1">
  <tr>
<%for i=0 to numColumnsCNA-1%>
 <%if i=0 Then %>
    <th width=900>
 <%Else%>
    <th width=30>
<%End if%>
<%=rstTotalCNA.Fields(i).Name%>
</th>
<%Next%>
  </tr>
</tbody>
<% Do While Not rstTotalCNA.EOF%>
  <tr>
<% for i=0 to numColumnsCNA-1%>
    <td><%=rstTotalCNA.Fields(i)%>
    <%CNAsum(i) = CNAsum(i) + rstTotalCNA.Fields(i).Value%>
</td>
<%Next
rstTotalCNA.MoveNext
Loop%>
  </tr>

<tr>
<td>Sub Total CNA Codes</td>
<% for i=1 to numColumnsCNA-1%>
    <td><%=CNAsum(i)%>
</td>
<%Next %>
</table>
<br>

<%rstTotalCNA.MoveFirst%>
<table Border="1">
  <tr>
<%for i=0 to numColumnsCNA-1%>
 <%if i=0 Then %>
    <th width=900>
 <%Else%>
    <th width=30><%=rstTotalCNA.Fields(i).Name%>
<%End if%>
</th>
<%Next%>

<tr>
<td>Total Company/CNA Codes</td>
<% for i=1 to numColumnsCNA-1%>
    <td><%=Companysum(i)+CNAsum(i)%>
</td>
<%Next %>
</tr>
<tr>
<td>Available CO Codes</td>
<% for i=1 to numColumnsCNA-2%>
    <td><%=800-(CNAsum(i) + Companysum(i))%>
</td>
<%Next %>
</tr>

</table>


<br>
<p align="center"><strong>The following CO Codes have been counted as available.</p></strong>
<%rstTotalMisc.MoveFirst%>

<table Border="1">
  <tr>
<%for i=0 to numColumnsMisc-1%>
 <%if i=0 Then %>
    <th width=900>
 <%Else%>
    <th width=30>
<%End if%>
<%=rstTotalMisc.Fields(i).Name%>

</th>
<%Next%>
  </tr>
</tbody>
<% Do While Not rstTotalMisc.EOF%>
  <tr>
<% for i=0 to numColumnsMisc-1%>
    <td><%=rstTotalMisc.Fields(i)%>
    <%Miscsum(i) = Miscsum(i) + rstTotalMisc.Fields(i).Value%>
</td>
<%Next
rstTotalMisc.MoveNext
Loop%>
  </tr>
<tr>
<td>Sub Total Misc Codes</td>
<% for i=1 to numColumnsMisc-1%>
    <td><%=Miscsum(i)%>
</td>
<%Next %>
</table>
<br>


<%objConnectionCompany.close
objConnectionCNA.close
objConnectionMisc.close %>
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
+"$RCSfile: Total.asp,v $\n"
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
