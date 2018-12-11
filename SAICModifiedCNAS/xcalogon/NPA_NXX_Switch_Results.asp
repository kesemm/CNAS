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
<title>NPA NXX Switch Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPA_NXX_Switch_Results.asp,v $
'* Commit Date:   $Date: 2006/05/17 15:54:01 $ (UTC)
'* Committed by:  $Author: SAIC-OTTAWA\browng $
'* CVS Revision:  $Revision: 1.3 $
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

<p><%

aSwitch = request.querystring("Switch")

sqlSwitchQry ="SELECT " &_
"CNAS.NPA as NPA, " &_
"CNAS.NXX As NXX, " &_
"COStatusDescription As CNASDescription, " &_
"EntityName As CNASCompany, " &_
"CNAS.OCN as CNASOCN, " &_
"SwitchID As CNASSwitch, " &_
"StrandedCodeComment, " &_
"RateCenter As CNASRatecenter, " &_
"LERGSTATUS.[Description] As LERGDescription, " &_
"[LERG1].OCN_NAME As LERGCOMPANY, " &_
"LERG6.OCN As LERGOCN, " &_
"LERG6.Switch As LERG6Switch, " &_
"LERG6.RC_NAME As LERGRatecenter " &_
"From xca_COCode As CNAS " &_
"Full Join LERG6 " &_
"On CNAS.NPA=LERG6.NPA " &_
"And CNAS.NXX=LERG6.NXX " &_
"Left JOIN xca_status_codes " &_
"On CNAS.status=xca_status_codes.COStatus " &_
"Left Join xca_Entity as Company " &_
"On CNAS.EntityID=Company.EntityID " &_
"Left Join [LERGSTATUS] " &_
"On [LERG6].STATUS = [LERGSTATUS].STATUS " &_
"Full Join [LERG1] " &_
"On [LERG1].[OCN] = [LERG6].OCN  " &_
"WHERE (CNAS.SwitchID='" & aSwitch & "' OR LERG6.Switch='" & aSwitch & "') " &_
"ORDER by CNAS.NPA,CNAS.NXX"



SET objConnection1 = server.createobject("ADODB.connection")
SET rstSwitchQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET rstSwitchQry = objConnection1.execute(sqlSwitchQry)

%> </p>

<p align="center"><strong>NPA NXXs assigned to Switch : <a HREF="Detailed_Switch_Results.asp?Switch=<% = UCASE(aSwitch) %> "><% = UCASE(aSwitch) %> </a> &nbsp;</strong></p></td>
<b>

<p><br>
<table align="center" BORDER="1">
  <tr>

 <tr>
    <th align="center">&nbsp; NPA &nbsp;</th>
    <th align="center">&nbsp; NXX &nbsp;</th>
    <th align="center">&nbsp; CNAS &nbsp;<br>&nbsp; Status &nbsp;</th>
    <th align="center">&nbsp; CNAS &nbsp;<br>&nbsp; Company &nbsp;</th>
    <th align="center">&nbsp; CNAS &nbsp;<br>&nbsp; OCN &nbsp;</th>
    <th align="center">&nbsp; CNAS &nbsp;<br>&nbsp; Rate Centre &nbsp;</th>
	<th align="center">&nbsp; CNAS &nbsp;<br>&nbsp; Stranded Code Comment &nbsp;</th>
    <th align="center">&nbsp; LERG &nbsp;<br>&nbsp; Status &nbsp;</th>
    <th align="center">&nbsp; LERG &nbsp;<br>&nbsp; Company &nbsp;</th>
    <th align="center">&nbsp; LERG &nbsp;<br>&nbsp; OCN &nbsp;</th>
    <th align="center">&nbsp; LERG &nbsp;<br>&nbsp; Rate Centre &nbsp;</th>
    <td><br></td>
 </tr>

<% Do Until rstSwitchQry.EOF %>

  <tr align="left">
    <td>&nbsp;<%= rstSwitchQry("NPA") %>&nbsp;</td>
    <td>&nbsp;<a HREF="NPA_NXX_Result.asp?NPA=<%= rstSwitchQry("NPA") %>&NXX=<%= rstSwitchQry("NXX") %> "><%= rstSwitchQry("NXX") %>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("CNASDescription") %>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("CNASCompany") %>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("CNASOCN") %>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("CNASRatecenter") %>&nbsp;</td>
	<td>&nbsp;<%= rstSwitchQry("StrandedCodeComment") %>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("LERGDescription") %>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("LERGCompany") %>&nbsp;</td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%= rstSwitchQry("LERGOCN") %> "><%= rstSwitchQry("LERGOCN") %> </a>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("LERGRatecenter") %>&nbsp;</td>
  </tr>
<% rstSwitchQry.moveNext
 loop %>
</table>


</table>
<% 
objConnection1.close
%>
</b>
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
+"$RCSfile: NPA_NXX_Switch_Results.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2006/05/17 15:54:01 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
