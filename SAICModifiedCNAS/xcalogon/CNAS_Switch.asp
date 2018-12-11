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
<title>CNAS Database Query</title>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: CNAS_Switch.asp,v $
'* Commit Date:   $Date: 2014/04/21 16:07:55 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.4 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 

' 2014-04-21  Changed date format throughout when updating SQL code to return latest LERG record /ktwalsh
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
SET objConnection1 = server.createobject("ADODB.connection")
SET rstSwitchQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLSwitchQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.OCN as OCN,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks,StrandedCodeComment FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.SwitchID)='" & aSwitch & "')) ORDER by NPA,NXX;"
SET rstSwitchQry = objConnection1.execute(SQLSwitchQry)

SET objConnection2 = server.createobject("ADODB.connection")
SET rstLERGSwitchQry =server.createobject("ADODB.recordset")
objConnection2.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

' 2014-09-21 Add the LERG import time to the top of the page like the others /ktwalsh
SQLLergDateLERG6 = "SELECT CONVERT(CHAR(19), LERG6DATE.LERG6DATE, 120) AS LERG6DATE from LERG6DATE"
SET objConnectionLERG6 = server.createobject("ADODB.connection")
SET rstLergDateLERG6 = server.createobject("ADODB.recordset")
objConnectionLERG6.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstLergDateLERG6 = objConnectionLERG6.execute(SQLLergDateLERG6)

' 2014-04-21 Change SQL code to ensure only the latest LERG record for a given NPA/NXX is selected of the bunch /ktwalsh
SQLLERGSwitchQry="SELECT A.NPA, A.NXX, A.OCN, [LERG1].OCN_NAME, A.SWITCH, A.[RC_NAME], [LERGSTATUS].[Description], CONVERT(CHAR(10),A.[Eff_DATE],120) AS [EffDate] " &_
"FROM [LERG6] as A " &_
"INNER JOIN [LERG1] ON [LERG1].[OCN] = A.OCN " &_
"INNER JOIN [LERGSTATUS] ON A.STATUS = [LERGSTATUS].STATUS " &_
"WHERE A.SWITCH='" & aSwitch & "' AND A.[Eff_Date]=(Select Max(LERG6.[Eff_Date]) From LERG6 where LERG6.NPA=A.NPA and LERG6.NXX=A.NXX) " &_
"ORDER by NPA,NXX"

SET rstLERGSwitchQry=objConnection2.execute(SQLLERGSwitchQry)

%> </p>

<p align="center"><strong>CNAS Switch Query for: <% = UCASE(aSwitch) %> </strong></p>
<p align="center">Listing based on <%=rstlergDateLERG6("LERG6DATE") %> data</p>
<b>

<p>
<table align="center" BORDER="1">
  <tr>

 <tr>
<th align="center">Tix</th>
    <th align="center">NPA</th>
    <th align="center">NXX</th>
    <th align="center">Status</th>
    <th align="center">Company</th>
    <th align="center">OCN</th>
    <th align="center">Rate Centre</th>
    <th align="center">Public Remarks</th>
    <th align="center">CNA Remarks</th>
    <th align="center">Stranded Code Comment</th>
    <br>

<% Do Until rstSwitchQry.EOF %>    </td>
  </tr>
  <tr align="center">
    <td nowrap><%= rstSwitchQry("Tix") %>
</td>

      <td nowrap><%= rstSwitchQry("NPA") %>
</td>
    <td nowrap><%= rstSwitchQry("NXX") %>
</td>
    <td nowrap><%= rstSwitchQry("COStatusDescription") %>
</td>
    <td nowrap><%= rstSwitchQry("EntityName") %>
</td>
    <td><%= rstSwitchQry("OCN") %>
</td>
    <td nowrap><%= rstSwitchQry("RateCenter") %>
</td>
    <td nowrap><%= rstSwitchQry("PublicRemarks") %> &nbsp;
</td>
    <td nowrap><%= rstSwitchQry("CNARemarks") %> &nbsp;
</td>
    <td nowrap><%= rstSwitchQry("StrandedCodeComment") %> &nbsp;
</td>

  </tr>
<% rstSwitchQry.moveNext
 loop %>
</table>
<br>

<p align="center"><strong>LERG Switch Query for: <% = UCASE(aSwitch) %> </strong></p>
<b>

<p>
<table align="center" BORDER="1">
  <tr>

 <tr>
    <th align="center">NPA</th>
    <th align="center">NXX</th>
    <th align="center">Status</th>
    <th align="center">Company</th>
    <th align="center">OCN</th>
    <th align="center">Rate Centre</th>
    <br>

<% Do Until rstLERGSwitchQry.EOF %>    </td>
  </tr>
      <td nowrap><%= rstLERGSwitchQry("NPA") %>
</td>
    <td nowrap><%= rstLERGSwitchQry("NXX") %>
</td>
    <td nowrap><%= rstLERGSwitchQry("Description") %>
</td>
    <td nowrap><%= rstLERGSwitchQry("OCN_NAME") %>
</td>
    <td><%= rstLERGSwitchQry("OCN") %>
</td>
    <td nowrap><%= rstLERGSwitchQry("RC_NAME") %>
</td>
  </tr>
<% rstLERGSwitchQry.moveNext
 loop %>


</table>

<% 
objConnection1.close
objConnection2.close
objConnectionLERG6.close
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
strAlertText="Leidos Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: CNAS_Switch.asp,v $\n"
+"$Revision: 1.4 $\n"
+"$Date: 2014/04/21 16:07:55 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
