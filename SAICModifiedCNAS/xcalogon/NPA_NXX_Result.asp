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
'* CVS File:      $RCSfile: NPA_NXX_Result.asp,v $
'* Commit Date:   $Date: 2014/12/17 13:00:19 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.6 $
'* Checkout Tag:  $Name:  $ (Version/Build)
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
<!-- #Include file="ADOVBS.INC" -->
</form>
<p><%

aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")

SET objConnection1 = server.createobject("ADODB.connection")
SET rstNPANXXCNAS =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET objConnection2 = server.createobject("ADODB.connection")
SET RSTPart1Qry =server.createobject("ADODB.recordset")
objConnection2.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET objConnectionLERG = server.createobject("ADODB.connection")
SET rstNPANXXLERG = server.createobject("ADODB.recordset")
' 2014-04-19 Change date format
SQLLergDateLERG6 = "SELECT CONVERT(CHAR(19), LERG6DATE.LERG6DATE, 120) AS LERG6DATE from LERG6DATE"
objConnectionLERG.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

SET objConnectionLERG6 = server.createobject("ADODB.connection")
SET rstLergDateLERG6 = server.createobject("ADODB.recordset")
objConnectionLERG6.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstLergDateLERG6 = objConnectionLERG6.execute(SQLLergDateLERG6)

' 2014-04-19 Change date format /ktwalsh
SQLNPANXXQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.EntityID As EntityID,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,CONVERT(char(10), InServiceDate,120) as [InServiceDate],PublicRemarks,CNARemarks,StrandedCodeComment FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SQLPart1Qry="Select CONVERT(char(10), Max(RequestedEffDate),120) As RequestedEffDate From xca_Part1 WHERE (((xca_Part1.NPA)='" & aNPA &"') AND ((xca_Part1.NXX1Preferred)='" & aNXX & "'));"
SET rstNPANXXCNAS = objConnection1.execute(SQLNPANXXQry)
SET RSTPart1Qry=objConnection2.execute(SQLPart1Qry)

' 2014-04-21 Change SQL code to ensure only the latest LERG record is selected of the bunch /ktwalsh
sqlNPANXXLERG = "SELECT [LERG6].NPA, [LERG6].NXX, [LERG6].OCN, [LERG1].OCN_NAME, [LERG6].SWITCH,  [LERG6].[RC_NAME], [LERGSTATUS].Description, CONVERT(CHAR(10),[LERG6].[Eff_DATE],120) AS [EffDate] " & _
		"FROM [LERG1] " & _ 
		"INNER JOIN [LERG6] ON [LERG1].[OCN] = [LERG6].OCN " & _
            "INNER JOIN [LERGSTATUS] ON [LERG6].STATUS = [LERGSTATUS].STATUS " & _
		"WHERE [LERG6].NPA='" & aNPA & "' AND [LERG6].NXX='" & aNXX & "' AND LERG6.[Eff_Date]=(Select Max(LERG6.[Eff_Date]) From LERG6 where LERG6.NPA='" & aNPA & "' and LERG6.NXX='" & aNXX & "')"

SET rstNPANXXLERG = objConnectionLERG.execute(sqlNPANXXLERG) %> </p>

<p align="center"><strong>CNAS NPA-NXX Query </strong></p>
<p align="center">Listing for NPA <strong><%= aNPA %> </strong>NXX <strong><%= aNXX %></strong> based on <%=rstlergDateLERG6("LERG6DATE") %> data</p>
<b>

<p><br>
<% if (rstNPANXXCNAS.EOF and rstNPANXXLERG.EOF) then %><b></p>

<p>No record found for NPA <%= aNPA %> NXX <%= aNXX %> in CNAS or the LERG.</b> </p>

<% Elseif rstNPANXXLERG.EOF then %>
<p align="center"><strong>CNAS Data Only.  No LERG Data</strong></p>

<table align="center" BORDER="1">

  <tr align="left">
    <td>&nbsp; &nbsp;</td>
    <td>&nbsp;CNAS Data&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Ticket Number&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("Tix") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Status&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("COStatusDescription") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Company&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("EntityName") %>&nbsp;</td>
  </tr>
  
  <tr align="left">
    <td><b>&nbsp;EntityID&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("EntityID") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;OCN&nbsp;</b></td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXCNAS("OCN1") %> "><%=rstNPANXXCNAS("OCN1") %> </a>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (NXX)</b> &nbsp;</td>
    <td>&nbsp;<a HREF="NPA_NXX_Switch_Results.asp?Switch=<%= rstNPANXXCNAS("SwitchID") %> "><%= rstNPANXXCNAS("SwitchID") %> </a> &nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (Details)&nbsp;</b></td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstNPANXXCNAS("SwitchID") %> "><%= rstNPANXXCNAS("SwitchID") %> </a> &nbsp;</td>
  </tr>


  <tr align="left">
    <td><b>&nbsp;WireCentre&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("WireCenter") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;RateCentre&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("RateCenter") %>&nbsp;</td>
  </tr>

<%if rstNPANXXCNAS("Status")="I" then %>
  <tr align="left">
    <td><b>&nbsp;InService Date&nbsp; <br>&nbsp;(yyyy-mm-dd)&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("InServiceDate") %>&nbsp;</td>
  </tr>
<% end if %>



<%if rstNPANXXCNAS("Status") = "A" then %>
  <tr align="left">
    <td><b>&nbsp;Effective Date &nbsp;<br>&nbsp;(yyyy-mm-dd)</b>&nbsp;</td>
    <td>&nbsp;<%= rstPart1Qry("RequestedEffDate") %>&nbsp;</td>
  </tr>
<% end if %>

  <tr align="left">
    <td><b>&nbsp;Public Remarks&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("PublicRemarks") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;CNA Remarks&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("CNARemarks") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Stranded Code Comment</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("StrandedCodeComment") %>&nbsp;</td>
  </tr>

</table>


<%if rstNPANXXCNAS("Status")="I" then %>
<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<% end if %>


<% Elseif rstNPANXXCNAS.EOF then %>
<p align="center"><strong>LERG Data Only.  No CNAS Data</strong></p>

<table align="center" BORDER="1">

  <tr align="left">
    <td>&nbsp; &nbsp;</td>
    <td>&nbsp;LERG Data&nbsp;</td>	
  </tr>

  <tr align="left">
    <td><b>&nbsp;Status</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("Description") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Company</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("OCN_NAME") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;OCN</b>&nbsp;</td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXLERG("OCN") %> "><%=rstNPANXXLERG("OCN") %> </a>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (NXX)</b> &nbsp;</td>
    <td>&nbsp;<a HREF="NPA_NXX_Switch_Results.asp?Switch=<%= rstNPANXXLERG("Switch") %> "><%= rstNPANXXLERG("Switch") %> </a> &nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (Details)</b> &nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstNPANXXLERG("Switch") %> "><%= rstNPANXXLERG("Switch") %> </a> &nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>RateCentre</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("RC_NAME") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Effective Date <br>(yyyy-mm-dd)</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("EFFDATE") %>&nbsp;</td>
  </tr>
</table>

<% Else %>

<table align="center" BORDER="1">

  <tr align="left">
    <td>&nbsp; &nbsp;</td>
    <td>&nbsp;CNAS Data&nbsp;</td>
    <td>&nbsp;LERG Data&nbsp;</td>	
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Ticket Number</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("Tix") %>&nbsp;</td>
    <td>&nbsp; &nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Status</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("COStatusDescription") %>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("Description") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Company</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("EntityName") %>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("OCN_NAME") %>&nbsp;</td>
  </tr>
  
   <tr align="left">
    <td><b>&nbsp;EntityID&nbsp;</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("EntityID") %>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>OCN</b>&nbsp;</td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXCNAS("OCN1") %> "><%=rstNPANXXCNAS("OCN1") %> </a>&nbsp;</td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXLERG("OCN") %> "><%=rstNPANXXLERG("OCN") %> </a>&nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (NXX)</b> &nbsp;</td>
    <td>&nbsp;<a HREF="NPA_NXX_Switch_Results.asp?Switch=<%= rstNPANXXCNAS("SwitchID") %> "><%= rstNPANXXCNAS("SwitchID") %> </a> &nbsp;</td>
    <td>&nbsp;<a HREF="NPA_NXX_Switch_Results.asp?Switch=<%= rstNPANXXLERG("Switch") %> "><%= rstNPANXXLERG("Switch") %> </a> &nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>&nbsp;Switch (Details)</b> &nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstNPANXXCNAS("SwitchID") %> "><%= rstNPANXXCNAS("SwitchID") %> </a> &nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstNPANXXLERG("Switch") %> "><%= rstNPANXXLERG("Switch") %> </a> &nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>WireCentre</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("WireCenter") %>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>RateCentre</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("RateCenter") %>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("RC_NAME") %>&nbsp;</td>
  </tr>

<%if rstNPANXXCNAS("Status")="I" then %>
  <tr align="left">
    <td>&nbsp;<b>InService Date &nbsp;<br>&nbsp;(yyyy-mm-dd)</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("InServiceDate") %>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Effective Date &nbsp; <br>&nbsp;(yyyy-mm-dd)</b>&nbsp;</td>
    <td>&nbsp;<%= rstPart1Qry("RequestedEffDate") %>&nbsp;</td>
   <td>&nbsp;<%= rstNPANXXLERG("EFFDATE") %>&nbsp;</td>
  </tr>
<% end if %>

<%if rstNPANXXCNAS("Status")="A" then %>
  <tr align="left">
    <td>&nbsp;<b>Effective Date &nbsp;<br>&nbsp;(yyyy-mm-dd)</b>&nbsp;</td>
    <td>&nbsp;<%= rstPart1Qry("RequestedEffDate") %>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXLERG("EFFDATE") %>&nbsp;</td>
  </tr>
<% end if %>

  <tr align="left">
    <td>&nbsp;<b>Public Remarks</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("PublicRemarks") %>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>CNA Remarks</b></td>
    <td>&nbsp;<%= rstNPANXXCNAS("CNARemarks") %>&nbsp;</td>
    <td>&nbsp; &nbsp;</td>
  </tr>

  <tr align="left">
    <td><b>Stranded Code Comment</b>&nbsp;</td>
    <td>&nbsp;<%= rstNPANXXCNAS("StrandedCodeComment") %>&nbsp;</td>
    <td>&nbsp; &nbsp;</td>
  </tr>

</table>


<%if rstNPANXXCNAS("Status")="I" then %>
<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<% end if %>

<% end if
objConnection1.close
objConnection2.close
objConnectionLERG.close
objConnectionLERG6.close %>
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
+"$RCSfile: NPA_NXX_Result.asp,v $\n"
+"$Revision: 1.6 $\n"
+"$Date: 2014/12/17 13:00:19 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</body>
</html>
