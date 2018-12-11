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
objConnectionLERG.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"



SQLNPANXXQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.OCN as OCN1,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks,StrandedCodeComment FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.NPA)='" & aNPA &"') AND ((xca_COCode.NXX)='" & aNXX & "'));"
SQLPart1Qry="Select Max(RequestedEffDate) As RequestedEffDate From xca_Part1 WHERE (((xca_Part1.NPA)='" & aNPA &"') AND ((xca_Part1.NXX1Preferred)='" & aNXX & "'));"
SET rstNPANXXCNAS = objConnection1.execute(SQLNPANXXQry)
SET RSTPart1Qry=objConnection2.execute(SQLPart1Qry)

sqlNPANXXLERG = "SELECT [LERG6].NPA, [LERG6].NXX, [LERG6].OCN, [LERG1].OCN_NAME, [LERG6].SWITCH,  [LERG6].[RC_NAME], [LERGSTATUS].Description, CONVERT(CHAR(10),[LERG6].[Eff_DATE],103) AS [EffDate] " & _
		"FROM [LERG1] " & _ 
		"INNER JOIN [LERG6] ON [LERG1].[OCN] = [LERG6].OCN " & _
            "INNER JOIN [LERGSTATUS] ON [LERG6].STATUS = [LERGSTATUS].STATUS " & _
		"WHERE [LERG6].NPA='" & aNPA & "' AND [LERG6].NXX='" & aNXX & "'"
SET rstNPANXXLERG = objConnectionLERG.execute(sqlNPANXXLERG) %> </p>

<p align="center"><strong>CNAS NPA-NXX Query </strong></p>
<p align="center">Lising for NPA <strong><%= aNPA %> </strong>NXX <strong><%= aNXX %></strong></p>
<b>

<p><br>
<% if (rstNPANXXCNAS.EOF and rstNPANXXLERG.EOF) then %><b></p>

<p>No record found for NPA <%= aNPA %> NXX <%= aNXX %>.</b> </p>
<% Elseif rstNPANXXLERG.EOF then %>
CNAS Only
<% Elseif rstNPANXXCNAS.EOF then %>
LERG Only
<% Else %>
<% Elseif rstNPANXXCNAS("Status")="I" then %>

<table align="center" BORDER="1">
  <tr align="center">
    <td></td>
    <td>CNAS Data</td>
    <td>LERG Data</td>	
  </tr>
  <tr align="center">
    <td><b>Ticket Number</b></td>
    <td><%= rstNPANXXCNAS("Tix") %></td>
    <td>N/A</td>
  </tr>
  <tr align="center">
    <td><b>Status</b></td>
    <td><%= rstNPANXXCNAS("COStatusDescription") %>
    <td><%= rstNPANXXLERG("Description") %>
</td>
  </tr>
  <tr align="center">
    <td><b>Company</b></td>
    <td><%= rstNPANXXCNAS("EntityName") %></td>
    <td><%= rstNPANXXLERG("OCN_NAME") %>
  </tr>
  <tr align="center">
    <td><b>OCN</b></td>
    <td><a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXCNAS("OCN1") %> "><%=rstNPANXXCNAS("OCN1") %> </a></td>
    <td><a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXLERG("OCN") %> "><%=rstNPANXXLERG("OCN") %> </a></td>

  </tr>
  <tr align="center">
    <td><b>SwitchID</b></td>
    <td><%= rstNPANXXCNAS("SwitchID") %></td>
    <td><%= rstNPANXXLERG("SWITCH") %>

  </tr>
  <tr align="center">
    <td><b>WireCentre</b></td>
    <td><%= rstNPANXXCNAS("WireCenter") %></td>
    <td>N/A</td>
  </tr>
  <tr align="center">
    <td><b>RateCentre</b></td>
    <td><%= rstNPANXXCNAS("RateCenter") %></td>
    <td><%= rstNPANXXLERG("RC_NAME") %>
  </tr>
  <tr align="center">
    <td><b>InService Date <br>(dd/mm/yyyy)</b></td>
    <td><%= rstNPANXXCNAS("InServiceDate") %></td>
    <td>N/A</td>
  </tr>
  </tr>
  <tr align="center">
    <td><b>Effective Date <br>(dd/mm/yyyy)</b></td>
    <td>N/A</td>
    <td><%= rstNPANXXLERG("EFFDATE") %>
  </tr>
  <tr align="center">
    <td><b>Public Remarks</b></td>
    <td><%= rstNPANXXCNAS("PublicRemarks") %></td>
  </tr>
  <tr align="center">
    <td><b>CNA Remarks</b></td>
    <td><%= rstNPANXXCNAS("CNARemarks") %></td>
  </tr>
  <tr align="center">
    <td><b>Stranded Code Comment</b></td>
    <td><%= rstNPANXXCNAS("StrandedCodeComment") %></td>
  </tr>

</table>

<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<% Elseif rstNPANXXCNAS("Status")="A" then %>
<table align="center" BORDER="1">
  <tr align="center">
    <td></td>
    <td>CNAS Data</td>
    <td>LERG Data</td>	
  </tr>
  <tr align="center">
    <td><b>Ticket Number</b></td>
    <td><%= rstNPANXXCNAS("Tix") %></td>
    <td>N/A</td>
  </tr>
  <tr align="center">
    <td><b>Status</b></td>
    <td><%= rstNPANXXCNAS("COStatusDescription") %>
    <td><%= rstNPANXXLERG("Description") %>
</td>
  </tr>
  <tr align="center">
    <td><b>Company</b></td>
    <td><%= rstNPANXXCNAS("EntityName") %></td>
    <td><%= rstNPANXXLERG("OCN_NAME") %>
  </tr>
  <tr align="center">
    <td><b>OCN</b></td>
    <td><a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXCNAS("OCN1") %> "><%=rstNPANXXCNAS("OCN1") %> </a></td>
    <td><a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXLERG("OCN") %> "><%=rstNPANXXLERG("OCN") %> </a></td>

  </tr>
  <tr align="center">
    <td><b>SwitchID</b></td>
    <td><%= rstNPANXXCNAS("SwitchID") %></td>
    <td><%= rstNPANXXLERG("SWITCH") %>

  </tr>
  <tr align="center">
    <td><b>WireCentre</b></td>
    <td><%= rstNPANXXCNAS("WireCenter") %></td>
    <td>N/A</td>
  </tr>
  <tr align="center">
    <td><b>RateCentre</b></td>
    <td><%= rstNPANXXCNAS("RateCenter") %></td>
    <td><%= rstNPANXXLERG("RC_NAME") %>
  </tr>
  <tr align="center">
    <td><b>Requested Effective Date <br>(dd/mm/yyyy)</b></td>
    <td><%= rstPart1Qry("RequestedEffDate") %></td>

  </tr>
  <tr align="center">
    <td><b>Public Remarks</b></td>
    <td><%= rstNPANXXCNAS("PublicRemarks") %></td>
  </tr>
  <tr align="center">
    <td><b>CNA Remarks</b></td>
    <td><%= rstNPANXXCNAS("CNARemarks") %></td>
  </tr>
  <tr align="center">
    <td><b>Stranded Code Comment</b></td>
    <td><%= rstNPANXXCNAS("StrandedCodeComment") %></td>
  </tr>

</table>

<%else%>

<table align="center" BORDER="1">
  <tr align="center">
    <td><b>Ticket Number</b></td>
    <td><%= rstNPANXXCNAS("Tix") %></td>
    <td>N/A</td>
  </tr>
  <tr align="center">
    <td><b>NPA</b></td>
    <td><%= rstNPANXXCNAS("NPA") %></td>
  </tr>
  <tr align="center">
    <td><b>NXX</b></td>
    <td><%= rstNPANXXCNAS("NXX") %></td>
  </tr>
  <tr align="center">
    <td><b>Status</b></td>
    <td><%= rstNPANXXCNAS("COStatusDescription") %></td>
<td><%= rstNPANXXLERG("Status") %></td>
  </tr>
  <tr align="center">
    <td><b>Company</b></td>
    <td><%= rstNPANXXCNAS("EntityName") %></td>
  </tr>
  <tr align="center">
    <td><b>OCN</b></td>
    <td><%= rstNPANXXCNAS("OCN1") %></td>
<td><a HREF="LERG_OCN_Contact.asp?OCN=<%=rstNPANXXLERG("OCN") %> "><%=rstNPANXXLERG("OCN") %> </a></td>

  </tr>
  <tr align="center">
    <td><b>SwitchID</b></td>
    <td><%= rstNPANXXCNAS("SwitchID") %></td>
  </tr>
  <tr align="center">
    <td><b>WireCentre</b></td>
    <td><%= rstNPANXXCNAS("WireCenter") %></td>
  </tr>
  <tr align="center">
    <td><b>RateCentre</b></td>
    <td><%= rstNPANXXCNAS("RateCenter") %></td>
  </tr>
  <tr align="center">
    <td><b>Public Remarks</b></td>
    <td><%= rstNPANXXCNAS("PublicRemarks") %></td>
  </tr>
  <tr align="center">
    <td><b>CNA Remarks</b></td>
    <td><%= rstNPANXXCNAS("CNARemarks") %></td>
  </tr>
  <tr align="center">
    <td><b>Stranded Code Comment</b></td>
    <td><%= rstNPANXXCNAS("StrandedCodeComment") %></td>
  </tr>

</table>

<p>Note: A ticket number of 999999999 implies we received the CO Code as Assigned or
InService. InService Date is not correct under this condition.</p>
<% end if
objConnection1.close
objConnection2.close
objConnectionLERG.close %>
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
+"$RCSfile: LERG_NPA_NXX.asp,v $\n"
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
