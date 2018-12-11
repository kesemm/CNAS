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
'* Filename:      Detailed_Switch_Results.asp
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
SET objConnection1 = server.createobject("ADODB.connection")
SET rstSwitchQry =server.createobject("ADODB.recordset")
objConnection1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"

sqlSwitchQry ="Select " &_
"LERG7.OCN, " &_
"[LERG1].OCN_NAME, " &_
"[Description], " &_
"Convert(Char(10),[CREATION_DATE],103) As [CREATION_DATE], " &_
"Convert(Char(10),[EFF_DATE],103) As [EFF_DATE], " &_
"EQPT_TYPE, " &_
"MAJOR_VC, " &_
"MAJOR_HC, " &_
"IDDD, " &_
"LERG7.STREET, " &_
"LERG7.CITY, " &_
"LERG7.STATE, " &_
"LERG7.ZIP, " &_
"ORIG_FG_D, " &_
"ORIG_FG_D_INT, " &_
"ORIG_LOCAL, " &_
"TERM_FG_D, " &_
"TERM_FG_D_INT, " &_
"TERM_LOCAL, " &_
"HOST, " &_
"STP_1, " &_
"STP_2, " &_
"ACTUAL_ID, " &_
"CALL_AGENT, " &_
"TRUNK_GATEWAY, " &_
"Convert(Char(10),[LAST_CHANGE],103) As [LAST_CHANGE] " &_
"From LERG7 " &_
"Left Join [LERGSTATUS] " &_
"On [LERG7].STATUS = [LERGSTATUS].STATUS " &_
"Left Join [LERG1] " &_
"On [LERG1].[OCN] = [LERG7].OCN " &_
"WHERE LERG7.Switch_ID='" & aSwitch & "' "
SET rstSwitchQry = objConnection1.execute(sqlSwitchQry)

' 2018-09-14 K.T. Walsh Add the LERG7 Import data
SQLLergDateLERG7 = "SELECT CONVERT(CHAR(19), LERG7DATE.LERG7DATE, 120) AS LERG7DATE from LERG7DATE"
SET objConnectionLERG7 = server.createobject("ADODB.connection")
SET rstLergDateLERG7 = server.createobject("ADODB.recordset")
objConnectionLERG7.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstLergDateLERG7 = objConnectionLERG7.execute(SQLLergDateLERG7)


' 2018-09-14 K.T. Walsh Add NPA-NXX lookup information from CNAS and LERG data
SET objConnection2 = server.createobject("ADODB.connection")
SET rstSwitchNPANXXQry =server.createobject("ADODB.recordset")
objConnection2.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLSwitchNPANXXQry = "SELECT Tix,NPA,NXX,Status,COStatusDescription,EntityName,xca_COCode.OCN as OCN,SwitchID,WireCenter,RateCenter,InServiceDate,PublicRemarks,CNARemarks,StrandedCodeComment FROM xca_COCode Left Join xca_status_codes ON xca_COCode.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID WHERE (((xca_COCode.SwitchID)='" & aSwitch & "')) ORDER by NPA,NXX;"
SET rstSwitchNPANXXQry = objConnection2.execute(SQLSwitchNPANXXQry)

SET objConnection3 = server.createobject("ADODB.connection")
SET rstLERGSwitchNPANXXQry =server.createobject("ADODB.recordset")
objConnection3.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLERGSwitchNPANXXQry="SELECT A.NPA, A.NXX, A.OCN, [LERG1].OCN_NAME, A.SWITCH, A.[RC_NAME], [LERGSTATUS].[Description], CONVERT(CHAR(10),A.[Eff_DATE],120) AS [EffDate] " &_
"FROM [LERG6] as A " &_
"INNER JOIN [LERG1] ON [LERG1].[OCN] = A.OCN " &_
"INNER JOIN [LERGSTATUS] ON A.STATUS = [LERGSTATUS].STATUS " &_
"WHERE A.SWITCH='" & aSwitch & "' AND A.[Eff_Date]=(Select Max(LERG6.[Eff_Date]) From LERG6 where LERG6.NPA=A.NPA and LERG6.NXX=A.NXX) " &_
"ORDER by NPA,NXX"
SET rstLERGSwitchNPANXXQry=objConnection3.execute(SQLLERGSwitchNPANXXQry)

SQLLergDateLERG6 = "SELECT CONVERT(CHAR(19), LERG6DATE.LERG6DATE, 120) AS LERG6DATE from LERG6DATE"
SET objConnectionLERG6 = server.createobject("ADODB.connection")
SET rstLergDateLERG6 = server.createobject("ADODB.recordset")
objConnectionLERG6.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstLergDateLERG6 = objConnectionLERG6.execute(SQLLergDateLERG6)



%> </p>

<p align="center"><strong>Details for Switch : <% = UCASE(aSwitch) %> </strong></p>
<p align="center">Listing based on LERG 7 Data import date: <%=rstlergDateLERG7("LERG7DATE") %></p>


<% if (rstSwitchQry.EOF) then %><b></p>

<p>No record found for Switch in the LERG.</b> </p>
<% Else %>

<table align="center" BORDER="1">
  <tr align="left">
    <td>&nbsp;<b>OCN</b>&nbsp;</td>
    <td>&nbsp;<a HREF="LERG_OCN_Contact.asp?OCN=<%= rstSwitchQry("OCN") %> "><%= rstSwitchQry("OCN") %> </a>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Company</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("OCN_NAME") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Effective Date</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("EFF_DATE") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Status</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("Description") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Equipment Type</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("EQPT_TYPE") %>&nbsp;</td>
  </tr>
  
<% If IsNull(rstSwitchQry("ORIG_FG_D")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Orignating FG D</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("ORIG_FG_D") %> "><%= rstSwitchQry("ORIG_FG_D") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("ORIG_FG_D_INT")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Orignating FG D Intermediate</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("ORIG_FG_D_INT") %> "><%= rstSwitchQry("ORIG_FG_D_INT") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("ORIG_LOCAL")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Orignating Local</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("ORIG_LOCAL") %> "><%= rstSwitchQry("ORIG_LOCAL") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("TERM_FG_D")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Terminating FG D</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("TERM_FG_D") %> "><%= rstSwitchQry("TERM_FG_D") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("TERM_FG_D_INT")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Terminating FG D Intermediate</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("TERM_FG_D_INT") %> "><%= rstSwitchQry("TERM_FG_D_INT") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("TERM_LOCAL")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Terminating Local</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("TERM_LOCAL") %> "><%= rstSwitchQry("TERM_LOCAL") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("HOST")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>HOST</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("HOST") %> "><%= rstSwitchQry("HOST") %>&nbsp;</td>
  </tr>
<%End If%>
<% If IsNull(rstSwitchQry("STP_1")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>STP 1</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("STP_1") %> "><%= rstSwitchQry("STP_1") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("STP_2")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>STP 2</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("STP_2") %> "><%= rstSwitchQry("STP_2") %>&nbsp;</td>
  </tr>
<%End If%>


<% If IsNull(rstSwitchQry("ACTUAL_ID")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Actual Switch</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("ACTUAL_ID") %> "><%= rstSwitchQry("ACTUAL_ID") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("CALL_AGENT")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Call Agent</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("CALL_AGENT") %> "><%= rstSwitchQry("CALL_AGENT") %>&nbsp;</td>
  </tr>
<%End If%>

<% If IsNull(rstSwitchQry("TRUNK_GATEWAY")) Then %>
<%Else %>
  <tr align="left">
    <td>&nbsp;<b>Trunk Gateway</b>&nbsp;</td>
    <td>&nbsp;<a HREF="Detailed_Switch_Results.asp?Switch=<%= rstSwitchQry("TRUNK_GATEWAY") %> "><%= rstSwitchQry("TRUNK_GATEWAY") %>&nbsp;</td>
  </tr>
<%End If%>

  <tr align="left">
    <td>&nbsp;<b>IDDD</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("IDDD") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Major Vertical Co-ordinate</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("MAJOR_VC") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Major Horizontal Co-ordinate</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("MAJOR_HC") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Address</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("STREET") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>City</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("CITY") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Province</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("STATE") %>&nbsp;</td>
  </tr>

  <tr align="left">
    <td>&nbsp;<b>Postal Code</b>&nbsp;</td>
    <td>&nbsp;<%= rstSwitchQry("ZIP") %>&nbsp;</td>
  </tr>

</table>

<br>
<br>
<br>
<% ' 2018-09-14 K.T. Walsh Add NPA-NXX lookup information from CNAS and LERG data
%>

<p align="center"><strong>CNAS Switch Query for: <% = UCASE(aSwitch) %> </strong></p>

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

<% Do Until rstSwitchNPANXXQry.EOF %>    </td>
  </tr>
  <tr align="center">
    <td nowrap><%= rstSwitchNPANXXQry("Tix") %>
</td>

      <td nowrap><%= rstSwitchNPANXXQry("NPA") %>
</td>
    <td nowrap><%= rstSwitchNPANXXQry("NXX") %>
</td>
    <td nowrap><%= rstSwitchNPANXXQry("COStatusDescription") %>
</td>
    <td nowrap><%= rstSwitchNPANXXQry("EntityName") %>
</td>
    <td><%= rstSwitchNPANXXQry("OCN") %>
</td>
    <td nowrap><%= rstSwitchNPANXXQry("RateCenter") %>
</td>
    <td nowrap><%= rstSwitchNPANXXQry("PublicRemarks") %> &nbsp;
</td>
    <td nowrap><%= rstSwitchNPANXXQry("CNARemarks") %> &nbsp;
</td>
    <td nowrap><%= rstSwitchNPANXXQry("StrandedCodeComment") %> &nbsp;
</td>

  </tr>
<% rstSwitchNPANXXQry.moveNext
 loop %>
</table>
<br>
<br>
<br>
</b>
<p align="center"><strong>LERG Switch Query for: <% = UCASE(aSwitch) %> </strong></p>
<p align="center">Listing based on LERG 6 Data import date: <%=rstlergDateLERG6("LERG6DATE") %></p>


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

<% Do Until rstLERGSwitchNPANXXQry.EOF %>    </td>
  </tr>
      <td nowrap><%= rstLERGSwitchNPANXXQry("NPA") %>
</td>
    <td nowrap><%= rstLERGSwitchNPANXXQry("NXX") %>
</td>
    <td nowrap><%= rstLERGSwitchNPANXXQry("Description") %>
</td>
    <td nowrap><%= rstLERGSwitchNPANXXQry("OCN_NAME") %>
</td>
    <td><%= rstLERGSwitchNPANXXQry("OCN") %>
</td>
    <td nowrap><%= rstLERGSwitchNPANXXQry("RC_NAME") %>
</td>
  </tr>
<% rstLERGSwitchNPANXXQry.moveNext
 loop %>


</table>




<% End If %>
<% 
objConnection1.close
objConnection2.close
objConnection3.close
objConnectionLERG6.close
'objConnectionLERG7.close
%>
</b>
</body>
</html>
