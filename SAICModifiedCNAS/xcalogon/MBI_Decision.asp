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
<title>MBI Part 1 Decision</title>
<%UserEntityType=session("UserEntityType")%>
 
 <script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">
Sub btnOK_onclick()
Response.Redirect "MBI_Select_NPA-NXX.asp"
End Sub
</script>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<p><%

aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXMBIQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXMBIQry = "SELECT Tix,NPA,NXX,MBI,COStatusDescription,RateCenter,EntityName,xca_MBI.OCN,xca_MBI.EntityID,PublicRemarks,CNARemarks FROM xca_MBI Left Join xca_status_codes ON xca_MBI.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_MBI.EntityID=xca_Entity.EntityID WHERE (((xca_MBI.NPA)='" & aNPA &"') AND ((xca_MBI.NXX)='" & aNXX & "')) ORDER BY MBI;"
SET RSTNPANXXMBIQry = objConnection.execute(SQLNPANXXMBIQry)

SET objConnectionNotAvailable = server.createobject("ADODB.connection")
SET rstNotAvailableQry =server.createobject("ADODB.recordset")
objConnectionNotAvailable.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
NotAvailableQry = "SELECT Count (*) As Number_of_Records FROM xca_MBI WHERE xca_MBI.Status <> 'S' And xca_MBI.NPA='" & aNPA &"' And xca_MBI.NXX='" & aNXX & "';"
SET rstNotAvailableQry = objConnectionNotAvailable.execute(NotAvailableQry)

SET objConnectionAvailable = server.createobject("ADODB.connection")
SET rstAvailableQry =server.createobject("ADODB.recordset")
objConnectionAvailable.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
AvailableQry = "SELECT Count (*) As Number_of_Records FROM xca_MBI WHERE xca_MBI.Status = 'S' And xca_MBI.NPA='" & aNPA &"' And xca_MBI.NXX='" & aNXX & "';"
SET rstAvailableQry = objConnectionAvailable.execute(AvailableQry)

%> </p>
<% if RSTNPANXXMBIQry.EOF then %>
<b><p>No record found for:  <% = (aNPA) %> - <% = (aNXX) %>
<tr><br></tr>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnOK style="HEIGHT: 27px; LEFT: 10px; TOP: 34px; WIDTH: 36px" 
            width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="OK">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnOK()
{
	btnOK.value = 'OK';
	btnOK.setStyle(0);
}
function _btnOK_ctor()
{
	CreateButton('btnOK', _initbtnOK, null);
}
</script>
<% btnOK.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<%end if%>

<% if rstAvailableQry ("Number_of_Records")= 10 then %>
<% Response.Redirect ("MBI_Full_Part1.asp") %>
<%end if%>

<% if rstNotAvailableQry ("Number_of_Records")= 10 then %>
<b><p>There are no 1000 blocks available for assignment for:  <% = (aNPA) %> - <% = (aNXX) %> 
<tr><br></tr>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnOK style="HEIGHT: 27px; LEFT: 10px; TOP: 34px; WIDTH: 36px" 
            width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="OK">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnOK()
{
	btnOK.value = 'OK';
	btnOK.setStyle(0);
}
function _btnOK_ctor()
{
	CreateButton('btnOK', _initbtnOK, null);
}
</script>
<% btnOK.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<%end if%></p>

<p align="center"><strong>MBIs Assigned to <% = aNPA %> - <% = aNXX %> </strong></p></td>
<b>
<tr><% = rstAvailableQry ("Number_of_Records") %></tr>
<tr><% = rstNotAvailableQry ("Number_of_Records") %></tr>
<p><br>
<table align="center" BORDER="1">
  <tr>

 <tr>
    <th align="center">&nbsp; Tix &nbsp;</th>
    <th align="center">&nbsp; NPA &nbsp;</th>
    <th align="center">&nbsp; NXX &nbsp;</th>
    <th align="center">&nbsp; MBI &nbsp;</th>
	<th align="center">&nbsp; Status &nbsp;</th>
	<th align="center">&nbsp; OCN &nbsp;</th>
   </tr>

<% Do Until RSTNPANXXMBIQry.EOF %>

  <tr align="center">
    <td>&nbsp;<%= RSTNPANXXMBIQry("Tix") %>&nbsp;</td>
    <td>&nbsp;<%= RSTNPANXXMBIQry("NPA") %>&nbsp;</td>
    <td>&nbsp;<%= RSTNPANXXMBIQry("NXX") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("MBI") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("COStatusDescription") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("OCN") %>&nbsp;</td>
  </tr>
<% RSTNPANXXMBIQry.moveNext
 loop %>
</p>
</table>



<p>Note: A ticket number of 999999999 implies a grandfathered MBI.</p>
 
<%
objConnection.close
objConnectionAvailable.close
 %>
</b>
</body>
</html>
