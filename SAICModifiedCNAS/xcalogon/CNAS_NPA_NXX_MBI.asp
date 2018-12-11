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


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<p><%
Sub btnGoToMainFrm_onclick()
	Response.Redirect "xca_MenuMBI.asp"
End Sub

aNPA = request.querystring("NPA")
aNXX = request.querystring("NXX")
SET objConnection = server.createobject("ADODB.connection")
SET rstNPANXXMBIQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLNPANXXMBIQry = "SELECT Tix,NPA,NXX,MBI,COStatusDescription,RateCenter,EntityName,xca_MBI.OCN,PublicRemarks,CNARemarks FROM xca_MBI Left Join xca_status_codes ON xca_MBI.Status=xca_status_codes.COStatus Left JOIN xca_Entity ON xca_MBI.EntityID=xca_Entity.EntityID WHERE (((xca_MBI.NPA)='" & aNPA &"') AND ((xca_MBI.NXX)='" & aNXX & "')) ORDER BY MBI;"
SET RSTNPANXXMBIQry = objConnection.execute(SQLNPANXXMBIQry)
%> </p>
<% if RSTNPANXXMBIQry.EOF then %><b><p>No record found for:  <% = (aNPA) %> - <% = (aNXX) %> 
<%else%></p>

<p align="center"><strong>MBIs Assigned to <% = (aNPA) %> - <% = (aNXX) %> </strong></p></td>
<b>

<p><br>
<table align="center" BORDER="1">
  <tr>

 <tr>
    <th align="center">&nbsp; Tix &nbsp;</th>
    <th align="center">&nbsp; NPA &nbsp;</th>
    <th align="center">&nbsp; NXX &nbsp;</th>
    <th align="center">&nbsp; MBI &nbsp;</th>
    <th align="center">&nbsp; Status &nbsp;</th>
    <th align="center">&nbsp; Company &nbsp;</th>
    <th align="center">&nbsp; OCN &nbsp;</th>
	<th align="center">&nbsp; Rate Center &nbsp;</th>
	<th align="center">&nbsp; Public Remarks &nbsp;</th>
	<th align="center">&nbsp; CNA Remarks &nbsp;</th>
  </tr>

<% Do Until RSTNPANXXMBIQry.EOF %>

  <tr align="center">
    <td>&nbsp;<%= RSTNPANXXMBIQry("Tix") %>&nbsp;</td>
    <td>&nbsp;<%= RSTNPANXXMBIQry("NPA") %>&nbsp;</td>
    <td>&nbsp;<%= RSTNPANXXMBIQry("NXX") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("MBI") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("COStatusDescription") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("EntityName") %>&nbsp;</td>
	<td>&nbsp;<%= RSTNPANXXMBIQry("OCN") %>&nbsp;</td>
	<td nowrap>&nbsp;<%= RSTNPANXXMBIQry("RateCenter") %>&nbsp;</td>
	<td>&nbsp;<%= RSTNPANXXMBIQry("PublicRemarks") %>&nbsp;</td>
	<td>&nbsp;<%= RSTNPANXXMBIQry("CNARemarks") %>&nbsp;</td>
  </tr>
<% RSTNPANXXMBIQry.moveNext
 loop %>
</p>
</table>



<p>Note: A ticket number of 999999999 implies a grandfathered MBI.</p>
 <%end if%>
 <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoToMainFrm 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoToMainFrm">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoToMainFrm()
{
	btnGoToMainFrm.value = 'Return';
	btnGoToMainFrm.setStyle(0);
}
function _btnGoToMainFrm_ctor()
{
	CreateButton('btnGoToMainFrm', _initbtnGoToMainFrm, null);
}
</script>
<% btnGoToMainFrm.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<%
objConnection.close %>
</b>
</body>
</html>
