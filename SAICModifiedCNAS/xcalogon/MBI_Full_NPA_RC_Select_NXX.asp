<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#include file="xca_CNASLib.inc"-->
<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
Sub btnGoToMainFrm_onclick()
	Response.Redirect "xca_MenuMBI.asp"
End Sub

session("aNPA")=request.querystring("NPA")

session("aRC")=Replace(Request.querystring("RC"),"'","''")
aNPA=session("aNPA")
aRC=session("aRC")
SET objConnection = server.createobject("ADODB.connection")
SET rstQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlQry = "Exec [Look_For_MBI_Based_On_RC] " &  aNPA & ", '" & aRC & "'"
SET rstQry = objConnection.execute(sqlQry)
%> </p>
<% if (rstQry.EOF) then %><b></p>

<p>No available MBIs for <%= aRC %> in <%= aNPA%></b> </p>
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
<% Else %>
<p><br>
<table align="center" BORDER="1">

<tr>
    <th align="center">&nbsp; Tix &nbsp;</th>
	<th align="center">&nbsp; NPA &nbsp;</th>
    <th align="center">&nbsp; NXX &nbsp;</th>
	<th align="center">&nbsp; CO Code Status &nbsp;</th>
    <th align="center">&nbsp; Rate Centre &nbsp;</th>
    <th align="center">&nbsp; OCN &nbsp;</th>
    <th align="center">&nbsp; Entity Name &nbsp;</th>
	<th align="center">&nbsp; Available Blocks &nbsp;</th>
    </tr>

<% Do Until rstQry.EOF %>
  <tr align="left">
    <td align="center">&nbsp;<%= rstQry("Tix") %>&nbsp;</td>
	<td align="center">&nbsp;<%= rstQry("NPA") %>&nbsp;</td>
	<td align="center"><a HREF="MBI_Full_NPA_RC_Part1.asp?&NPA=<%= aNPA%>&NXX=<%= rstQry("NXX") %>"><%= rstQry("NXX") %></a></td>
    <td align="center">&nbsp;<%= rstQry("COStatusDescription") %>&nbsp;</td>
	<td align="left">&nbsp;<%= rstQry("RateCenter") %>&nbsp;</td>
	<td align="center">&nbsp;<%= rstQry("OCN") %>&nbsp;</td>
	<td align="left">&nbsp;<%= rstQry("EntityName") %>&nbsp;</td>
	 <td align="center">&nbsp;<%= rstQry("Counter") %>&nbsp;</td>
</tr>
<% rstQry.moveNext
 loop %>
</table>
<%End If%>
<%
objConnection.close %>
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
