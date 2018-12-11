<%@ Language=VBScript %>

<%
Response.Buffer = true
Response.Expires=0
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
</form>
<!--#include file="xca_CNASLib.inc"-->

<form action="xca_MenuInt.asp" method="post" id="formP4" name="formAdminMenu" ONLOAD= history.go(0)">
<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>CNAS Administration Menu</title>

<SCRIPT LANGUAGE="JavaScript"><!--


 function go(url) {
     location.href = url;
 }
 

	
--></script>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnReturntoMain_onclick()
Response.Redirect "xca_MenuMainPost.asp"
End Sub


</SCRIPT>
</head>
<body text="black" bgProperties=fixed bgColor="#d7c7a4">
<%
session("NoTixSent")=""


session("HereP3")="xca_MenuCNASadmin.asp"
	%>

<P><center>
<font face="Arial Black" color=maroon size=5><strong>Security and System 
Administration Menu</font></STRONG></center>
<P></P>
<table align=center border=0 cellPadding=1 cellSpacing=1 width=35%>
    <TR>
        <TD width = 5 align=right><IMG height=36 src="../images/ball25.gif" width=36> 
        </TD>
        <TD colspan=3><FONT
    face=
        Arial><A href="javascript:go('xca_EntityOnly.asp')" target=page >Entity Administration</A>  </FONT>
       </TD>
    <TR>
        <TD width = 5 align=right><IMG height=36 src="../images/ball25.gif" 
            width=35 >   </TD>
        <TD colspan=3><FONT face=Arial><A href="javascript:go('xca_UserOnly.asp')" 
            target=page>User Administration</A></FONT>
        </TD>
        
    <TR>
        <TD width = 5 align=right><IMG height=36 src="../images/ball25.gif" width=36> </TD>
        <TD colspan=3><A href="javascript:go('xca_CleanUp.asp')"><FONT face=Arial >Database Clean Up</FONT></A>
        </TD>
        
    <TR>
        <TD width = 5 align=right><IMG height=36 src="../images/ball25.gif" width=36> </TD>
        <TD colspan=3><FONT face=Arial><FONT><A 
            href="javascript:go('xca_Parms.asp')" 
           >System 
            Parameters</A></FONT></FONT>
        </TR>
    <tr>
       <td>
         </td>
<td>
            <p>&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturntoMain 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 115px" width=115>
	<PARAM NAME="_ExtentX" VALUE="3043">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturntoMain">
	<PARAM NAME="Caption" VALUE="Return to Main">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturntoMain()
{
	btnReturntoMain.value = 'Return to Main';
	btnReturntoMain.setStyle(0);
}
function _btnReturntoMain_ctor()
{
	CreateButton('btnReturntoMain', _initbtnReturntoMain, null);
}
</script>
<% btnReturntoMain.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</p></td>
</tr>

 





<p>&nbsp;</p>             
</table> 
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
