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
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnOK_onclick()
Response.Redirect "xca_RptPrtsFrmsPre.asp"
End Sub


</SCRIPT>
</HEAD>
<body leftmargin=15 bgColor="#d7c7a4" bgProperties="fixed" text="black">
<TABLE align=center>
	<tr>
		<TD align=middle><IMG alt="" src="../images/stop.gif" 
            height=52 
            style="HEIGHT: 52px; WIDTH: 56px" width=56> 
		</TD>
	</tr>
	<tr>
		<td>
            <DIV align=center><font size=3 face=Arial color=black><Strong>
		No Part 4: Confirmation of Code Activation form has been 
            submitted for this ticket number. To view the contents of this 
            ticket, view Parts 1 and 3.</Strong></font></DIV>
		</td>
	</tr>
	<tr>
		<td align=center>
    <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
    id=btnOK style="HEIGHT: 27px; LEFT: 10px; TOP: 34px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="   OK   ">
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
	btnOK.value = '   OK   ';
	btnOK.setStyle(0);
}
function _btnOK_ctor()
{
	CreateButton('btnOK', _initbtnOK, null);
}
</script>
<% btnOK.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TABLE>
<P>&nbsp;</P>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
