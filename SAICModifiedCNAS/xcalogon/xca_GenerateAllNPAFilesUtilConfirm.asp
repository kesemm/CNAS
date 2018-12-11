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
<%
'Tix=session("P1TixCook")
'NPA=session("P1NPACook")
'NXX=session("P1NXXCook")
'twoEmail=session("P1TwoEmailsCook")
%>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnOK_onclick()
Response.Redirect "xca_MenuMainPost.asp"
End Sub


Sub Button1_onclick()
Response.Redirect "xca_WebNPAFilePre2m.asp"
End Sub

Sub Button2_onclick()
Response.Redirect "xca_WebNPAFileMenu.htm"
End Sub

</SCRIPT>
</HEAD>
<body leftmargin=15 bgColor="#d7c7a4" bgProperties="fixed" text="black">
<TABLE align=center WIDTH=65.02% BORDER=0 CELLSPACING=0 CELLPADDING=0 height=138 style="HEIGHT: 138px; WIDTH: 578px">
    
    <TR>
        <TD >&nbsp;&nbsp;&nbsp;&nbsp; 
    <TR>
        <TD align=middle><STRONG><FONT face=Arial size=4>You have 
            generated the Utilized NPA Files for all the NPAs </FONT></STRONG> 
    <TR>
        <TD > 
    <TR>
        <TD align=middle noWrap>
	<TR>
		<TD align=middle noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnOK style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 90px" 
            width=90>
	<PARAM NAME="_ExtentX" VALUE="2381">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="Main Menu">
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
	btnOK.value = 'Main Menu';
	btnOK.setStyle(0);
}
function _btnOK_ctor()
{
	CreateButton('btnOK', _initbtnOK, null);
}
</script>
<% btnOK.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=Button2 style="HEIGHT: 27px; LEFT: 90px; TOP: 0px; WIDTH: 194px" 
            width=194>
	<PARAM NAME="_ExtentX" VALUE="5133">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Button2">
	<PARAM NAME="Caption" VALUE="Generate NPA Files Menu">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initButton2()
{
	Button2.value = 'Generate NPA Files Menu';
	Button2.setStyle(0);
}
function _Button2_ctor()
{
	CreateButton('Button2', _initButton2, null);
}
</script>
<% Button2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=Button1 style="HEIGHT: 27px; LEFT: 284px; TOP: 0px; WIDTH: 175px" 
            width=175>
	<PARAM NAME="_ExtentX" VALUE="4630">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Button1">
	<PARAM NAME="Caption" VALUE="Generate All NPA Files">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initButton1()
{
	Button1.value = 'Generate All NPA Files';
	Button1.setStyle(0);
}
function _Button1_ctor()
{
	CreateButton('Button1', _initButton1, null);
}
</script>
<% Button1.display %>

<!--METADATA TYPE="DesignerControl" endspan--></TD>
	</TR>
    <TR>
        <TD >&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR>
</TABLE>

<P>&nbsp;</P>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
