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
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnGo_onclick()

	session("pPart4NPA")=txtNPA.value
	session("pPart4NXX")=txtNXX.value
	if session("pPart4NPA")<>"" and isnumeric(session("pPart4NPA")) and session("pPart4NXX")<>"" and isnumeric(session("pPart4NXX")) then
		if session("UserEntityType") = "a" then
			Result = NewPart4Data(session("pPart4NPA"),session("pPart4NXX"))
		else
			Result = NewPart4DataEntiy(session("pPart4NPA"),session("pPart4NXX"),session("UserEntityID"))
		end if
		 
		if not Result then
			session("part4_pre") ="no part4 record"
			Response.Redirect "xca_Part4Pre.asp"
		else
			session("part4_pre") =""
			session("Part4Act")=""
			Response.Redirect "xca_Part4.asp"
		end if	
	else
		session("part4_pre") ="invalid NPA-NXX number"
	end if
		
End Sub


function NewPart4Data(pNPA,pNXX)
	'check if record exists

	dim objConn
	dim objCmd
	dim objRec
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn

	objCmd.CommandText="Get_NewPart4_Data " & pNPA & ", " & pNXX 
			
	set objRec=objCmd.Execute
	if not objRec.EOF then
		NewPart4DataTemp=true
	else
		NewPart4DataTemp=false						
	end if	
				
	objRec.close
	objConn.close
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing
	
	NewPart4Data=NewPart4DataTemp
end function

function NewPart4DataEntiy(pNPA,pNXX,pEntity)
	'check if record exists

	dim objConn
	dim objCmd
	dim objRec
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn

	objCmd.CommandText="Get_NewPart4_Data_Entity " & pNPA & ", " & pNXX & ", " & pEntity 
			
	set objRec=objCmd.Execute
	if not objRec.EOF then
		NewPart4DataEntiyTemp=true
	else
		NewPart4DataEntiyTemp=false						
	end if	
				
	objRec.close
	objConn.close
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing
	
	NewPart4DataEntiy=NewPart4DataEntiyTemp
end function

</SCRIPT>
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
	if session("part4_pre") ="no part4 record" then
	
	%>
 		<SCRIPT Language="JavaScript">
		alert("No part 4 record exists for this NPA-NXX or the NPA-NXX does not belong to your entity.")
		</SCRIPT>

<%
	
	elseif session("part4_pre") ="invalid NPA-NXX number" then
	
	%>
 		<SCRIPT Language="JavaScript">
		alert("The NPA-NXX number is invalid or does not belong to your entity.")
		</SCRIPT>

<%	
	
	end if
	session("part4_pre") =""
%>
&nbsp;<BR><BR>
<P><center><font face="Arial Black" color=maroon size=5><strong>Input Part 
4</strong></font></center>
<P></P>
<P>&nbsp;<P>
<BR>
<TABLE WIDTH=60.71% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1 height=50 style="HEIGHT: 50px; WIDTH: 462px">
    
    <TR>
        <TD align=left vAlign=center><STRONG><FONT face=Arial size=2>Please enter the In-Service NPA -&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtNPA style="HEIGHT: 19px; LEFT: 10px; TOP: 56px; WIDTH: 30px" 
            width=30>
	<PARAM NAME="_ExtentX" VALUE="794">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="5">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNPA()
{
	txtNPA.setStyle(TXT_TEXTBOX);
	txtNPA.setMaxLength(3);
	txtNPA.setColumnCount(5);
}
function _txtNPA_ctor()
{
	CreateTextbox('txtNPA', _inittxtNPA, null);
}
</script>
<% txtNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 and NXX -&nbsp; 
            </FONT></STRONG>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtNXX style="HEIGHT: 19px; LEFT: 10px; TOP: 75px; WIDTH: 30px" 
            width=30>
	<PARAM NAME="_ExtentX" VALUE="794">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="5">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNXX()
{
	txtNXX.setStyle(TXT_TEXTBOX);
	txtNXX.setMaxLength(3);
	txtNXX.setColumnCount(5);
}
function _txtNXX_ctor()
{
	CreateTextbox('txtNXX', _inittxtNXX, null);
}
</script>
<% txtNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        <TD rowSpan=2>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnGo style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 40px" 
            width=40>
	<PARAM NAME="_ExtentX" VALUE="1058">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGo">
	<PARAM NAME="Caption" VALUE="Go ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGo()
{
	btnGo.value = 'Go ';
	btnGo.setStyle(0);
}
function _btnGo_ctor()
{
	CreateButton('btnGo', _initbtnGo, null);
}
</script>
<% btnGo.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
</TR>
</TABLE><BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</P>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
