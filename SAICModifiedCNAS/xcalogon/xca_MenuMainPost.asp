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
UserEntityType=session("UserEntityType")
UserUserType=session("UserUserType")
If UserEntityType = "a" and UserUserType= "a" then
		
        Response.Redirect("xca_MenuCNASadmin.asp")
        
Elseif UserEntityType = "a" and UserUserType= "m" then
		
		Response.Redirect ("xca_MenuCNASmgr.asp")
		
  else
  
    Response.Redirect("xca_MenuCNASapp.asp")   
    
end if%>
<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
