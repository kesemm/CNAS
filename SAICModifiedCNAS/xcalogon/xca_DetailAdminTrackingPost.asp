<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post">
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
UserName=session("UserLogon")
If UserEntityType = "a" and UserUserType= "a" then
Session("lhd_ext_uid")=Session("UserUserLogon")
Session("aid")=Cint(Request.QueryString("aid"))
Response.Redirect("/CNATracking/detaillogon.asp")
Else
Response.Redirect("xca_Login2.asp")  
end if%>
<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
