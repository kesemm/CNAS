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

Response.Redirect("/mlm/Default.aspx?Page=MLM_Menu&Auth=cnas")

Else

'SOMETHING WRONG HERE
'SEND TO DEFAULT PAGE TO GIVE THEM A CLOSE ME WINDOW

Response.Redirect("/mlm/Default.aspx")

end if%>
<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
