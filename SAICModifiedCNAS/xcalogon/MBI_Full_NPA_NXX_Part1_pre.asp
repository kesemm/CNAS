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
session("aNPA")=Request.Form("NPA")
session("aNXX")=Request.Form("NXX")
aNPA=session("aNPA")
aNXX=session("aNXX")
SET objConnection = server.createobject("ADODB.connection")
SET rstCount =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlCountQry = "SELECT Count (*) As Number FROM xca_MBI WHERE xca_MBI.Status='S' And xca_MBI.NPA='" & aNPA &"' AND xca_MBI.NXX='" & aNXX & "';"
SET rstCount = objConnection.execute(sqlCountQry)
%> </p>
<% If rstCount("Number")=10 Then 
Response.Redirect("MBI_Full_NPA_NXX_Part1.asp") 
else
Response.Redirect("MBI_Full_Part1_NPA_NXX_Not_Available.asp")
end if%>


<%
objConnection.close %>
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
