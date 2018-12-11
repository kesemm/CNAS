<%@ Language=VBScript %>

<%
Response.Buffer = true
Response.Expires=0
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#Include file="xca_CNASlib.inc"-->

<HTML>
<HEAD>

<%
'Build some Page Variables from Session Variables
'================================================
Dim afname,asubject,abody,strResult
aNPA = session("selectedNPA")
aNXX = session("selectedNXX")
aRC=session("selectedRC")
afname="D:\CNA\Assign\"+aNPA+aNXX+".xls"
asubject="Assign check for NPA-NXX "+aNPA+"-"+aNXX
abody="The attached file is an assignment check for "+chr(13)+"NPA-NXX "+aNPA+"-"+aNXX+" for the Rate Centre of "+aRC+" "

'This ECHO's out the body of the E-Mail
'======================================
Set Executor = Server.CreateObject("ASPExec.Execute")
Executor.Application = "cmd"
Executor.Parameters = "/C echo " & abody & " > D:\CNA\Assign\assignbody.txt"
Executor.ShowWindow = False
strResult = Executor.ExecuteWinApp

'This BLAT's the Attachment
'==========================
Set Executor = Server.CreateObject("ASPExec.Execute")
Executor.Application = "cmd"
Executor.Parameters = "/C D:\CNA\Assign\assign.cmd "+aNPA+""+aNXX
Executor.ShowWindow = False
strResult = Executor.ExecuteWinApp

Response.Redirect "xca_MenuNANPCAN.asp"
%>

</HEAD>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>