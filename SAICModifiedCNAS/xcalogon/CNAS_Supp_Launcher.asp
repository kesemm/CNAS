<%@ Language=VBScript %>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: CNAS_Supp_Launcher.asp,v $
'* Commit Date:   $Date: 2015/01/19 13:34:14 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.4 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
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
RequestedPage=request.querystring("Page")
UserEntityType=session("UserEntityType")
UserUserType=session("UserUserType")
UserName=session("UserLogon")
UserID=session("UserUserID")

'If UserEntityType = "a" and UserUserType= "a" then

select case request.querystring("s")

case "1"
' REQUESTED CODE VALIDATOR APPLICATION
Response.Redirect("/CNAS_Supp/Default.aspx?Auth=cnas&UserID=" & UserID & "&Page=CodeValidator")


case "2"
' CNA DATE CALCULATOR APPLICATION
Response.Redirect("/CNAS_Supp/Default.aspx?Auth=cnas&UserID=" & UserID & "&Page=DateCalculator")


case "3"
' NONGEOGRAPHIC CODE LOOKUP
Response.Redirect("/CNAS_Supp/Default.aspx?Auth=cnas&UserID=" & UserID & "&Page=NonGeo_RqstViewCode")


case "4"
' NONGEOGRAPHIC FORMA FILL FORM
Response.Redirect("/CNAS_Supp/Default.aspx?Auth=cnas&UserID=" & UserID & "&Page=NonGeo_FormA_FillForm")


case "5"
' NONGEOGRAPHIC FORMC FILL FORM
Response.Redirect("/CNAS_Supp/Default.aspx?Auth=cnas&UserID=" & UserID & "&Page=NonGeo_FormC_FillForm")

end select

'Else

'SOMETHING WRONG HERE
'SEND TO DEFAULT PAGE TO GIVE THEM A CLOSE ME WINDOW

'Response.Redirect("/CNAS_Supp/Default.aspx")

'end if%>
<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
