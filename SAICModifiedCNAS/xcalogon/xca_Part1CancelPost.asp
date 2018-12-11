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
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<%

CanTix=int(session("P1CanTix"))
CanNPA=session("CanNPA")
CanNXX=session("CanNXX")

%>
</HEAD>
<BODY>


<P>&nbsp;</P>

<%
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
    objCmd.ActiveConnection = objConn	
	
'See if CO Code was reserved
SQLstmtR = "Select * from xca_Part1 where RequestStatus='RS' and NPA= '"&CanNPA&"' and NXX1Preferred= '"&CanNXX&"'"

		
	objCmd.CommandText=SQLStmtR
	set objRec=objCmd.Execute
	
	if not objRec.EOF then
		
		'WasReserved=true
		RetStatus="R"
		RetEntityID=objRec("EntityID")
		RetTix=objRec("Tix")
		LATA=objRec("LATA")
		OCN=objRec("OCN")
		SwitchID=objRec("SwitchID")
		WireCenter=objRec("WireCenter")
		
	else
		'WasReserved=false
		RetStatus="S"  
		RetEntityID=0
		RetTix=0
		LATA=""
		OCN=""
		SwitchID=""
		WireCenter=""				
	end if	


 'Update Part 1
 


SQLstmt = "Update xca_Part1 set RequestStatus= 'CC' where Tix='"&CanTix&"'"
	objCmd.CommandText=SQLStmt
	objCmd.Execute
	
	'Update COCodes
	
 


SQLstmt1="Update xca_COCode Set LATA='"&LATA&"', OCN='"&OCN&"', SwitchID='"&SwitchID&"', WireCenter='"&WireCenter&"', Status= '"&RetStatus&"', EntityID= '"&RetEntityID&"', Tix='"&RetTix&"' where NPA= '"&CanNPA&"' and NXX='"&CanNXX&"'"
	objCmd.CommandText=SQLStmt1
	objCmd.Execute
	
 



If  err.number>0 then
      response.write "VBScript Errors Occured:" & "<P>"
      response.write "Error Number=" & err.number & "<P>"
      response.write "Error Descr.=" & err.description & "<P>"
      response.write "Help Context=" & err.helpcontext & "<P>" 
      response.write "Help Path=" & err.helppath & "<P>"
      response.write "Native Error=" & err.nativeerror & "<P>"
      response.write "Source=" & err.source & "<P>"
      response.write "SQLState=" & err.sqlstate & "<P>"
end if

UserID= session("UserUserID")		
Action="Cancel"
ActionText="Cancelled"
twoEmail=(session("UserUserEmail")) & ", " & (session("EntityUserEmail"))
emailText="Ticket # " & CanTix & " for CO Code " & CanNPA & " " & CanNXX & " has been " & ActionText & ".  "

log  "R",CanNPA,CanNXX,UserID,Now,CanTix,Action,ActionText,"Part1"   
'
'This section was added by G. Brown Sep 7,1999.  S. Khare only wants an 'email sent when the applicant fills out the form on-line.
'
UserEntityType=session("UserEntityType")
UserUserType=session("UserUserType")
If UserEntityType <> "a" and UserUserType <> "a" then
'   
' This section was modified by G. Brown Oct 15 21,1999 to reflect a change in message format.
'
'email  session("AdminEntityEmail"),twoEmail,"","P1 Canceled",emailText
email "cnas@cnac.ca",twoEmail,"","Part One Cancelled",emailText
end if
session("P1CanEmail")=twoEmail
	


	Response.Redirect "xca_Part1CanConfirm.asp"



%>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>


