<script  language=jscript>
void function refCheck(){
	var re;
	re = new RegExp("cnas");
	gr = document.referrer;
	hist = history.length;
	loc= window.location;
	//history.go(-1);
	document.writeln(hist);
	document.writeln(loc);
	if (re.test(gr)==false) {
		//window.navigate("xca_Timeout.asp");
		document.writeln("false<br>"+gr);
	}
	document.writeln("true"+gr);
}
//refCheck();
</script>
<%



If session("UserEntityID") = ""   then
		
        Response.Redirect("..\xca_Timeout.asp")
 end if 

If session("UserEntityType") = "" then
		
        Response.Redirect("..\xca_Login2.asp")
	end if 


sub email(from,to1,cc,subject,body)
' Commented out by Glen Brown on May 31, 2002 to ensure mail is not sent
	'set objMail=CreateObject("CDONTS.Newmail")
	'objMail.From=from '"webmaster@cnas.com"
	'objMail.To=To1
	'objMail.Cc=cc
	'objMail.Subject=subject
	'objMail.Body=body
	'objMail.send
	'set objMail=nothing
end sub	
sub log(LogType,NPA,NXX,UserID,Date1,Tix,Action,ActionText,Process)
	Set objConn1=server.CreateObject("ADODB.Connection")
	Set objCmd1=server.CreateObject("ADODB.Command")
	objConn1.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd1.ActiveConnection = objConn1
	
	on error resume next
	objCmd1.CommandText=	"LogEvent '" & LogType & "', '" & NPA & "', '" & NXX & "', " _
										& UserID & ", '" & Date1 & "', " & Tix _
										& ", '" & Action & "', '" & ActionText & "', '" & Process & "'"
	objCmd1.Execute
	objConn1.close
	Set objConn1=Nothing
	Set objCmd1=Nothing
end sub

function IsDateReal(d)

	dim IsDateRealTemp
	
	dim NUM1,NUM2,NUM3
	
	if not isdate(d) then
	
		IsDateRealTemp=false
	
	else	
	
		NUM1 = cint(Instr(d,"/"))
		NUM2 = cint(Instr(Num1+1,d,"/"))
		NUM3 = Mid(d,NUM1+1,NUM2-NUM1-1)
		
		if cint(NUM3)<13 then
			
			IsDateRealTemp=true
		else
			IsDateRealTemp=false
			
		end if
		
	end if
	
	IsDateReal = IsDateRealTemp	
			
end function
%>
