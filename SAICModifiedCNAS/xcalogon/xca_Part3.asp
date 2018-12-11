<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
</form>
<!--#include file="xca_CNASLib.inc"-->
<form action="xca_Part3int.asp" method="post" id="formP3" name="formP3app" onSubmit="return validateForm()">

<html>
<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

</head>

<title>Part 3 Form</title>

<script LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
 function checkdate(a) {						//a=document.frm.field.value

				var err=0,result
				
				if (a.length != 10) err=1
					d = a.substring(0, 2)// month
					c = a.substring(2, 3)// '/'
					b = a.substring(3, 5)// day
					e = a.substring(5, 6)// '/'
					f = a.substring(6, 10)// year
				if (b<1 || b>12) err = 1
				if (c != '/') err = 1
				if (d<1 || d>31) err = 1
				if (e != '/') err = 1
				if (f<1999) err = 1
				if (b==4 || b==6 || b==9 || b==11){
				if (d==31) err=1
				}
				if (b==2){
				var g=parseInt(f/4)
				if (isNaN(g)) {
				err=1
				}
				if (d>29) err=1
				if (d==29 && ((f/4)!=parseInt(f/4))) err=1
				}
				if (err==1) {
				return false;
				}
				else {
					return true;
			   }
		}  
function validDate2(effDate1Str,effDate2Str) //new 
		{
			var err=0
					//startDateStr to Date()
					monSt = effDate1Str.substring(0, 2)//  StartDate month 
					chr0 = effDate1Str.substring(2, 3)// '/'
					daySt = effDate1Str.substring(3, 5)//StartDate day
					chr1 = effDate1Str.substring(5, 6)// '/'
					yrSt = effDate1Str.substring(6, 10)//StartDate year
					//endDateStr to Date()
					monEf = effDate2Str.substring(0, 2)//EffectiveDate month
					chr0 = effDate2Str.substring(2, 3)// '/'
					dayEf = effDate2Str.substring(3, 5)//EffectiveDate day
					chr1 = effDate2Str.substring(5, 6)// '/'
					yrEf = effDate2Str.substring(6, 10)//EffectiveDate year
			effD1 = new Date(yrSt,monSt,daySt);
			effD1.setMonth(effD1.getMonth()-1);
			effD2 = new Date(yrEf,monEf,dayEf);
			effD2.setMonth(effD2.getMonth()-1);
			validEffDate = new Date();
			validEffDate.setTime(effD1.getTime());
			//validEffDate.setDate(validEffDate.getDate()+45);			
			result = Date.parse(validEffDate)- Date.parse(effD2)
			if ((result <= 0)) {
				err=0;
				//document.writeln("good validEffDate =>",	validEffDate);			
			}
			else {
				err=1;
				// document.writeln("bad validEffDate =>",	validEffDate);
			}
			if (err == 1){
				return false;
			}
			else {
				return true;
			}
		}       

function validDate3(effDate1Str,effDate2Str) //new 
		{
			var err=0
					//startDateStr to Date()
					daySt = effDate1Str.substring(0, 2)//  StartDate month 
					chr0 = effDate1Str.substring(2, 3)// '/'
					monSt = effDate1Str.substring(3, 5)//StartDate day
					chr1 = effDate1Str.substring(5, 6)// '/'
					yrSt = effDate1Str.substring(6, 10)//StartDate year
					//endDateStr to Date()
					dayEf = effDate2Str.substring(0, 2)//EffectiveDate month
					chr0 = effDate2Str.substring(2, 3)// '/'
					monEf = effDate2Str.substring(3, 5)//EffectiveDate day
					chr1 = effDate2Str.substring(5, 6)// '/'
					yrEf = effDate2Str.substring(6, 10)//EffectiveDate year
			effD1 = new Date(yrSt,monSt,daySt);
			effD1.setMonth(effD1.getMonth()-1);
			effD2 = new Date(yrEf,monEf,dayEf);
			effD2.setMonth(effD2.getMonth()-1);
			effD2.setDate(effD2.getDate()+ 180); // appDAte +180
			effD3 = new Date(yrEf,monEf,dayEf);
			effD3.setMonth(effD3.getMonth()-1);
			effD3.setDate(effD3.getDate()+540);// appDate +540
			validEffDate = new Date();
			validEffDate.setTime(effD1.getTime());
			//validEffDate.setDate(validEffDate.getDate()+45);	
			document.writeln();
			document.writeln();
					
			result = Date.parse(validEffDate)- Date.parse(effD2);
			document.writeln(effD1+"<-effD1  "+monSt+daySt+yrSt+"<-ed1  "+"  "+result+" <-r")
			result2 =Date.parse(validEffDate)- Date.parse(effD3);
			document.writeln(effD2+"<-effD2  "+monEf+dayEf+yrEf+"<-ed2  "+"  "+result2+" <-r2")
			document.writeln(effD3+"<-effD3  ")
			
			if ((result >= 0) && (result2 <= 0)) {
				err=0;
				//document.writeln("good validEffDate =>",	validEffDate);			
			}
			else {
				err=1;
				// document.writeln("bad validEffDate =>",	validEffDate);
			}
			if (err == 1){
				return false;
			}
			else {
				return true;
			}
		}       
 
function validateForm() {
     var yes = 0, rryes=0, cayes=0;
 //    tmpDate=document.formP3app.ExtentionDate.value;


 //    extDate = new Date(tmpDate);
     
     
		 if (document.formP3app.ExtentionDate.value != ""){
			 var result=checkdate(document.formP3app.ExtentionDate.value) 
			     if (result==false)	{
						alert("The Extention Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
				        document.formP3app.ExtentionDate.focus();
                return false;
                }
			}
	     
		
     
         if (eval("document.formP3app.Part3Result[0].checked") == true)
             yes++;
         if (eval("document.formP3app.Part3Result[1].checked") == true)
             yes++;
		 if (eval("document.formP3app.Part3Result[2].checked") == true)
             yes++;
         if (eval("document.formP3app.Part3Result[3].checked") == true)
             yes++;
             
     if  (yes==0){
    
                alert("You have not checked a Part 3 Action Type. Please select one and submit again");
                          
                return false;
     
     }
     
 
    if (document.formP3app.EffDate.value != ""){
	
          var result=checkdate(document.formP3app.EffDate.value) //START HERE MARTIN....RequestedEffDate
            if (result==false)	{
				alert("The NXX Effective Date field is invalid. Please type in a valid date (including leading zeros) and submit again");
               document.formP3app.EffDate.focus();
                return false;
                }
              }
     if (document.formP3app.RRComplete[1].checked)  {
				if (document.formP3app.RRDescription.value == ""){
						alert("Routing and Rating information is not complete. Please enter details in the 'Additional Information' field and submit again");
						document.formP3app.RRDescription.focus();
                return false;
             }
            }
        
	if (document.formP3app.LERGDate.value != "") {
               
      var result=checkdate(document.formP3app.LERGDate.value) 
            if (result==false)	{
				alert("The LERG Date field is invalid. Please type in a valid date (including leading zeros) and submit again");
                document.formP3app.LERGDate.focus();
                return false;
          }   
         }   
     if (document.formP3app.RRReturnDate.value != "") {
                
        var result=checkdate(document.formP3app.RRReturnDate.value) 
            if (result==false)	{
				alert("The Return Date field is invalid. Please type in a valid date (including leading zeros) and submit again");
                document.formP3app.RRReturnDate.focus();
                return false;
			           }  
                  }         
       

    if (document.formP3app.Part3Result[1].checked)  {
				if (document.formP3app.Part3IncompleteDescription.value == ""){
						alert("You have selected Part 3 Incomplete.  Please enter a description and submit again");
						document.formP3app.Part3IncompleteDescription.focus();
                return false;
             } 
             
           }  		
      if (document.formP3app.Part3Result[2].checked)  {
				if (document.formP3app.Part3DenialDescription.value == ""){
						alert("You have selected Part 3 Denied.  Please enter a description and submit again");
						document.formP3app.Part3DenialDescription.focus();
               return false;
             } 
             
           }  	
     if (document.formP3app.Part3Result[3].checked)  {
				if (document.formP3app.Part3SuspendedDescription.value == ""){
						alert("You have selected Part 3 Suspend.  Please enter a description and submit again");
						document.formP3app.Part3SuspendedDescription.focus();
               return false;
             } 
             
         }  			  
    }      
           
           
           
 

        
    

 
function NoTix(){
 
	alert("That Ticket does not exist Part3.  Please try again.....");

}
// end hiding -->
</script>


<%



session("Part3Complete") = ""
Sub btnGoToMainFrm_onclick()
	Response.Redirect "xca_MenuC0CReqAdmin.asp"
End Sub

'Session("Here")="xca_Pending.asp"
'session("HereP3")="xca_MenuC0CReqAdmin.asp"
'Session("HereP3")="xca_Pending.asp"
'Session("Pending")=session("Return")


uname = session("UserUserName")
TixManual=int(Request.Form("Tix"))
Tix=session("Tix")

if  session("Tix") =""  then
		Tix=TixManual
end if
session("P1P3COTix")=Tix



'Check for invalid tix
sqlnoTix="Select * from xca_Part1 where Tix= '"&Tix&"'"
	GetPart1Data.setSQLText(sqlnoTix)
	GetPart1Data.Open
	checkTIX= GetPart1Data.fields.getValue("Tix")
	P1EntityID= GetPart1Data.fields.getValue("EntityID")
	P1UserID= GetPart1Data.fields.getValue("UserID")
	P1P3CONPA= GetPart1Data.fields.getValue("NPA")
	P1RequestStatus= GetPart1Data.fields.getValue("RequestStatus")
if checkTix="" then	
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

end if

Select case P1RequestStatus

case "CA" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

case "CU" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

case "CC" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

case "CS" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

case "CD" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

case "CI" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

case "CP" 
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

end select

AdminEntityName=session("ADMIN")
session("NoTixSent")=""

AdminUserName=session("UserUserName")



session("P1EntityID")=P1EntityID
session("P1P3CONPA")=P1P3CONPA

'sql gets user entity recordset of user
sql = "SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = '"&Tix&"' AND xca_User.UserID = '"&P1UserID&"' AND xca_Entity.EntityID = '"&P1EntityID&"'"
'sql = "Select * from xca_Entity where EntityID = '"&P1EntityID&"'"
	GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open

''sql 1 gets the entity recordset of the ADMIN
sql1="Select EntityID from xca_Entity where EntityName= '"&AdminEntityID&"'"
	
'sql2 gets part1 recordset using tix from previous page
sql2 = "Select * from xca_Part1 where Tix= '"&Tix&"'"
	GetPart1Data.setSQLText(sql2)
	GetPart1Data.Open
'Used to pass The ApplicationDate field to the p3int
	session("ApplicationDate") = GetPart1Data.fields.getValue("ApplicationDate")	

AdminData=session("ADMIN")

'get Admin info for top of form
sqlADMIN="Select * from xca_Entity where EntityName ='"&AdminData&"'"
	GetAdminEntityName.setSQLText(sqlADMIN)
	GetAdminEntityName.Open


sqlrsv="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'R'"
	ToRRsvRec.setSQLText(sqlrsv)
	ToRRsvRec.open

sqlass="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'A'"
	ToRAssRec.setSQLText(sqlass)
	ToRAssRec.open

effDatee=GetPart1Data.fields.getValue("RequestedEffDate")
datee = datevalue(effDatee)
LERGDatee = DateAdd("d",-45,datee)
NXX2=GetPart1Data.fields.getValue("NXX2")
NXX3=GetPart1Data.fields.getValue("NXX3")
'NXX4=GetPart1Data.fields.getValue("NXX4")
'NXX5=GetPart1Data.fields.getValue("Nxx5")
NoNXX1=GetPart1Data.fields.getValue("NoNXX1")
NoNXX2=GetPart1Data.fields.getValue("NoNXX2")
NoNXX3=GetPart1Data.fields.getValue("NoNXX3")
NoNXX4=GetPart1Data.fields.getValue("NoNXX4")
NoNXX5=GetPart1Data.fields.getValue("NoNXX5")


NXXsql="Select * From xca_COCode Where NPA= '"&P1P3CONPA&"' and NXX= '"&NXX2&"' and Status='S'"
	GetCOCodeData.setSQLText(NXXsql)
	GetCOCodeData.open
	isNXX2= GetCOCodeData.fields.getValue("NXX")

	
NXXsql="Select * From xca_COCode Where NPA= '"&P1P3CONPA&"' and NXX= '"&NXX3&"' and Status='S'"
	GetCOCodeData.setSQLText(NXXsql)
	GetCOCodeData.requery
	isNXX3= GetCOCodeData.fields.getValue("NXX")

	
'NXXsql="Select * From xca_COCode Where NPA= '"&P1P3CONPA&"' and NXX= '"&NXX4&"' and Status='S'"
'	GetCOCodeData.setSQLText(NXXsql)
'	GetCOCodeData.requery
'	isNXX4= GetCOCodeData.fields.getValue("NXX")
	
	
'NXXsql="Select * From xca_COCode Where NPA= '"&P1P3CONPA&"' and NXX= '"&NXX5&"' and Status='S'"
''	GetCOCodeData.setSQLText(NXXsql)
'	GetCOCodeData.requery
'	isNXX5= GetCOCodeData.fields.getValue("NXX")
	
	
	P3data=	"Select * from xca_Part3 where Tix= '"&Tix&"' "
	P3ExtentionDate.setSQLText(P3data)
	P3ExtentionDate.open
'	Response.Write P3ExtentionDate.fields.getValue("ExtentionDate")&"<-da real"

	IF P3ExtentionDate.fields.getValue("ExtentionDate")<> "01/01/1900" then
		P3ExtentionDateVal=P3ExtentionDate.fields.getValue("ExtentionDate")
'		Response.Write P3ExtentionDateVal &"<- not null(01/01/1900)<br>"
	else 
		P3ExtentionDateVal=""
'		Response.Write P3ExtentionDateVal&"<-  null(01/01/1900)<br>"
	end if
	
RequestStatusValue=GetPart1Data.fields.getValue("RequestStatus")
Select Case RequestStatusValue
Case "NW"
	RequestStatuschar="Pending - New Code"
Case "UP"
	RequestStatuschar="Pending - Update"
Case "AS"
	RequestStatuschar="Pending - Assigned"
Case "RS"
	RequestStatuschar="Pending - Reserved"
Case "CU"
	RequestStatuschar="Closed - Updated"
Case "CD"
	RequestStatuschar="Closed - Denied"
Case "CI"
	RequestStatuschar="Closed - Incomplete"
Case "CP"
	RequestStatuschar="Closed - Suspended"
Case "CS"
	RequestStatuschar="Closed - InService"
Case "CA"
	RequestStatuschar="Closed - Assigned"
Case "CC"
	RequestStatuschar="Closed - Cancelled by Code Applicant"
End Select

	

CodeReqNew=GetPart1Data.fields.getValue("CodeRequestNew")
Select Case CodeReqNew
Case "c"
	CodeReqNewchar1="**"
Case "o"
	CodeReqNewchar2="**"
End Select


CertReqValue=GetPart1Data.fields.getValue("CertificationRequired")
Select Case CertReqValue
Case "Y"
	CertReqchar="YES"
Case "N"
	CertReqchar="NO"
End Select


ReqCertReadyValue=GetPart1Data.fields.getValue("RequiredCertificationReady")
Select Case ReqCertReadyValue
Case "Y"
	ReqCertReadychar="YES"
Case "N"
	ReqCertReadychar="NO"
End Select


TypeEntityValue=GetPart1Data.fields.getValue("CarrierType")
Select Case TypeEntityValue
Case "l"
	TypeEntitychar="Local Exchange Carrier"
Case "w"
	TypeEntitychar="Wireless Service Provider"
Case "o"
	TypeEntitychar="Other(Specify)"
End Select


AuthPart2Value=GetPart1Data.fields.getValue("AuthorizationPart2")
Select Case AuthPart2Value
Case "y"
	AuthPart2char1="**"
Case "n"
	AuthPart2char2="**"
End Select
ReserveDate=Date()
ReserveDAte450=date() + 450
EffDateLabel="dd/mm/ccyy"

TyReqvalue=GetPart1Data.fields.getValue("TypeOfRequest")
session("P1TypeofRequest")=TyReqvalue
Select Case TyReqvalue
Case "A"
	TyReqchar1="**"
	UpdatedNPA.hide
	NXXUpdate.hide
	ReservedNPA.hide
	ReservedNXX.hide
	ReserveDate=""
	ReserveDate450=""
	ReservedSwitchID.hide
	AisNXX2=isNXX2
	AisNXX3=isNXX3
'	AisNXX4=isNXX4
'	AisNXX5=isNXX5
	RisNXX2=""
	RisNXX3=""
'	RisNXX4=""
'	RisNXX5=""
	AnisNXX1=NoNXX1
	AnisNXX2=NoNXX2
	AnisNXX3=NoNXX3
	AnisNXX4=NoNXX4
	AnisNXX5=NoNXX5
	RnisNXX1=""
	RnisNXX2=""
	RnisNXX3=""
	RnisNXX4=""
	RnisNXX5=""
	
	
	
Case "U"
	TyReqchar2="**"
	AssignedNPA.hide
	AssignedNXX.hide
	ReserveDate=""
	ReserveDate450=""
	ReservedSwitchID.hide
	SwitchID.hide
	RateCenter.hide
	ReservedNPA.hide
	ReservedNXX.hide
	ReserveDate=""
	ReserveDate450=""
	ReservedSwitchID.hide
	AisNXX2=""
	AisNXX3=""
'	AisNXX4=""
'	AisNXX5=""
	RisNXX2=""
	RisNXX3=""
	RisNXX4=""
	RisNXX5=""
	RnisNXX1=""
	RnisNXX2=""
	RnisNXX3=""
'	RnisNXX4=""
'	RnisNXX5=""
	AnisNXX1=""
	AnisNXX2=""
	AnisNXX3=""
	AnisNXX4=""
	AnisNXX5=""
	
	
	
Case "R"
	TyReqchar3="**"
	AssignedNPA.hide
	AssignedNXX.hide
	UpdatedNPA.hide
	NXXUpdate.hide
	SwitchID.hide
	RateCenter.hide
	RisNXX2= isNXX2
	RisNXX3= isNXX3
'	RisNXX4= isNXX4
'	RisNXX5= isNxx5
	AisNXX2=""
	AisNXX3=""
'	AisNXX4=""
'	AisNXX5=""
	RnisNXX1=NoNXX1
	RnisNXX2=NoNXX2
	RnisNXX3=NoNXX3
	RnisNXX4=NoNXX4
	RnisNXX5=NoNXX5
	AnisNXX1=""
	AnisNXX2=""
	AnisNXX3=""
	AnisNXX4=""
	AnisNXX5=""
	
	
End Select


Reas4ReqValue=GetPart1Data.fields.getValue("ReasonForRequest")
Select Case Reas4ReqValue
Case "aic"
	Reas4Reqchar="a) Initial Code for new Switching Entity or new Point of Interconnection (Complete Part 2)"
Case "aau"
	Reas4Reqchar="b) Code request for New Application for existing switching entity or point of interconnection (Code Aplicant must complete Section 1.7)"
Case "aag"
	Reas4Reqchar="c) Additional Code for Growth (Code Applicant must complete Section 1.6)"
End Select


ReasForReqValue=GetPart1Data.fields.getValue("ReasonForRequest")
Select Case ReasForReqValue
Case "ric"
	ReasForReqchar="a) Initial Code"
Case "rau"
	ReasForReqchar="b) New Application (Complete Section 1.7)"
Case "rag"
	ReasForReqchar="c) Growth (Complete Section 1.6)"
End Select


JeopardyValue = GetPart1Data.fields.getValue("NPAinJeopardy")
select case JeopardyValue
case "y"
	JeopardyName1="YES"
	JeopardyNameP3="YES"
case "n"
	JeopardyName2="NO"
	JeopardyNameP3="NO"
case else
	JeopardyNameP3="NO"
end select

sqlParm = "Select * from xca_Parms where name='P1DAYS'"

	P1Parms.setSQLText(sqlParm)
	P1Parms.Open
	P1getDays= P1Parms.fields.getValue("Value")
	Part1Days.setCaption(P1getDays)



%>

<script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

Sub btnSubmit_onclick()
Response.Redirect "xca_Part3int.asp"
End Sub


Sub btnReturn_onclick()

Response.Redirect "xca_MenuC0CReqAdmin.asp"

'HereP3=Session("HereP3")
'Here=Session("Here")
'If Session("Here")="" then
'	Response.Redirect "xca_MenuC0CReqAdmin.asp"
'End If

'If Session("HereP3")="" then
'	Response.Redirect "xca_Pending.asp"
'end if 

'If HereP3<>"" then
'Response.Write HereP3
'Response.Write Here

	'Response.Redirect HereP3
'else
	'Response.Redirect Here
'end if 
End Sub

Sub btnPending_onclick()
	Response.Redirect "xca_Pending.asp"
End Sub
</script>
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black" leftmargin="20" rightmargin="20">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=P1Parms style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sValue\sfrom\sxca_Parms\swhere\sName=?\q,TCControlID_Unmatched=\qP1Parms\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sValue\sfrom\sxca_Parms\swhere\sName=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initP1Parms()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select Value from xca_Parms where Name=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	P1Parms.setRecordSource(rsTmp);
	if (thisPage.getState('pb_P1Parms') != null)
		P1Parms.setBookmark(thisPage.getState('pb_P1Parms'));
}
function _P1Parms_ctor()
{
	CreateRecordset('P1Parms', _initP1Parms, null);
}
function _P1Parms_dtor()
{
	P1Parms._preserveState();
	thisPage.setState('pb_P1Parms', P1Parms.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserEntityName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part1\sWHERE\sxca_Part1.Tix\s=\s?\sAND\sxca_User.UserName\s=\s?\sAND\sxca_Entity.EntityID\s=\s?\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part1\sWHERE\sxca_Part1.Tix\s=\s?\sAND\sxca_User.UserName\s=\s?\sAND\sxca_Entity.EntityID\s=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserEntityName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = ? AND xca_User.UserName = ? AND xca_Entity.EntityID = ?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetUserEntityName.setRecordSource(rsTmp);
}
function _GetUserEntityName_ctor()
{
	CreateRecordset('GetUserEntityName', _initGetUserEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_Parms\sWHERE\s(xca_Entity.EntityName\s=\s?)\q,TCControlID_Unmatched=\qGetAdminEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_Parms\sWHERE\s(xca_Entity.EntityName\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminEntityName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_Parms WHERE (xca_Entity.EntityName = ?)';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetAdminEntityName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetAdminEntityName') != null)
		GetAdminEntityName.setBookmark(thisPage.getState('pb_GetAdminEntityName'));
}
function _GetAdminEntityName_ctor()
{
	CreateRecordset('GetAdminEntityName', _initGetAdminEntityName, null);
}
function _GetAdminEntityName_dtor()
{
	GetAdminEntityName._preserveState();
	thisPage.setState('pb_GetAdminEntityName', GetAdminEntityName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminUserName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sUserName\sfrom\sxca_User\swhere\sEntityID=?\q,TCControlID_Unmatched=\qGetAdminUserName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_User\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sUserName\sfrom\sxca_User\swhere\sEntityID=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminUserName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select UserName from xca_User where EntityID=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetAdminUserName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetAdminUserName') != null)
		GetAdminUserName.setBookmark(thisPage.getState('pb_GetAdminUserName'));
}
function _GetAdminUserName_ctor()
{
	CreateRecordset('GetAdminUserName', _initGetAdminUserName, null);
}
function _GetAdminUserName_dtor()
{
	GetAdminUserName._preserveState();
	thisPage.setState('pb_GetAdminUserName', GetAdminUserName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart1Data style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxcaPart1.Tix\s=?\q,TCControlID_Unmatched=\qGetPart1Data\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxcaPart1.Tix\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart1Data()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from xca_part1 where xcaPart1.Tix =?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart1Data.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetPart1Data') != null)
		GetPart1Data.setBookmark(thisPage.getState('pb_GetPart1Data'));
}
function _GetPart1Data_ctor()
{
	CreateRecordset('GetPart1Data', _initGetPart1Data, null);
}
function _GetPart1Data_dtor()
{
	GetPart1Data._preserveState();
	thisPage.setState('pb_GetPart1Data', GetPart1Data.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=ToRAssRec style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'A'\q,TCControlID_Unmatched=\qToRAssRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'A'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initToRAssRec()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'SELECT * FROM xca_Part1 WHERE TypeOfRequest = \'A\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	ToRAssRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_ToRAssRec') != null)
		ToRAssRec.setBookmark(thisPage.getState('pb_ToRAssRec'));
}
function _ToRAssRec_ctor()
{
	CreateRecordset('ToRAssRec', _initToRAssRec, null);
}
function _ToRAssRec_dtor()
{
	ToRAssRec._preserveState();
	thisPage.setState('pb_ToRAssRec', ToRAssRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 id=Part3_connect 
	style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\s*\sfrom\sxca_COCodes\swhere\sStatus=?\q,TCControlID_Unmatched=\qPart3_connect\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part3\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\s*\sfrom\sxca_COCodes\swhere\sStatus=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initPart3_connect()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'select * from xca_COCodes where Status=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Part3_connect.setRecordSource(rsTmp);
	if (thisPage.getState('pb_Part3_connect') != null)
		Part3_connect.setBookmark(thisPage.getState('pb_Part3_connect'));
}
function _Part3_connect_ctor()
{
	CreateRecordset('Part3_connect', _initPart3_connect, null);
}
function _Part3_connect_dtor()
{
	Part3_connect._preserveState();
	thisPage.setState('pb_Part3_connect', Part3_connect.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=ToRRsvRec style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'R'\q,TCControlID_Unmatched=\qToRRsvRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'R'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initToRRsvRec()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'SELECT * FROM xca_Part1 WHERE TypeOfRequest = \'R\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	ToRRsvRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_ToRRsvRec') != null)
		ToRRsvRec.setBookmark(thisPage.getState('pb_ToRRsvRec'));
}
function _ToRRsvRec_ctor()
{
	CreateRecordset('ToRRsvRec', _initToRRsvRec, null);
}
function _ToRRsvRec_dtor()
{
	ToRRsvRec._preserveState();
	thisPage.setState('pb_ToRRsvRec', ToRRsvRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 id=GetCOCodeData 
	style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 462px" width=462>
	<PARAM NAME="ExtentX" VALUE="12223">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qTables\q,TCDBObjectName_Unmatched=\q\q,TCControlID_Unmatched=\qGetCOCodeData\q,TCPPConn=\qcnasadmin\q,RCDBObject=\qRCDBObject\q,TCPPDBObject_Unmatched=\qTables\q,TCPPDBObjectName_Unmatched=\q\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetCOCodeData()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
//Recordset DTC error: Failed to get command text
	cmdTmp.CommandText = '';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetCOCodeData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetCOCodeData') != null)
		GetCOCodeData.setBookmark(thisPage.getState('pb_GetCOCodeData'));
}
function _GetCOCodeData_ctor()
{
	CreateRecordset('GetCOCodeData', _initGetCOCodeData, null);
}
function _GetCOCodeData_dtor()
{
	GetCOCodeData._preserveState();
	thisPage.setState('pb_GetCOCodeData', GetCOCodeData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=P3ExtentionDate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sExtentionDate\sfrom\sxca_Part3\s\q,TCControlID_Unmatched=\qP3ExtentionDate\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sExtentionDate\sfrom\sxca_Part3\s\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initP3ExtentionDate()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select ExtentionDate from xca_Part3 ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	P3ExtentionDate.setRecordSource(rsTmp);
	if (thisPage.getState('pb_P3ExtentionDate') != null)
		P3ExtentionDate.setBookmark(thisPage.getState('pb_P3ExtentionDate'));
}
function _P3ExtentionDate_ctor()
{
	CreateRecordset('P3ExtentionDate', _initP3ExtentionDate, null);
}
function _P3ExtentionDate_dtor()
{
	P3ExtentionDate._preserveState();
	thisPage.setState('pb_P3ExtentionDate', P3ExtentionDate.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<table align="left" border="0" cellPadding="0" cellSpacing="0" width="174" background height="48" style="HEIGHT: 48px; WIDTH: 174px">
    
    <tr>
        <td align="left">&nbsp;
        <td align="left"><strong><font face="Arial" color="maroon" size="4">
	CNA 
            Ticket #:&nbsp;&nbsp;
            <% Response.Write(Tix) %></font></strong></td>
		</td>
	</tr>
	<tr>
		<td align="left">&nbsp;
		<td align="left"><font color="maroon" size="4"><strong>
	Request 
            Status:&nbsp;&nbsp;
            <% Response.write RequestStatuschar %></strong></font>
		</td>
	</tr>
</table>&nbsp;
<p>&nbsp;</p>


<table border="0" cellpadding="0"><tr>
        <td>
	<td><font color="maroon" face="Arial Black" size="4"><strong>
Part 1 - 
            Canadian Central Office Code (NXX) Assignment Request 
            Form</strong></font> 
    
		</td>
	</tr>
</table>
<font face="arial" size="2">

<p>Please complete the following form. Use one form per NXX 
code request. Mail, fax, or submit online the completed form to the Code 
Administrator.</p>
<p>The Code Applicants are granted subject to the condition 
that all code holders are subject to the assignment guidelines which are 
published and available from the appropriate Code Administrator. A code assigned 
to an entity, either directly by the Code Administrator or through transfer from 
another entity, should be placed in service within 6 months after the initially 
published effective date.</p>
<p>These guidelines may be modified from time-to-time. The 
assignment guidelines in effect shall apply equally to all Code Applicants and 
all existing code holders.</p> 
<p>The Code Applicant and the Code Administrator 
acknowledge that the information contained on this request form is sensitive and 
will be treated as confidential. Prior to confirmation the information in this 
form will only be shared with the appropriate administrator and/or regulators. 
Information requested for RDBS and BRIDS will become available to the public 
upon input into those systems.</p>
<p>I hereby certify that the following information 
requesting an NXX code is true and accurate to the best of my knowledge and that 
this application has been prepared in accordance with the Canadian Central 
Office Code (NXX) Assignment Guidelines dated October 23, 1997 which were 
adopted by the CSCN on April 2, 1998.</p>
<p>It is understood that the Code Applicant will return the 
CO Code to the administrator for reassignment if the resource is no longer in 
use by the Code Applicant, no longer required for the service for which it was 
intended, not activated within the time frame specified in these guidelines (an 
extension can be applied for), or not used in conformance with these assignment 
guidelines.</p></font>
<p>
<br>
<table align="left" border="0" cellPadding="0" cellSpacing="0">
<tr>
<td>
<strong><font size="2" face="arial"><strong>Code Applicants are required to retain a copy of all 
            application forms, appendices and supporting data in the event of an 
            audit.</strong></font>
            </strong></td></tr>
</table>
<br>
<br>
<br>

<table align="center" border="0" cellPadding="0" cellSpacing="0">
<tr>
<td align="right"><label><font face="arial" size="2"><strong>Authorized Representative 
            Name:&nbsp;&nbsp;</strong></font></label></td>
<td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AuthorizedRep 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 98px" width=98>
	<PARAM NAME="_ExtentX" VALUE="2593">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AuthorizedRep">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AuthorizedRep">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRep()
{
	AuthorizedRep.setDataSource(GetPart1Data);
	AuthorizedRep.setDataField('AuthorizedRep');
}
function _AuthorizedRep_ctor()
{
	CreateLabel('AuthorizedRep', _initAuthorizedRep, null);
}
</script>
<% AuthorizedRep.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
<tr>
<td align="right"><label><font face="arial" size="2"><strong>Title:&nbsp;&nbsp;</strong></font></label></td>
<td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AuthorizedRepTitle 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 126px" width=126>
	<PARAM NAME="_ExtentX" VALUE="3334">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AuthorizedRepTitle">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AuthorizedRepTitle">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRepTitle()
{
	AuthorizedRepTitle.setDataSource(GetPart1Data);
	AuthorizedRepTitle.setDataField('AuthorizedRepTitle');
}
function _AuthorizedRepTitle_ctor()
{
	CreateLabel('AuthorizedRepTitle', _initAuthorizedRepTitle, null);
}
</script>
<% AuthorizedRepTitle.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
<tr>
<td align="right"><label><font face="arial" size="2"><strong>Date of 
            Receipt:&nbsp;&nbsp;</strong></font></label></td>
<td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=DateofReceipt1 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofReceipt1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateofReceipt">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofReceipt1()
{
	DateofReceipt1.setDataSource(GetPart1Data);
	DateofReceipt1.setDataField('DateofReceipt');
}
function _DateofReceipt1_ctor()
{
	CreateLabel('DateofReceipt1', _initDateofReceipt1, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("DateofReceipt"),vbShortDate)
'DateofReceipt1.display
%>
<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
<tr>
<td align="right"><label><font face="arial" size="2"><strong>Last Correspondence Date:&nbsp;&nbsp;</strong></font></label></td>
<td align="left">
 <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=CorrespondenceDate 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 116px" width=116>
	<PARAM NAME="_ExtentX" VALUE="3069">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CorrespondenceDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CorrespondenceDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCorrespondenceDate()
{
	CorrespondenceDate.setDataSource(GetPart1Data);
	CorrespondenceDate.setDataField('CorrespondenceDate');
}
function _CorrespondenceDate_ctor()
{
	CreateLabel('CorrespondenceDate', _initCorrespondenceDate, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("CorrespondenceDate"),vbShortDate)
'CorrespondenceDate.display
%>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            

</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
</table>
<br><br>
<br><br>
<strong><center><font size="4" face="arial" color="maroon">General Information</font></strong></center>
<table align="left" border="0" cellPadding="0" cellSpacing="1">
<tr>
        <td style="FONT-WEIGHT: bold"><label><strong><font size="3" face="arial" color="maroon">1.1 Contact 
            Information:</font></strong></label> 
 
 </td></tr>
 
 </table>
 <br>
 <br>


<table align="center" border="0" cellPadding="1" cellSpacing="1">
    <tbody>
    
    <tr>
        <td align="left" colSpan="2">
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code 
            Applicant Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left"><font face="Arial"> </font>
        <td align="left" colSpan="2">
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Entity 
            Name</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityname 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 72px" width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityname">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityname()
{
	AppEntityname.setDataSource(GetUserEntityName);
	AppEntityname.setDataField('EntityName');
}
function _AppEntityname_ctor()
{
	CreateLabel('AppEntityname', _initAppEntityname, null);
}
</script>
<% AppEntityname.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>

</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"> 
            </font></font>CNA Admin</font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><strong><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CO Code Manager</font></strong><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>

</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>

</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityContact 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityContact">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityContact()
{
	AppEntityContact.setDataSource(GetUserEntityName);
	AppEntityContact.setDataField('UserName');
}
function _AppEntityContact_ctor()
{
	CreateLabel('AppEntityContact', _initAppEntityContact, null);
}
</script>
<% AppEntityContact.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityContact 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 83px" width=83>
	<PARAM NAME="_ExtentX" VALUE="2196">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityContact()
{
	AdminEntityContact.setDataSource(GetAdminEntityName);
	AdminEntityContact.setDataField('EntityContact');
}
function _AdminEntityContact_ctor()
{
	CreateLabel('AdminEntityContact', _initAdminEntityContact, null);
}
</script>
<% AdminEntityContact.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserAddress 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserAddress">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserAddress()
{
	AppUserAddress.setDataSource(GetUserEntityName);
	AppUserAddress.setDataField('UserAddress');
}
function _AppUserAddress_ctor()
{
	CreateLabel('AppUserAddress', _initAppUserAddress, null);
}
</script>
<% AppUserAddress.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityAddress 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityAddress()
{
	AdminEntityAddress.setDataSource(GetAdminEntityName);
	AdminEntityAddress.setDataField('EntityAddress');
}
function _AdminEntityAddress_ctor()
{
	CreateLabel('AdminEntityAddress', _initAdminEntityAddress, null);
}
</script>
<% AdminEntityAddress.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserCity 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserCity">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserCity()
{
	AppUserCity.setDataSource(GetUserEntityName);
	AppUserCity.setDataField('UserCity');
}
function _AppUserCity_ctor()
{
	CreateLabel('AppUserCity', _initAppUserCity, null);
}
</script>
<% AppUserCity.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityCity 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityCity()
{
	AdminEntityCity.setDataSource(GetAdminEntityName);
	AdminEntityCity.setDataField('EntityCity');
}
function _AdminEntityCity_ctor()
{
	CreateLabel('AdminEntityCity', _initAdminEntityCity, null);
}
</script>
<% AdminEntityCity.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
            
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserProvince 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 87px" width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserProvince">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserProvince()
{
	AppUserProvince.setDataSource(GetUserEntityName);
	AppUserProvince.setDataField('UserProvince');
}
function _AppUserProvince_ctor()
{
	CreateLabel('AppUserProvince', _initAppUserProvince, null);
}
</script>
<% AppUserProvince.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityProvince 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 87px" width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityProvince()
{
	AdminEntityProvince.setDataSource(GetAdminEntityName);
	AdminEntityProvince.setDataField('EntityProvince');
}
function _AdminEntityProvince_ctor()
{
	CreateLabel('AdminEntityProvince', _initAdminEntityProvince, null);
}
</script>
<% AdminEntityProvince.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
         
            
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal 
            Code</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserPostalCode 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserPostalCode">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserPostalCode()
{
	AppUserPostalCode.setDataSource(GetUserEntityName);
	AppUserPostalCode.setDataField('UserPostalCode');
}
function _AppUserPostalCode_ctor()
{
	CreateLabel('AppUserPostalCode', _initAppUserPostalCode, null);
}
</script>
<% AppUserPostalCode.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal Code 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityPostalCode 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityPostalCode()
{
	AdminEntityPostalCode.setDataSource(GetAdminEntityName);
	AdminEntityPostalCode.setDataField('EntityPostalCode');
}
function _AdminEntityPostalCode_ctor()
{
	CreateLabel('AdminEntityPostalCode', _initAdminEntityPostalCode, null);
}
</script>
<% AdminEntityPostalCode.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
           
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail 
            Address 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserEmail 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 64px" width=64>
	<PARAM NAME="_ExtentX" VALUE="1693">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserEmail">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserEmail()
{
	AppUserEmail.setDataSource(GetUserEntityName);
	AppUserEmail.setDataField('UserEmail');
}
function _AppUserEmail_ctor()
{
	CreateLabel('AppUserEmail', _initAppUserEmail, null);
}
</script>
<% AppUserEmail.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityEmail 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 71px" width=71>
	<PARAM NAME="_ExtentX" VALUE="1879">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityEmail()
{
	AdminEntityEmail.setDataSource(GetAdminEntityName);
	AdminEntityEmail.setDataField('EntityEmail');
}
function _AdminEntityEmail_ctor()
{
	CreateLabel('AdminEntityEmail', _initAdminEntityEmail, null);
}
</script>
<% AdminEntityEmail.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>    
            
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserFax 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 53px" width=53>
	<PARAM NAME="_ExtentX" VALUE="1402">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserFax">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserFax()
{
	AppUserFax.setDataSource(GetUserEntityName);
	AppUserFax.setDataField('UserFax');
}
function _AppUserFax_ctor()
{
	CreateLabel('AppUserFax', _initAppUserFax, null);
}
</script>
<% AppUserFax.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityFax 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityFax()
{
	AdminEntityFax.setDataSource(GetAdminEntityName);
	AdminEntityFax.setDataField('EntityFax');
}
function _AdminEntityFax_ctor()
{
	CreateLabel('AdminEntityFax', _initAdminEntityFax, null);
}
</script>
<% AdminEntityFax.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
          
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserTelephone 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 90px" width=90>
	<PARAM NAME="_ExtentX" VALUE="2381">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserTelephone">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserTelephone()
{
	AppUserTelephone.setDataSource(GetUserEntityName);
	AppUserTelephone.setDataField('UserTelephone');
}
function _AppUserTelephone_ctor()
{
	CreateLabel('AppUserTelephone', _initAppUserTelephone, null);
}
</script>
<% AppUserTelephone.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityTelephone 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityTelephone()
{
	AdminEntityTelephone.setDataSource(GetAdminEntityName);
	AdminEntityTelephone.setDataField('EntityTelephone');
}
function _AdminEntityTelephone_ctor()
{
	CreateLabel('AdminEntityTelephone', _initAdminEntityTelephone, null);
}
</script>
<% AdminEntityTelephone.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
            
            
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppUserExtension 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppUserExtension">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppUserExtension()
{
	AppUserExtension.setDataSource(GetUserEntityName);
	AppUserExtension.setDataField('UserExtension');
}
function _AppUserExtension_ctor()
{
	CreateLabel('AppUserExtension', _initAppUserExtension, null);
}
</script>
<% AppUserExtension.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
        </td>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityExtension 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityExtension()
{
	AdminEntityExtension.setDataSource(GetAdminEntityName);
	AdminEntityExtension.setDataField('EntityExtension');
}
function _AdminEntityExtension_ctor()
{
	CreateLabel('AdminEntityExtension', _initAdminEntityExtension, null);
}
</script>
<% AdminEntityExtension.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
           
</td></tr></tbody>

</table>


<br><br>

<table align="left" border="0" cellPadding="0" cellSpacing="0">
    
    <tr>
        <td align="left" colSpan="8"><strong><font face="arial" color="maroon" size="3">
	1.2 
            CO Code Information:</font></strong>
    <tr>
        <td align="right" colSpan="8">
            <div align="left">&nbsp; </div>
    <tr>
        <td align="left" colSpan="2" width="100"><strong><font face="arial" size="2">&nbsp;NPA:&nbsp;</font></strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NPA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 31px" 
	width=31>
	<PARAM NAME="_ExtentX" VALUE="820">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPA()
{
	NPA.setDataSource(GetPart1Data);
	NPA.setDataField('NPA');
}
function _NPA_ctor()
{
	CreateLabel('NPA', _initNPA, null);
}
</script>
<% NPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
        <td align="left" colSpan="2" width="100"><strong><font face="arial" size="2">&nbsp;LATA:&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=LATA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
	width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="LATA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="LATA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLATA()
{
	LATA.setDataSource(GetPart1Data);
	LATA.setDataField('LATA');
}
function _LATA_ctor()
{
	CreateLabel('LATA', _initLATA, null);
}
</script>
<% LATA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
        <td align="left" colSpan="4" width="100"><strong><font face="arial" size="2">&nbsp;OCN:&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=OCN style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 32px" 
	width=32>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="OCN">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="OCN">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOCN()
{
	OCN.setDataSource(GetPart1Data);
	OCN.setDataField('OCN');
}
function _OCN_ctor()
{
	CreateLabel('OCN', _initOCN, null);
}
</script>
<% OCN.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            

    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">Switch 
            Identification (Switching Entity / POI):&nbsp;&nbsp;</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=SwitchID 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="SwitchID">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initSwitchID()
{
	SwitchID.setDataSource(GetPart1Data);
	SwitchID.setDataField('SwitchID');
}
function _SwitchID_ctor()
{
	CreateLabel('SwitchID', _initSwitchID, null);
}
</script>
<% SwitchID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
    <tr>
        <td align="left" colSpan="5">
        <td align="left" colSpan="2"><font face="Arial" size="2">This is an eleven-character descriptor 
            of the switch provided by the owning entity for the purpose of 
            routing calls. This is the 11 character COMMON LANGUAGE Location 
            Identification - (CLLI) of the switch or POI.</font>
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">
	City or Wire 
            Center:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=WireCenter 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="WireCenter">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="WireCenter">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initWireCenter()
{
	WireCenter.setDataSource(GetPart1Data);
	WireCenter.setDataField('WireCenter');
}
function _WireCenter_ctor()
{
	CreateLabel('WireCenter', _initWireCenter, null);
}
</script>
<% WireCenter.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">Rate 
            Center:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RateCenter 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 75px" width=75>
	<PARAM NAME="_ExtentX" VALUE="1984">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RateCenter">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RateCenter">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRateCenter()
{
	RateCenter.setDataSource(GetPart1Data);
	RateCenter.setDataField('RateCenter');
}
function _RateCenter_ctor()
{
	CreateLabel('RateCenter', _initRateCenter, null);
}
</script>
<% RateCenter.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<font face="Arial" size="2">Rate Center Name must be a 
            tariffed Rate Center associated with toll billing.</font>
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">Route Same 
            as<strong><font face="arial" size="2">&nbsp;NPA:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RouteNPA 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 68px" width=68>
	<PARAM NAME="_ExtentX" VALUE="1799">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RouteNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RouteNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNPA()
{
	RouteNPA.setDataSource(GetPart1Data);
	RouteNPA.setDataField('RouteNPA');
}
function _RouteNPA_ctor()
{
	CreateLabel('RouteNPA', _initRouteNPA, null);
}
</script>
<% RouteNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<strong><font face="Arial" size="2">&nbsp; NXX:&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RouteNXX 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 68px" width=68>
	<PARAM NAME="_ExtentX" VALUE="1799">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RouteNXX">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RouteNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNXX()
{
	RouteNXX.setDataSource(GetPart1Data);
	RouteNXX.setDataField('RouteNXX');
}
function _RouteNXX_ctor()
{
	CreateLabel('RouteNXX', _initRouteNXX, null);
}
</script>
<% RouteNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;<strong><font face="Arial" size="2">Use Same Rate Center as<strong><font face="Arial" size="2">&nbsp;NPA:&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=CenterNPA 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 73px" width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CenterNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CenterNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNPA()
{
	CenterNPA.setDataSource(GetPart1Data);
	CenterNPA.setDataField('CenterNPA');
}
function _CenterNPA_ctor()
{
	CreateLabel('CenterNPA', _initCenterNPA, null);
}
</script>
<% CenterNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<strong><font face="Arial" size="2">&nbsp; NXX:&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=CenterNXX 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 73px" width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CenterNXX">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CenterNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNXX()
{
	CenterNXX.setDataSource(GetPart1Data);
	CenterNXX.setDataField('CenterNXX');
}
function _CenterNXX_ctor()
{
	CreateLabel('CenterNXX', _initCenterNXX, null);
}
</script>
<% CenterNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font></strong></font></strong></font></strong></font></strong></font> </strong>
    <tr>
        <td align="left" colSpan="7">&nbsp;&nbsp;
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="3" color="maroon" style="FONT-WEIGHT: bold">
1.3 Dates:</font></strong>
    <tr>
        <td align="left" colSpan="7">&nbsp;
    <tr>
        <td align="left" colSpan="7"><strong><font face="Arial" size="2">Application 
Date:&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ApplicationDate 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ApplicationDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initApplicationDate()
{
	ApplicationDate.setDataSource(GetPart1Data);
	ApplicationDate.setDataField('ApplicationDate');
}
function _ApplicationDate_ctor()
{
	CreateLabel('ApplicationDate', _initApplicationDate, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("ApplicationDate"),vbShortDate)
'ApplicationDate.display
%>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<font face="arial" size="1">dd/mm/ccyy</font></font></strong> 
    <tr>
        <td align="left" colSpan="7"><strong><font face="Arial" size="2"><strong>Requested Effective Date:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequestedEffDate 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 116px" width=116>
	<PARAM NAME="_ExtentX" VALUE="3069">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestedEffDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestedEffDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedEffDate()
{
	RequestedEffDate.setDataSource(GetPart1Data);
	RequestedEffDate.setDataField('RequestedEffDate');
}
function _RequestedEffDate_ctor()
{
	CreateLabel('RequestedEffDate', _initRequestedEffDate, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("RequestedEffDate"),vbShortDate)
'RequestedEffDate.display
%>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<font face="arial" size="1">dd/mm/ccyy</font> 
            
</strong></font></strong>
    <tr>
        <td align="left" colSpan="7">&nbsp;
    <tr>
        <td align="left" colSpan="7">

<p><font face="Arial" size="2">The nationwide cut-over is a minimum of 45 days after the NXX 
            code request is input to RDBS and BRIDS. To the extent possible, 
            code applicants should avoid requesting an effective date that is an 
            interval less than 
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Part1Days 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 70px" width=70>
	<PARAM NAME="_ExtentX" VALUE="1852">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part1Days">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="P1getDays">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Blue">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Blue"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart1Days()
{
	Part1Days.setCaption('P1getDays');
}
function _Part1Days_ctor()
{
	CreateLabel('Part1Days', _initPart1Days, null);
}
</script>
<% Part1Days.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 calendar days from the submission of this 
            form. It should be noted that interconnection arrangements and 
            facilities need to be in place prior to activation of a code. Such 
            arrangements are outside the scope of these guidelines.</font></p>
    <tr>
        <td align="left" colSpan="7">&nbsp;
    <tr>
        <td align="left" colSpan="7">
<p><font face="Arial" size="2">Requests for code assignment should not be made more than 6 
            months prior to the requested effective date.</font></p>
    <tr>
        <td align="left" colSpan="7">&nbsp;
    <tr>
        <td align="left" colSpan="7">
<p><font face="Arial" size="2">Acknowledgment and indication of disposition of this 
            application will be provided to applicant as noted in Section 1.2 
            within ten working days from the date of receipt of this 
            application.</font></p>

<tr>
	<td align="left" colSpan="7">
</td></tr>
</table>&nbsp;&nbsp;
<p>
<!-- mld here -->
<br><br><br><br>
<br><br><br><br>
<br><br><br><br>
<br><br><br><br>
<br><br><br><br></p>
<p>&nbsp;</p>
<p><br></p>

<table align="left" background border="0" cellPadding="0" cellSpacing="0">
    
    <tr>
        <td align="left" colSpan="3"><strong><font color="maroon" face="Arial" Size="3">
1.4 
            Type of Entity Requesting the Code:</font></strong></font> 
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
<tr>
<td align="left" colSpan="3"><strong><font face="Arial" size="2"> A)&nbsp;</font><font color="maroon" face="Arial" Size="2">
            <% Response.Write TypeEntitychar %>
             &nbsp;</font></strong>&nbsp;&nbsp; 
<strong><font face="Arial" size="2">&nbsp; Other 
            Explained:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=OtherCarrierType 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 114px" width=114>
	<PARAM NAME="_ExtentX" VALUE="3016">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="OtherCarrierType">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="OtherCarrierType">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOtherCarrierType()
{
	OtherCarrierType.setDataSource(GetPart1Data);
	OtherCarrierType.setDataField('OtherCarrierType');
}
function _OtherCarrierType_ctor()
{
	CreateLabel('OtherCarrierType', _initOtherCarrierType, null);
}
</script>
<% OtherCarrierType.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font></strong>     

</td><tr></tr>
    <tr>
        <td align="left" colSpan="3" vAlign="top">&nbsp;


<tr>
        <td align="left" colSpan="3" vAlign="top"><font face="arial" size="2"><strong>B)&nbsp; Type of Service for which code is being 
            requested:</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=TypeOfService 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="TypeOfService">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="TypeOfService">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTypeOfService()
{
	TypeOfService.setDataSource(GetPart1Data);
	TypeOfService.setDataField('TypeOfService');
}
function _TypeOfService_ctor()
{
	CreateLabel('TypeOfService', _initTypeOfService, null);
}
</script>
<% TypeOfService.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</td></tr>
    <tr>
        <td align="left" colSpan="3">&nbsp;


<tr>
<td align="left" colSpan="3"><strong><font face="Arial" size="2">C)&nbsp; Is 
            certification or authorization required to provide this type of 
            service in the relevant geographic area?</strong></font><font face="Arial" size="2" color="maroon"><strong>
            <% Response.Write CertReqchar %></strong></font>
		</td>
	</tr>
	<tr>
<td width="25"></td>
        <td colSpan="2"><strong><font face="Arial" size="2">(1)&nbsp; If no, 
            explain:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=CertificationNoExplained 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 162px" width=162>
	<PARAM NAME="_ExtentX" VALUE="4286">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CertificationNoExplained">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CertificationNoExplained">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCertificationNoExplained()
{
	CertificationNoExplained.setDataSource(GetPart1Data);
	CertificationNoExplained.setDataField('CertificationNoExplained');
}
function _CertificationNoExplained_ctor()
{
	CreateLabel('CertificationNoExplained', _initCertificationNoExplained, null);
}
</script>
<% CertificationNoExplained.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></strong>
		</td>
	</tr>
	<tr>
		<td align="left"><font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;</strong></font></td>
		<td align="left" colSpan="2"><font face="Arial" size="2"><strong>(2)&nbsp; 
            If yes, does your company have such certification or 
            authorization?</strong></font><font face="Arial" size="2" color="maroon"><strong>
            <% Response.write ReqCertReadychar %></strong></font>
		</td>
    <tr>
        <td align="left"></td>
        <td align="left" colSpan="2">

<tr>
<td align="left">&nbsp;</td>
        <td align="left" width="35"></td>
        <td align="left"><strong><font face="Arial" size="2">(i)&nbsp; If yes, 
            indicate type and date of certification or authorization(e.g. letter 
            of authorization, license, Certificate of Public Convenience &amp; 
            Necessity (CPCN), tarriff, etc.):
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequiredYesExplanation 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 160px" width=160>
	<PARAM NAME="_ExtentX" VALUE="4233">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequiredYesExplanation">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequiredYesExplanation">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequiredYesExplanation()
{
	RequiredYesExplanation.setDataSource(GetPart1Data);
	RequiredYesExplanation.setDataField('RequiredYesExplanation');
}
function _RequiredYesExplanation_ctor()
{
	CreateLabel('RequiredYesExplanation', _initRequiredYesExplanation, null);
}
</script>
<% RequiredYesExplanation.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </font></strong>
            

<tr>
<td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        <td align="left">
</td>
        <td align="left"><font face="Arial" size="2"><strong>(ii)&nbsp; If no, 
            explain:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequiredNoExplanationel1 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 155px" width=155>
	<PARAM NAME="_ExtentX" VALUE="4101">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequiredNoExplanationel1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequiredNoExplanation">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequiredNoExplanationel1()
{
	RequiredNoExplanationel1.setDataSource(GetPart1Data);
	RequiredNoExplanationel1.setDataField('RequiredNoExplanation');
}
function _RequiredNoExplanationel1_ctor()
{
	CreateLabel('RequiredNoExplanationel1', _initRequiredNoExplanationel1, null);
}
</script>
<% RequiredNoExplanationel1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</strong></font>
    <tr>
        <td align="left" colSpan="3">&nbsp; 
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;&nbsp; 
    <tr>
        <td align="left" colSpan="3"><strong><font face="Arial" size="3" color="maroon">1.5&nbsp; Type of Request: 
    
	</font></strong>
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3"><font face="Arial" color="maroon" size="4"><strong>&nbsp;
            <% Response.Write TyReqchar1 %>
             &nbsp;</font></strong>
		<font face="Arial" size="2"><strong> 1)&nbsp; Code Assignment - 
            Requested NXX:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXXAssign 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXAssign">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX1preferred">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXAssign()
{
	NXXAssign.setDataSource(ToRAssRec);
	NXXAssign.setDataField('NXX1preferred');
}
function _NXXAssign_ctor()
{
	CreateLabel('NXXAssign', _initNXXAssign, null);
}
</script>
<% NXXAssign.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</strong></font>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <p><font face="Arial" size="2"><strong>Secondary NXXs if requested 
            becomes unavailable (optional, you can identify 2 
            NXXs):</strong></font></font></p>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXX2A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
	width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX2A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX2A()
{
	NXX2A.setDataSource(ToRAssRec);
	NXX2A.setDataField('NXX2');
}
function _NXX2A_ctor()
{
	CreateLabel('NXX2A', _initNXX2A, null);
}
</script>
<% NXX2A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXX3A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
	width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX3A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX3A()
{
	NXX3A.setDataSource(ToRAssRec);
	NXX3A.setDataField('NXX3');
}
function _NXX3A_ctor()
{
	CreateLabel('NXX3A', _initNXX3A, null);
}
</script>
<% NXX3A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2"><strong>Undesirable NXXs 
            (optional, you can identify 5 NXXs):</strong></font> 
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX1A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX1A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1A()
{
	NoNXX1A.setDataSource(ToRAssRec);
	NoNXX1A.setDataField('NoNXX1');
}
function _NoNXX1A_ctor()
{
	CreateLabel('NoNXX1A', _initNoNXX1A, null);
}
</script>
<% NoNXX1A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX2A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX2A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2A()
{
	NoNXX2A.setDataSource(ToRAssRec);
	NoNXX2A.setDataField('NoNXX2');
}
function _NoNXX2A_ctor()
{
	CreateLabel('NoNXX2A', _initNoNXX2A, null);
}
</script>
<% NoNXX2A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX3A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX3A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3A()
{
	NoNXX3A.setDataSource(ToRAssRec);
	NoNXX3A.setDataField('NoNXX3');
}
function _NoNXX3A_ctor()
{
	CreateLabel('NoNXX3A', _initNoNXX3A, null);
}
</script>
<% NoNXX3A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX4A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX4A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4A()
{
	NoNXX4A.setDataSource(ToRAssRec);
	NoNXX4A.setDataField('NoNXX4');
}
function _NoNXX4A_ctor()
{
	CreateLabel('NoNXX4A', _initNoNXX4A, null);
}
</script>
<% NoNXX4A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX5A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX5A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5A()
{
	NoNXX5A.setDataSource(ToRAssRec);
	NoNXX5A.setDataField('NoNXX5');
}
function _NoNXX5A_ctor()
{
	CreateLabel('NoNXX5A', _initNoNXX5A, null);
}
</script>
<% NoNXX5A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan--> 
        <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" color="maroon" size="2"><strong>
            <% Response.Write Reas4Reqchar %></strong></font>            
		</td>    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3"><strong><font face="Arial" color="maroon" size="4">&nbsp;
            <% Response.Write TyReqchar2 %>
             &nbsp;</font></strong>&nbsp;<font face="Arial" size="2"> 
            <strong>2) Update Information (Complete 
            Section 2).&nbsp;&nbsp; NXX requiring update:</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=NXXUpdate 
	style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="NXXUpdate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXXUpdate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXUpdate()
{
	NXXUpdate.setDataSource(GetPart1Data);
	NXXUpdate.setDataField('NXXUpdate');
}
function _NXXUpdate_ctor()
{
	CreateLabel('NXXUpdate', _initNXXUpdate, null);
}
</script>
<% NXXUpdate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</td>
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3"><font face="Arial" color="maroon" size="4"><strong>&nbsp;
            <% Response.Write TyReqchar3 %>
             &nbsp;</strong></font>&nbsp;
        <font face="Arial" size="2"><strong>3) Code Reservation Only - 
            Requested NXX:</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXXReserve 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXReserve">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX1preferred">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXReserve()
{
	NXXReserve.setDataSource(ToRRsvRec);
	NXXReserve.setDataField('NXX1preferred');
}
function _NXXReserve_ctor()
{
	CreateLabel('NXXReserve', _initNXXReserve, null);
}
</script>
<% NXXReserve.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <p><font face="Arial" size="2"><strong>Secondary NXXs if requested 
            becomes unavailable (optional, you can identify 2 
            NXXs):</strong></font></font></p>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXX2R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
	width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX2R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX2R()
{
	NXX2R.setDataSource(ToRRsvRec);
	NXX2R.setDataField('NXX2');
}
function _NXX2R_ctor()
{
	CreateLabel('NXX2R', _initNXX2R, null);
}
</script>
<% NXX2R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXX3R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
	width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX3R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX3R()
{
	NXX3R.setDataSource(ToRRsvRec);
	NXX3R.setDataField('NXX3');
}
function _NXX3R_ctor()
{
	CreateLabel('NXX3R', _initNXX3R, null);
}
</script>
<% NXX3R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2"><strong>Undesirable NXXs 
            (optional, you can identify 5 NXXs):</strong></font> 
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX1R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX1R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1R()
{
	NoNXX1R.setDataSource(ToRRsvRec);
	NoNXX1R.setDataField('NoNXX1');
}
function _NoNXX1R_ctor()
{
	CreateLabel('NoNXX1R', _initNoNXX1R, null);
}
</script>
<% NoNXX1R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX2R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX2R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2R()
{
	NoNXX2R.setDataSource(ToRRsvRec);
	NoNXX2R.setDataField('NoNXX2');
}
function _NoNXX2R_ctor()
{
	CreateLabel('NoNXX2R', _initNoNXX2R, null);
}
</script>
<% NoNXX2R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX3R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX3R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3R()
{
	NoNXX3R.setDataSource(ToRRsvRec);
	NoNXX3R.setDataField('NoNXX3');
}
function _NoNXX3R_ctor()
{
	CreateLabel('NoNXX3R', _initNoNXX3R, null);
}
</script>
<% NoNXX3R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX4R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX4R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4R()
{
	NoNXX4R.setDataSource(ToRRsvRec);
	NoNXX4R.setDataField('NoNXX4');
}
function _NoNXX4R_ctor()
{
	CreateLabel('NoNXX4R', _initNoNXX4R, null);
}
</script>
<% NoNXX4R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NoNXX5R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX5R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5R()
{
	NoNXX5R.setDataSource(ToRRsvRec);
	NoNXX5R.setDataField('NoNXX5');
}
function _NoNXX5R_ctor()
{
	CreateLabel('NoNXX5R', _initNoNXX5R, null);
}
</script>
<% NoNXX5R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="2">
		<td align="left"><font face="Arial" color="maroon" size="2"><strong>
            <% Response.Write ReasForReqchar %></strong></font>
		</td>
	</tr>
    <tr>
        <td align="left" colSpan="3">
            <p><font face="arial" size="2">
            When the Code Applicant desires to change the status of a CO 
            Code from reserved to assigned within the time frame contained 
            within the guidelines, the Code Applicant should complete and submit 
            a new Canadian Central Office Code (NXX) Assignment Request 
            Form.&nbsp;</font></p>
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
    <tr>
        <td align="left" colSpan="3"><font face="Arial" size="3" color="maroon" style="FONT-WEIGHT: bold">
	<strong>1.6 Additional Code Request For 
            Growth:</strong></font> 
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="2">
<p>&nbsp;</p>
        <td align="left">
<p><font face="Arial" size="2">Basis of eligibility for an additional code for growth 
            assigned to the switching entity/POI assumes the following: the 
            initial code or the code previously assigned to a new application 
            meets the exhaust criteria, as specified in the Central Office Code 
            (NXX) Assignment Guidelines, depending on whether the NPA is in a 
            non-jeopardy situation as described in Section 7.3 of the 
            guidelines. The appropriate situation shall be indicated below 
            (select one).</font></p>
    <tr>
        <td align="left" colSpan="3"><font face="Arial" size="2" color="maroon"><strong>&nbsp;
            <% Response.Write JeopardyName2 %>
             &nbsp;&nbsp;</font></strong>
        <font face="Arial" size="2"><strong>Non-Jeopardy NPA 
            Situation</strong></font> 
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2">I hereby certify that the existing CO Code(s) 
            (NXX) at this Switching Entity/POI is/(are) projected to exhaust 
            within 12 months of the date of this application. This fact is 
            documented on Appendix B and will be supplied to an auditor when 
            requested to do so per Appendix A of the Guidelines.</font>
    <tr>
        <td align="left" colSpan="3"><font face="Arial" size="2" color="maroon"><strong>&nbsp;
            <% Response.Write JeopardyName1 %>
             &nbsp;&nbsp;</font></strong>
        <font face="Arial" size="2"><strong>Jeopardy NPA Situation (see 
            Section 7.4(c) of the Guidelines) 
        </strong></font>
	<tr>
        <td align="left" colSpan="2"><font face="Arial"></font>
        <td align="left"><p><font face="Arial" size="2">I hereby certify that the existing CO Code(s) (NXX) at this 
            Switching Entity/POI is/(are) projected to exhaust within 6 months 
            of the date of this application. This fact is documented on Appendix 
            B and will be supplied to an auditor when requested to do so per 
            Appendix A of the Guidelines.</font></p><font face size="2"></font>
    <tr>
        <td align="left" colSpan="3">&nbsp;
<p>&nbsp;<p>
            <table background border="0" height="280" style="HEIGHT: 280px; WIDTH: 969px" width="969">
                
                <tr>
                    <td align="left" colSpan="12"><strong><font color="#993300" face="Arial" size="3">APPENDIX B:</font></strong> 
                <tr>
                    <td align="left" colSpan="12">
                <tr>
                    <td align="left" colSpan="12"><font face="Arial" size="2"><strong>NXXs included in growth 
                        calculation:</strong></font>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NXXGrowthCal 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 98px" width=98>
	<PARAM NAME="_ExtentX" VALUE="2593">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXGrowthCal">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXXGrowthCal">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXGrowthCal()
{
	NXXGrowthCal.setDataSource(GetPart1Data);
	NXXGrowthCal.setDataField('NXXGrowthCal');
}
function _NXXGrowthCal_ctor()
{
	CreateLabel('NXXGrowthCal', _initNXXGrowthCal, null);
}
</script>
<% NXXGrowthCal.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                <tr>
                    <td align="left" colSpan="12"><strong><font face="Arial" size="2">A.&nbsp; Telephone Numbers (TNs) 
                        Available for Assignment (See Glossary):</font></strong>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=TNs style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 27px" 
	width=27>
	<PARAM NAME="_ExtentX" VALUE="714">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="TNs">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="TNs">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTNs()
{
	TNs.setDataSource(GetPart1Data);
	TNs.setDataField('TNs');
}
function _TNs_ctor()
{
	CreateLabel('TNs', _initTNs, null);
}
</script>
<% TNs.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                <tr>
                    <td align="left" colSpan="12"><font face="Arial" size="2">Definitions of 
                        terms may be found in the Glossary section of the 
                        Central Office Code (NXX) Assignment Guidelines.</font> 
                <tr>
                    <td align="left" colSpan="6">
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Month1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
	width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth1()
{
	Month1.setCaption('Month #1');
}
function _Month1_ctor()
{
	CreateLabel('Month1', _initMonth1, null);
}
</script>
<% Month1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Month2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
	width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month2">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth2()
{
	Month2.setCaption('Month #2');
}
function _Month2_ctor()
{
	CreateLabel('Month2', _initMonth2, null);
}
</script>
<% Month2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Month3 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
	width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth3()
{
	Month3.setCaption('Month #3');
}
function _Month3_ctor()
{
	CreateLabel('Month3', _initMonth3, null);
}
</script>
<% Month3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Month4 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
	width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month4">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth4()
{
	Month4.setCaption('Month #4');
}
function _Month4_ctor()
{
	CreateLabel('Month4', _initMonth4, null);
}
</script>
<% Month4.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=MOnth5 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
	width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="MOnth5">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMOnth5()
{
	MOnth5.setCaption('Month #5');
}
function _MOnth5_ctor()
{
	CreateLabel('MOnth5', _initMOnth5, null);
}
</script>
<% MOnth5.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Month6 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
	width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month6">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #6">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth6()
{
	Month6.setCaption('Month #6');
}
function _Month6_ctor()
{
	CreateLabel('Month6', _initMonth6, null);
}
</script>
<% Month6.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
                <tr>
                    <td align="left" colSpan="6"><strong><font face="Arial" size="2">B.&nbsp; Previous 6-month growth 
                        history:</font></strong></td>
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Prev6Month1 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Prev6Month1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="Prev6Month1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPrev6Month1()
{
	Prev6Month1.setDataSource(GetPart1Data);
	Prev6Month1.setDataField('Prev6Month1');
}
function _Prev6Month1_ctor()
{
	CreateLabel('Prev6Month1', _initPrev6Month1, null);
}
</script>
<% Prev6Month1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Prev6Month2 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Prev6Month2">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="Prev6Month2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPrev6Month2()
{
	Prev6Month2.setDataSource(GetPart1Data);
	Prev6Month2.setDataField('Prev6Month2');
}
function _Prev6Month2_ctor()
{
	CreateLabel('Prev6Month2', _initPrev6Month2, null);
}
</script>
<% Prev6Month2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Prev6Month3 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Prev6Month3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="Prev6Month3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPrev6Month3()
{
	Prev6Month3.setDataSource(GetPart1Data);
	Prev6Month3.setDataField('Prev6Month3');
}
function _Prev6Month3_ctor()
{
	CreateLabel('Prev6Month3', _initPrev6Month3, null);
}
</script>
<% Prev6Month3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Prev6Month4 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Prev6Month4">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="Prev6Month4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPrev6Month4()
{
	Prev6Month4.setDataSource(GetPart1Data);
	Prev6Month4.setDataField('Prev6Month4');
}
function _Prev6Month4_ctor()
{
	CreateLabel('Prev6Month4', _initPrev6Month4, null);
}
</script>
<% Prev6Month4.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Prev6Month5 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Prev6Month5">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="Prev6Month5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPrev6Month5()
{
	Prev6Month5.setDataSource(GetPart1Data);
	Prev6Month5.setDataField('Prev6Month5');
}
function _Prev6Month5_ctor()
{
	CreateLabel('Prev6Month5', _initPrev6Month5, null);
}
</script>
<% Prev6Month5.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Prev6Month6 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Prev6Month6">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="Prev6Month6">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPrev6Month6()
{
	Prev6Month6.setDataSource(GetPart1Data);
	Prev6Month6.setDataField('Prev6Month6');
}
function _Prev6Month6_ctor()
{
	CreateLabel('Prev6Month6', _initPrev6Month6, null);
}
</script>
<% Prev6Month6.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
                <tr>
                    <td align="left" colSpan="12"><font face="Arial" size="2">Telephone Numbers 
                        (TNs) assigned in each previous month, starting with the 
                        most distant month as Month #1, and Month #6 as the 
                        current month.</font></td></tr>
                <tr>
                    <td align="left" colSpan="6"><strong><font face="Arial" size="2">C.&nbsp; Projected growth - 
                        Months&nbsp;&nbsp; 1-6:</font></strong></td>
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth16Month1 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 136px" width=136>
	<PARAM NAME="_ExtentX" VALUE="3598">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth16Month1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth16Month1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth16Month1()
{
	ProjGrowth16Month1.setDataSource(GetPart1Data);
	ProjGrowth16Month1.setDataField('ProjGrowth16Month1');
}
function _ProjGrowth16Month1_ctor()
{
	CreateLabel('ProjGrowth16Month1', _initProjGrowth16Month1, null);
}
</script>
<% ProjGrowth16Month1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth16Month2 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 136px" width=136>
	<PARAM NAME="_ExtentX" VALUE="3598">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth16Month2">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth16Month2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth16Month2()
{
	ProjGrowth16Month2.setDataSource(GetPart1Data);
	ProjGrowth16Month2.setDataField('ProjGrowth16Month2');
}
function _ProjGrowth16Month2_ctor()
{
	CreateLabel('ProjGrowth16Month2', _initProjGrowth16Month2, null);
}
</script>
<% ProjGrowth16Month2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth16Month3 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 136px" width=136>
	<PARAM NAME="_ExtentX" VALUE="3598">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth16Month3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth16Month3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth16Month3()
{
	ProjGrowth16Month3.setDataSource(GetPart1Data);
	ProjGrowth16Month3.setDataField('ProjGrowth16Month3');
}
function _ProjGrowth16Month3_ctor()
{
	CreateLabel('ProjGrowth16Month3', _initProjGrowth16Month3, null);
}
</script>
<% ProjGrowth16Month3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth16Month4 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 136px" width=136>
	<PARAM NAME="_ExtentX" VALUE="3598">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth16Month4">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth16Month4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth16Month4()
{
	ProjGrowth16Month4.setDataSource(GetPart1Data);
	ProjGrowth16Month4.setDataField('ProjGrowth16Month4');
}
function _ProjGrowth16Month4_ctor()
{
	CreateLabel('ProjGrowth16Month4', _initProjGrowth16Month4, null);
}
</script>
<% ProjGrowth16Month4.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth16Month5 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 136px" width=136>
	<PARAM NAME="_ExtentX" VALUE="3598">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth16Month5">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth16Month5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth16Month5()
{
	ProjGrowth16Month5.setDataSource(GetPart1Data);
	ProjGrowth16Month5.setDataField('ProjGrowth16Month5');
}
function _ProjGrowth16Month5_ctor()
{
	CreateLabel('ProjGrowth16Month5', _initProjGrowth16Month5, null);
}
</script>
<% ProjGrowth16Month5.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth16Month6 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 136px" width=136>
	<PARAM NAME="_ExtentX" VALUE="3598">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth16Month6">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth16Month6">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth16Month6()
{
	ProjGrowth16Month6.setDataSource(GetPart1Data);
	ProjGrowth16Month6.setDataField('ProjGrowth16Month6');
}
function _ProjGrowth16Month6_ctor()
{
	CreateLabel('ProjGrowth16Month6', _initProjGrowth16Month6, null);
}
</script>
<% ProjGrowth16Month6.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
                <tr>
                    <td align="left" colSpan="6">&nbsp;&nbsp;&nbsp;&nbsp; 
                        <strong><font face="Arial" size="2">Projected growth - Months&nbsp; 
                        7-12:</font></strong></td>
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth712Month1 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 143px" width=143>
	<PARAM NAME="_ExtentX" VALUE="3784">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth712Month1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth712Month1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth712Month1()
{
	ProjGrowth712Month1.setDataSource(GetPart1Data);
	ProjGrowth712Month1.setDataField('ProjGrowth712Month1');
}
function _ProjGrowth712Month1_ctor()
{
	CreateLabel('ProjGrowth712Month1', _initProjGrowth712Month1, null);
}
</script>
<% ProjGrowth712Month1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth712Month2 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 143px" width=143>
	<PARAM NAME="_ExtentX" VALUE="3784">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth712Month2">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth712Month2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth712Month2()
{
	ProjGrowth712Month2.setDataSource(GetPart1Data);
	ProjGrowth712Month2.setDataField('ProjGrowth712Month2');
}
function _ProjGrowth712Month2_ctor()
{
	CreateLabel('ProjGrowth712Month2', _initProjGrowth712Month2, null);
}
</script>
<% ProjGrowth712Month2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth712Month3 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 143px" width=143>
	<PARAM NAME="_ExtentX" VALUE="3784">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth712Month3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth712Month3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth712Month3()
{
	ProjGrowth712Month3.setDataSource(GetPart1Data);
	ProjGrowth712Month3.setDataField('ProjGrowth712Month3');
}
function _ProjGrowth712Month3_ctor()
{
	CreateLabel('ProjGrowth712Month3', _initProjGrowth712Month3, null);
}
</script>
<% ProjGrowth712Month3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth712Month4 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 143px" width=143>
	<PARAM NAME="_ExtentX" VALUE="3784">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth712Month4">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth712Month4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth712Month4()
{
	ProjGrowth712Month4.setDataSource(GetPart1Data);
	ProjGrowth712Month4.setDataField('ProjGrowth712Month4');
}
function _ProjGrowth712Month4_ctor()
{
	CreateLabel('ProjGrowth712Month4', _initProjGrowth712Month4, null);
}
</script>
<% ProjGrowth712Month4.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth712Month5 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 143px" width=143>
	<PARAM NAME="_ExtentX" VALUE="3784">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth712Month5">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth712Month5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth712Month5()
{
	ProjGrowth712Month5.setDataSource(GetPart1Data);
	ProjGrowth712Month5.setDataField('ProjGrowth712Month5');
}
function _ProjGrowth712Month5_ctor()
{
	CreateLabel('ProjGrowth712Month5', _initProjGrowth712Month5, null);
}
</script>
<% ProjGrowth712Month5.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <td align="middle">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ProjGrowth712Month6 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 143px" width=143>
	<PARAM NAME="_ExtentX" VALUE="3784">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ProjGrowth712Month6">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ProjGrowth712Month6">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initProjGrowth712Month6()
{
	ProjGrowth712Month6.setDataSource(GetPart1Data);
	ProjGrowth712Month6.setDataField('ProjGrowth712Month6');
}
function _ProjGrowth712Month6_ctor()
{
	CreateLabel('ProjGrowth712Month6', _initProjGrowth712Month6, null);
}
</script>
<% ProjGrowth712Month6.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
                <tr>
                    <td align="left" colSpan="12"><font face="Arial" size="2">TNs assigned in 
                        each following month, starting with the most recent 
                        month as Month #1.&nbsp; In a jeopardy situation, only 6 
                        months growth porjection is required.</font></td></tr>
                <tr>
                    <td align="left" colSpan="12"><strong><font face="Arial" size="2">D.&nbsp; Average Monthly Growth 
                        Rate (From Part C above):
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=AvgMonGrowthRate 
	style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 158px" width=158>
	<PARAM NAME="_ExtentX" VALUE="4180">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="AvgMonGrowthRate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AvgMonGrowthRate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAvgMonGrowthRate()
{
	AvgMonGrowthRate.setDataSource(GetPart1Data);
	AvgMonGrowthRate.setDataField('AvgMonGrowthRate');
}
function _AvgMonGrowthRate_ctor()
{
	CreateLabel('AvgMonGrowthRate', _initAvgMonGrowthRate, null);
}
</script>
<% AvgMonGrowthRate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></strong>
</td></tr>
                <tr>
                    <td align="left" colSpan="12"><strong><font face="Arial" size="2">E.&nbsp; Months to Exhaust = TNs 
                        Available for Assignment (A) / Average Monthly Growth 
                        Rate (D) =</strong>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=MonthsToExhaust 
	style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 144px" width=144>
	<PARAM NAME="_ExtentX" VALUE="3810">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="MonthsToExhaust">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="MonthsToExhaust">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonthsToExhaust()
{
	MonthsToExhaust.setDataSource(GetPart1Data);
	MonthsToExhaust.setDataField('MonthsToExhaust');
}
function _MonthsToExhaust_ctor()
{
	CreateLabel('MonthsToExhaust', _initMonthsToExhaust, null);
}
</script>
<% MonthsToExhaust.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></td></tr>
                <tr>
                    <td align="left" colSpan="12"><font face="Arial" size="2">To be assigned an 
                        additional CO Code for growth, &quot;Months to 
                        Exhaust&quot; must be less than or equal to 12 month for 
                        a non -jeopardy NPA (See Section 4.2.1 of the 
                        Guidelines), or less than or equal to 6 months for a 
                        jeopardy NPA (See Section 8.4(c) of the 
                        Guidelines).</font></td></tr>
                <tr>
                    <td align="left" colSpan="12"><strong><font face="Arial" size="2">Explanation:
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppendixBExplanation 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" width=150>
	<PARAM NAME="_ExtentX" VALUE="3969">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppendixBExplanation">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AppendixBExplanation">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppendixBExplanation()
{
	AppendixBExplanation.setDataSource(GetPart1Data);
	AppendixBExplanation.setDataField('AppendixBExplanation');
}
function _AppendixBExplanation_ctor()
{
	CreateLabel('AppendixBExplanation', _initAppendixBExplanation, null);
}
</script>
<% AppendixBExplanation.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </font></strong>
</td></tr></table>
<p>&nbsp;<p></p>
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
    <tr>
        <td align="left" colSpan="3"><font face="Arial" size="3" color="maroon" style="FONT-WEIGHT: bold">
		<strong>1.7 Code Request for New 
            Application(see Section 4.2 of the Guidelines)</strong></font>
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2">
	Basis of eligibility for an additional code 
            means that there has not been a code assigned to this switching 
            entity/point of interconnection for this purpose. (Check the 
            applicable space and, if applicable, provide the requested 
            information). If eligibility is based on a category that requires 
            additional explanation or documentation and the code administrator 
            denies a request, the applicant has the option to pursue an appeals 
            process.</font>
    <tr>
        <td align="left" colSpan="3">
			 <dd><font face="Arial" color="maroon" size="4"><strong>&nbsp;
            <% Response.Write CodeReqNewchar1 %>
             &nbsp;&nbsp;</font></strong><strong><font face="Arial" size="2"> Code is necessary for 
            distinct routing, rating or billing purposes.<font face="Arial" Size="2"><strong> Any additional 
            information that can be provided by the Code Applicant may 
            facilitate the processing of that 
            application.</strong></font></strong></font> 
			</dd>
		</td>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <strong><font face="Arial" size="2">Description:</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequestNewNecessary 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 147px" width=147>
	<PARAM NAME="_ExtentX" VALUE="3889">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestNewNecessary">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestNewNecessary">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestNewNecessary()
{
	RequestNewNecessary.setDataSource(GetPart1Data);
	RequestNewNecessary.setDataField('RequestNewNecessary');
}
function _RequestNewNecessary_ctor()
{
	CreateLabel('RequestNewNecessary', _initRequestNewNecessary, null);
}
</script>
<% RequestNewNecessary.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</td>            
    <tr>
        <td align="left" colSpan="3">
			<dd><font face="Arial" color="maroon" size="4"><strong>&nbsp;
            <% Response.Write CodeReqNewchar2 %>
             &nbsp;&nbsp;</font></strong>
		<font face="Arial" size="2"><strong>Other <font size="2">The Code Applicant must provide an explanation of why existing 
            resources assigned to that entity cannot satisfy this 
            requirement.</strong></font></font> 
			</dd>
		</td>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <font face="Arial" size="2"><strong>Description:</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequestNewOther 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 118px" width=118>
	<PARAM NAME="_ExtentX" VALUE="3122">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestNewOther">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestNewOther">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestNewOther()
{
	RequestNewOther.setDataSource(GetPart1Data);
	RequestNewOther.setDataField('RequestNewOther');
}
function _RequestNewOther_ctor()
{
	CreateLabel('RequestNewOther', _initRequestNewOther, null);
}
</script>
<% RequestNewOther.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</td>            
    <tr>
        <td align="left" colSpan="3">
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
    <tr>
        <td align="left" colSpan="3"><strong><font face="Arial" size="3" color="maroon" style="FONT-WEIGHT: bold">
	1.8 Authorization for entry of Part 2 
            Information into Bellcore databases (Check applicable 
            space):</font></strong>
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
    <tr>
        <td align="right" valign="top" colSpan="2"><font face="Arial" color="maroon" size="4"><strong>&nbsp;
            <%Response.Write AuthPart2char1%>
             &nbsp;&nbsp;</font></strong>
        <td align="left"><font face="Arial" size="2"><strong>Yes - </strong>I 
            have attached a completed Part 2 of this form. This is the Code 
            Administrator's authorization to input/revise the indicated RDBS 
            and/or BRIDS data. Further, I understand that the Code Administrator 
            may not be the authorized party to input the data. The authorization 
            and/or data input responsibilities are determined on an Operating 
            Company Number level. If the Code Administrator advises me that said 
            Code Administrator does not have Administrative Operating Company 
            Number (AOCN) responsibility for my data inputs, I will contact 
            Bellcore-TRA to determine the correct AOCN company. Upon that 
            determination, I will submit Part 2 directly to the AOCN company for 
            input to RDBS and BRIDS.</font></font></strong> 
		</td>
	</tr>
    <tr>
        <td align="right" valign="top" colSpan="2"><font face="Arial" color="maroon" size="4"><strong>
            <% Response.Write AuthPart2char2 %>
            &nbsp;</font></strong></td>
        <td align="left"><font face="arial" size="2"><strong>No - </strong>Part 
            2 of this form is not attached. RDBS and BRIDS input will be the 
            responsibility of the Code Applicant. The 66 calendar day 
            nation-wide minimum interval cut-over for RDBS and BRIDS will not 
            begin until input into RDBS and BRIDS has been 
            completed.</font></font></td></tr>
</table> 

<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p>
<p>&nbsp;<p><p>&nbsp;<p><p>&nbsp;<p>



<table border="0" cellPadding="1" cellSpacing="1" width="80%" background style="HEIGHT: 1755px; WIDTH: 80%" height="1755">
    
    <tr>
        <td>
<table align="left" cellPadding="0" cellSpacing="0" width="74.91%" border="0" height="97" style="HEIGHT: 97px; WIDTH: 645px">
    <tr>
        <td align="left"><strong><font face="Arial Black" color="maroon" size="4">Part 3: Canadian CNA's 
                        Response/Confirmation Form</font></strong>
	<tr>
		<td align="left"><strong><font face="Arial" color="maroon" size="4">CNA Ticket #:
                        <%Response.Write(Tix)%></font></strong></td>
 </tr></table>
    <tr>
        <td>

<table align="left" background border="0" cellPadding="0" cellSpacing="0" width="745" height="224" style="HEIGHT: 224px; WIDTH: 745px">
    
    <tr>
        <td align="left"><strong><font face="Arial" size="4">Applicant Requested Dates:</font></strong>
        <td>
        <td align="right">
        <td align="left">
    
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><strong>Date of Requested 
                        Application:</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=DateofApp 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 95px" width=95>
	<PARAM NAME="_ExtentX" VALUE="2514">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofApp">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofApp()
{
	DateofApp.setDataSource(GetPart1Data);
	DateofApp.setDataField('ApplicationDate');
}
function _DateofApp_ctor()
{
	CreateLabel('DateofApp', _initDateofApp, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("ApplicationDate"),vbShortDate)
'DateofApp.display
%>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><strong>Date of 
                        Receipt:</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=DateofReceipt 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 85px" width=85>
	<PARAM NAME="_ExtentX" VALUE="2249">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofReceipt">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateofReceipt">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofReceipt()
{
	DateofReceipt.setDataSource(GetPart1Data);
	DateofReceipt.setDataField('DateofReceipt');
}
function _DateofReceipt_ctor()
{
	CreateLabel('DateofReceipt', _initDateofReceipt, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("DateofReceipt"),vbShortDate)
'DateofReceipt.display
%>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><strong>Date Response Due 
                        from CNA Admin:</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=DateofResponse 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 112px" width=112>
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofResponse">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateResponseDue">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofResponse()
{
	DateofResponse.setDataSource(GetPart1Data);
	DateofResponse.setDataField('DateResponseDue');
}
function _DateofResponse_ctor()
{
	CreateLabel('DateofResponse', _initDateofResponse, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("DateResponseDue"),vbShortDate)
'DateofResponse.display
%>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
        <td align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><strong>Requested Effective 
                        Date of CO Code:</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequestedEffDate 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 108px" width=108>
	<PARAM NAME="_ExtentX" VALUE="2858">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestedEffDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestedEffDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedEffDate()
{
	RequestedEffDate.setDataSource(GetPart1Data);
	RequestedEffDate.setDataField('RequestedEffDate');
}
function _RequestedEffDate_ctor()
{
	CreateLabel('RequestedEffDate', _initRequestedEffDate, null);
}
</script>
<%
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("RequestedEffDate"),vbShortDate)
'RequestedEffDate.display
%>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td></tr>
    <tr>
        <td align="right"></td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="right">
            <div align="right"><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><strong>The Preferred NPA-NXX 
                        Split ID:</strong></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font> 
                        </div></td>
        <td align="left">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=NPASplitID 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 69px" width=69>
	<PARAM NAME="_ExtentX" VALUE="1826">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NPASplitID">
	<PARAM NAME="DataSource" VALUE="GetCOCodeData">
	<PARAM NAME="DataField" VALUE="NPASplitID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPASplitID()
{
	NPASplitID.setDataSource(GetCOCodeData);
	NPASplitID.setDataField('NPASplitID');
}
function _NPASplitID_ctor()
{
	CreateLabel('NPASplitID', _initNPASplitID, null);
}
</script>
<% NPASplitID.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td align="right"><font face="Arial" size="2" nowrap><strong>Administrator who is Approving Part 
                        3:</font></font></font></strong>
           <td align="left"><font face="Arial" size="2" color="maroon">
                        <%Response.Write(AdminUserName)%></font>
           <td align="left">
            <div align="right">&nbsp;</div>
        <td align="left">
    <tr>
        <td align="right">
        <td align="right">
        <td align="left">
        <td align="left">
    <tr>
        <td align="right">
            <div align="left">&nbsp;</div>
        <td align="right">
        <td align="left">
        <td align="left">
    <tr>
        <td align="right">
            <div align="left">&nbsp;</div>
        <td align="right">
        <td align="left">
        <td align="left">
    <tr>
        <td align="right">
            <p><font face size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><strong>Extension Date:</strong></font></font></font></p>
        <td align="right">
<%  ExtentionDate.value = P3ExtentionDateVal
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=ExtentionDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="ExtentionDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="P3ExtentionDate">
	<PARAM NAME="DataField" VALUE="ExtentionDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initExtentionDate()
{
	ExtentionDate.setStyle(TXT_TEXTBOX);
	ExtentionDate.setDataSource(P3ExtentionDate);
	ExtentionDate.setDataField('ExtentionDate');
	ExtentionDate.setMaxLength(10);
	ExtentionDate.setColumnCount(10);
}
function _ExtentionDate_ctor()
{
	CreateTextbox('ExtentionDate', _initExtentionDate, null);
}
</script>
<% ExtentionDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
  
        <td align="left">

        <font face="arial" size="1"><strong>dd/mm/ccyy</strong></font>   
        <td align="left">
    <tr>
        <td align="right"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="right"></td>
        <td align="left">
            <div align="right">&nbsp; </div>
        <td align="left">
</td></tr></table>
    <tr>
        <td>
<table align="left" border="0" cellPadding="1" cellSpacing="1" width="75%" background="../brian_int/xca_Part3infield.asp#d7c7a4" style="WIDTH: 75%">
    
    <tr>
        <td align="left" colSpan="12"><strong><font face="Arial" size="4">SELECT THE APPROPRIATE PART 3 
                        ACTION:</strong> </font>
    <tr>
        <td align="left" colSpan="12">&nbsp;
 
    <tr>
        <td align="left" colSpan="12">&nbsp;</td>
    <tr>
        <td colSpan="12" nowrap><font face="Arial">
         
<input type="radio" NAME="Part3Result" value="a"> 
           <strong>Approve Part 1 
                        Request</strong></font>
    <tr>
        <td colSpan="12">&nbsp; 
    <tr>
        <td colSpan="12" nowrap><font face="Arial">
        <strong>-Code Reserved-</strong> </font>
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <p align="left"><strong><font size="2" face="Arial">Requested NPA:</font></strong></p>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=ReservedNPA 
	style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 38px" width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="ReservedNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedNPA()
{
	ReservedNPA.setDataSource(GetPart1Data);
	ReservedNPA.setDataField('NPA');
}
function _ReservedNPA_ctor()
{
	CreateLabel('ReservedNPA', _initReservedNPA, null);
}
</script>
<% ReservedNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <div align="left"><font face="Arial" size="2"><strong>Reserved NXX: 
            </strong></font></div>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=ReservedNXX 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="ReservedNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX1preferred">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedNXX()
{
	ReservedNXX.setStyle(TXT_TEXTBOX);
	ReservedNXX.setDataSource(GetPart1Data);
	ReservedNXX.setDataField('NXX1preferred');
	ReservedNXX.setMaxLength(3);
	ReservedNXX.setColumnCount(3);
}
function _ReservedNXX_ctor()
{
	CreateTextbox('ReservedNXX', _initReservedNXX, null);
}
</script>
<% ReservedNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td>
        <td align="left" colSpan="2" noWrap><font face="Arial" size="2"><strong>Secondary NXXs chosen 
                        that are still available:</strong></font>
        <td align="left" colSpan="9"><strong><font face="arial" color="green" size="2">
                        <%=RisNXX2%>&nbsp;

                        <%=RisNXX3%>&nbsp;

                        <!-- <%=RisNXX4%>-->

                        <!-- <%=RisNXX5%>--></font></strong>
    <tr>
        <td>
        <td align="left" colSpan="2" noWrap><font face="Arial" size="2"><strong>Undesirable 
                        NXXs:</strong></font>
        <td align="left" colSpan="9"><strong><font face="arial" color="red" size="2">
                        <%=RnisNXX1%>&nbsp;

                        <%=RnisNXX2%>&nbsp;

                        <%=RnisNXX3%>&nbsp;

                        <%=RnisNXX4%>&nbsp;
                        
                        <%=RnisNXX5%></font></strong>
    <tr>
        <td>
        <td align="left" colSpan="2" nowrap>
             <font face="Arial" size="2"><strong>Date of 
                        Reservation:</font></strong> 
        <td align="left" colSpan="9"><strong><font face="Arial" size="2">
                        <%=ReserveDate%></strong></font>
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <p align="left"><font face="Arial" size="2"><strong>Your code reservation will be honored 
                        until:</strong></font></p>
        <td align="left" colSpan="9"><strong><font face="Arial" size="2">
                        <%=ReserveDate450%></strong></font>
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <p align="left"><font face="Arial" size="2"><strong>Switch Identification (Switching 
                        Entity/POI):</font> </strong> </p>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=ReservedSwitchID 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ReservedSwitchID">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedSwitchID()
{
	ReservedSwitchID.setDataSource(GetPart1Data);
	ReservedSwitchID.setDataField('SwitchID');
}
function _ReservedSwitchID_ctor()
{
	CreateLabel('ReservedSwitchID', _initReservedSwitchID, null);
}
</script>
<% ReservedSwitchID.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td colSpan="12">&nbsp; 
    <tr>
        <td colSpan="12" nowrap>           
 <strong> <font face="Arial">-Code Update- 
            </font></strong>
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <p align="left"><strong><font size="2" face="Arial">Requested NPA:</font></strong></p>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=UpdatedNPA 
	style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 38px" width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="UpdatedNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUpdatedNPA()
{
	UpdatedNPA.setDataSource(GetPart1Data);
	UpdatedNPA.setDataField('NPA');
}
function _UpdatedNPA_ctor()
{
	CreateLabel('UpdatedNPA', _initUpdatedNPA, null);
}
</script>
<% UpdatedNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <div align="left"><font face="Arial"><strong><font face size="2">Updated NXX:</font> 
                        </strong></font></div>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=NXXUpdate 
	style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="NXXUpdate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXXUpdate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Green">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Green"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXUpdate()
{
	NXXUpdate.setDataSource(GetPart1Data);
	NXXUpdate.setDataField('NXXUpdate');
}
function _NXXUpdate_ctor()
{
	CreateLabel('NXXUpdate', _initNXXUpdate, null);
}
</script>
<% NXXUpdate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td colSpan="12">&nbsp; 
    <tr>
        <td colSpan="12">
        <font face="Arial"><strong>-Code 
                        Assigned-</strong></font>
    <tr>
        <td>
        <td align="right" colSpan="2" nowrap>
            <p align="left"><strong><font size="2" face="Arial"> Requested NPA:</font></strong> 
            </p>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=26 id=AssignedNPA 
	style="HEIGHT: 26px; LEFT: 0px; TOP: 0px; WIDTH: 43px" width=43>
	<PARAM NAME="_ExtentX" VALUE="1138">
	<PARAM NAME="_ExtentY" VALUE="688">
	<PARAM NAME="id" VALUE="AssignedNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="4">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="4" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAssignedNPA()
{
	AssignedNPA.setDataSource(GetPart1Data);
	AssignedNPA.setDataField('NPA');
}
function _AssignedNPA_ctor()
{
	CreateLabel('AssignedNPA', _initAssignedNPA, null);
}
</script>
<% AssignedNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
	</td>
    <tr>
        <td></td>
        <td align="left" colSpan="2" nowrap>
            <p><font face="Arial" size="2"><strong>Assigned NXX:</strong></font> 
            </p>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=AssignedNXX 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="AssignedNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX1Preferred">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAssignedNXX()
{
	AssignedNXX.setStyle(TXT_TEXTBOX);
	AssignedNXX.setDataSource(GetPart1Data);
	AssignedNXX.setDataField('NXX1Preferred');
	AssignedNXX.setMaxLength(3);
	AssignedNXX.setColumnCount(3);
}
function _AssignedNXX_ctor()
{
	CreateTextbox('AssignedNXX', _initAssignedNXX, null);
}
</script>
<% AssignedNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
  
        </td>
    <tr>
        <td>
        <td align="left" colSpan="2" noWrap><font face="Arial" size="2"><strong>Secondary NXXs chosen 
                        that are sill available:</strong></font>
        <td align="left" colSpan="9"><strong><font face="arial" color="green" size="2">
                        <%=AisNXX2%>&nbsp;

                        <%=AisNXX3%>&nbsp;

                       <!-- <%=AisNXX4%> -->

                       <!-- <%=AisNXX5%> --></font></strong>
    <tr>
        <td>
        <td align="left" colSpan="2" noWrap><font face="Arial" size="2"><strong>Undesirable 
                        NXXs:</strong></font>
        <td align="left" colSpan="9"><strong><font face="arial" color="red" size="2">
                        <%=AnisNXX1%>&nbsp;

                        <%=AnisNXX2%>&nbsp;

                        <%=AnisNXX3%>&nbsp;

                        <%=AnisNXX4%>&nbsp;
                     
                        <%=AnisNXX5%></font></strong>
    <tr>
        <td>
        <td align="left" colSpan="2" nowrap>
            <p><font face="Arial" size="2"><strong>NXX 
                        Effective Date:</strong></font></p>
        <td align="left" colSpan="9">
                        <% If TyReqvalue = "A" then
						'KT MADE CHANGES 2013-06-12; Format Date inside TextBox
                        EffDate.value=FormatDateTime(effDatee,vbShortDate)
                        else 
                        EffDate.value=""
                        end if%>

                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=EffDate style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
	width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EffDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEffDate()
{
	EffDate.setStyle(TXT_TEXTBOX);
	EffDate.setMaxLength(10);
	EffDate.setColumnCount(10);
}
function _EffDate_ctor()
{
	CreateTextbox('EffDate', _initEffDate, null);
}
</script>
<% EffDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 
            <font face="arial" size="1">
                        <%=EffDateLabel%></font>   
        
        </td>
    <tr>
        <td><strong></strong>
        <td align="left" colSpan="2" nowrap><font face="Arial" size="2"><strong>Switch Identification 
                        (Switching Entity/POI): </strong> 
            
</font></td>
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=SwitchID 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="SwitchID">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initSwitchID()
{
	SwitchID.setDataSource(GetPart1Data);
	SwitchID.setDataField('SwitchID');
}
function _SwitchID_ctor()
{
	CreateLabel('SwitchID', _initSwitchID, null);
}
</script>
<% SwitchID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
        </td>
    <tr>
        <td> <font face="Arial"> </font>&nbsp;            
</td>
        <td align="left" colSpan="2"> <font face="Arial" size="2">&nbsp;&nbsp;&nbsp; <strong>Rate Center: 
 
            </strong> </font> 
        <td align="left" colSpan="9">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RateCenter 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 75px" width=75>
	<PARAM NAME="_ExtentX" VALUE="1984">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RateCenter">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RateCenter">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRateCenter()
{
	RateCenter.setDataSource(GetPart1Data);
	RateCenter.setDataField('RateCenter');
}
function _RateCenter_ctor()
{
	CreateLabel('RateCenter', _initRateCenter, null);
}
</script>
<% RateCenter.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
         
</td>
            <tr>
        <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        <td align="left" colSpan="2"><font face="Arial" size="2"><strong>Is Routing and Rating information 
                        complete?</strong></font></td>
       <td align="left" colSpan="9">
        
        <input type="radio" name="RRComplete" value="Y "><strong><font face="Arial">Yes 
                </font> </strong>
         <input type="radio" name="RRComplete" value="N"><strong><font face="Arial">No 
            </font> </strong>
            </td>
    <tr>
        <td></td>
        <td align="left" colSpan="2"><font face="Arial"><font size="2"><strong>If no, 
                        Additional RDBS and BRIDS information is required as 
                        follows: 
                        </strong></font></font>
        <td align="left" colSpan="9">
            
<input name="RRDescription" Size="50" Maxlength="100">
      
    <tr>
        <td></td>
        <td colSpan="11"><font face="Arial" size="2"><strong>The Code Administrator
         <input type="radio" name="CNAResponsible" value="Y "><font face="Arial">IS 
                </font>
          <input type="radio" name="CNAResponsible" value="N"><font face="Arial">IS NOT 
 
            </font>responsible for inputting Part 2 Information into 
                        RDBS and BRIDS.</strong> 
            </font></td>
    <tr>
        <td></td>
        <td colSpan="11"><font face="Arial" size="2">

<strong>To be published in the LERG and TPM by:
                        <% If TyReqvalue = "A" then 
						' KT MADE CHANGES 2013-06-12; Format Date inside TextBox
                        LERGDate.value= FormatDateTime(LERGDatee,vbShortDate)
						RRReturnDate.value=FormatDateTime(LERGDatee,vbShortDate)
                        else
                        LERGDate.value=""
                        RRReturnDate.value=""
                        end if%>

<!--<input name="LERGDate" Size="10" Maxlength="10"> -->
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=LERGDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="LERGDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLERGDate()
{
	LERGDate.setStyle(TXT_TEXTBOX);
	LERGDate.setMaxLength(10);
	LERGDate.setColumnCount(10);
}
function _LERGDate_ctor()
{
	CreateTextbox('LERGDate', _initLERGDate, null);
}
</script>
<% LERGDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 <font face="arial" size="1">(dd/mm/ccyy),<font face="Arial" size="2"> additional RDBS and BRIDS 
                        information needs to be received by the Code 
                        Administrator no later than:
            
<!-- <input name="RRReturnDate" Size="10" Maxlength="10">  -->
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=RRReturnDate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RRReturnDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRRReturnDate()
{
	RRReturnDate.setStyle(TXT_TEXTBOX);
	RRReturnDate.setMaxLength(20);
	RRReturnDate.setColumnCount(20);
}
function _RRReturnDate_ctor()
{
	CreateTextbox('RRReturnDate', _initRRReturnDate, null);
}
</script>
<% RRReturnDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <font face="arial" size="1">(dd/mm/ccyy).<font face="Arial" size="2"> 
 
            
</font></font></font></font></strong></font></td></tr>
   
   <tr>
        
        <td align="left" colSpan="12">&nbsp;</td></tr>
       
    
   
    <tr>
        
        <td align="left" colSpan="12">&nbsp;</td></tr>
    <tr>
        <td colSpan="12"><strong>           
 
 <input type="radio" NAME="Part3Result" value="i"> <font face="Arial">Part 1 </font> <font face="Arial">Form</strong><strong> 
                        Incomplete. 
            </strong></font></td></tr>
    <tr>
        <td></td>
        <td align="left" colSpan="11">
            <p><font face="Arial" size="2">Additional information required in 
                        the following section(s):</font></p></td></td>
    <tr>
        <td> </td>
        <td colSpan="11">
            
<input name="Part3IncompleteDescription" Size="50" Maxlength="100"></td></tr>
    <tr>
        <td colSpan="12">&nbsp;
    <tr>
        <td colSpan="12"><font face="Arial">
            
<input type="radio" NAME="Part3Result" value="d"><strong>Part 1 Form completed, Code request denied. 
                        </strong> </font>
        
        </td></tr>
    <tr>
        <td>
        <td align="left" colSpan="11">  <font size="2" face="Arial">Explanation is: 
                        </font>
      
    <tr>
        <td></td>
        <td align="left" colSpan="11">
         <input name="Part3DenialDescription" Size="50" Maxlength="100">
       </td>
    <tr>
        
        <td align="right" colSpan="12">
            <div align="left">&nbsp;&nbsp;</div> </td>
        </tr>
    <tr>
       
        <td colSpan="12"><font face="Arial"><strong>
            
<input type="radio" NAME="Part3Result" value="s">Part 1 Assignment Activity Suspended by the 
                        Administrator.</strong> 
    </font></td></tr>
    <tr>
        <td>
        <td align="left" colSpan="11"> <font face="Arial" size="2">Explanation 
                        is:</font> 
        
    <tr>
        <td>
        <td align="left" colSpan="11">
             <input name="Part3SuspendedDescription" Size="50" Maxlength="100">
        </td>
    <tr>
        <td>
        <td align="left" colSpan="11">
            <font face="arial" size="3"><strong>Further Action:</strong></font>
        </td>
    <tr>
        <td>
        <td align="left" colSpan="11">
           <input name="Part3SuspendedFurtherAction" Size="50" Maxlength="100">
        </td>
     <tr>
    <td colSpan="12" align="left"> <font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td>
            </tr>
            <tr>
    <td colSpan="12" align="left"> <font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td>
            </tr>
    <tr>
      
        <td colSpan="3" align="left">
                        <% NPAinJeopardy.selectByCaption(JeopardyNameP3)%><font face="Arial" size="2"><strong>&nbsp; NPA Jeopardy 
                        =</strong></font> </td>
<td align="bottom">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E45D-DC5F-11D0-9846-0000F8027CA0" height=38 id=NPAinJeopardy 
	style="HEIGHT: 38px; LEFT: 0px; TOP: 0px; WIDTH: 132px" width=132>
	<PARAM NAME="_ExtentX" VALUE="3493">
	<PARAM NAME="_ExtentY" VALUE="820">
	<PARAM NAME="id" VALUE="NPAinJeopardy">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="1">
	<PARAM NAME="BType" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="2">
	<PARAM NAME="CLED1" VALUE="YES">
	<PARAM NAME="CLEV1" VALUE="y">
	<PARAM NAME="CLED2" VALUE="NO">
	<PARAM NAME="CLEV2" VALUE="n">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/OptionGrp.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPAinJeopardy()
{
	NPAinJeopardy.addItem('YES', 'y');
	NPAinJeopardy.addItem('NO', 'n');
	NPAinJeopardy.setAlignment(1);
}
function _NPAinJeopardy_ctor()
{
	CreateOptionGroup('NPAinJeopardy', _initNPAinJeopardy, null);
}
</script>
<% NPAinJeopardy.display %>

<!--METADATA TYPE="DesignerControl" endspan-->


  
 <!-- <font face="Arial" color=maroon size="3"> <%=JeopardyNameP3%> --></font></strong></font>           
    <tr>
        <td></td>
        <td align="left" colSpan="3">
            <font face="Arial" size="1">
            If YES, refer to 
                        Section 7 of the assignment guidelines</font>
             </td></tr>
    <tr>
    <td colSpan="12" align="left"> <font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td>
            </tr>
            <tr>
    <td colSpan="12" align="left"> <font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td>
            </tr>
    <tr>
        <td colSpan="12">&nbsp;&nbsp;
            <font face="Arial" size="3"><strong>Remarks: 
            </strong>
           
            <input name="Remarks" Size="50" Maxlength="100"></font></td>
</tr><tr>
<td align="left" colSpan="12">&nbsp;&nbsp; 
</td>
</tr><tr>
<td align="left" colSpan="12">
 <input type="submit" value="Submit" id="button2" name="submitt">
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnPending 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 135px" width=135>
	<PARAM NAME="_ExtentX" VALUE="3572">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnPending">
	<PARAM NAME="Caption" VALUE="Pending Request">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnPending()
{
	btnPending.value = 'Pending Request';
	btnPending.setStyle(0);
}
function _btnPending_ctor()
{
	CreateButton('btnPending', _initbtnPending, null);
}
</script>
<% btnPending.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturn 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturn">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturn()
{
	btnReturn.value = 'Return';
	btnReturn.setStyle(0);
}
function _btnReturn_ctor()
{
	CreateButton('btnReturn', _initbtnReturn, null);
}
</script>
<% btnReturn.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr></table></td></tr></table></p><br></form>
<%
session("Tix")=""
%>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
