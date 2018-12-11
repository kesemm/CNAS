<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">
</form>
<!--#include file="xca_CNASLib.inc"-->
<form action="xca_Part1int.asp" method="post" id="FormP1app" name="FormP1app" onSubmit="return validateForm()">
<html>
<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<title>Part 1 - Canadian Central Office Code (NXX) Assignment Request Form (Required)</title>
<script LANGUAGE="JavaScript"> <!--

       
    function checkdate(a) {						//a=document.frm.field.value

				var err=0,result
//new comment Make these vars meaningful for other to debug  -- Martin				
				if (a.length != 10) err=1
					d = a.substring(0, 2)//day  was-> b = a.substring(0, 2)// day
					c = a.substring(2, 3)// '/'
					b = a.substring(3, 5)//month was->d = a.substring(3, 5)// month
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
//				mon1 = b
//				day1= d
//				yr1 = f
//				appDate = new Date(mon1,day1,yr1)
//				TodaysDate = new Date()
//				TodaysDate.setHours("0")
//				TodaysDate.setMinutes("0")
//				TodaysDate.setSeconds("0")
//				if (Date.parse(TodaysDate)>Date.parse(appDate) {
//					err = 1
//				}
				if (err==1) {
				return false;
				}
				else {
					return true;
			   }
		}  
 function validDate(startDateStr,endDateStr,diffnum) //new 
		{
			
			var err=0
			//var DifNum=document.FormP1app.Part1Days.value
					//startDateStr to Date()
					 daySt= startDateStr.substring(0, 2)//  StartDate month 
					chr0 = startDateStr.substring(2, 3)// '/'
					monSt = startDateStr.substring(3, 5)//StartDate day
					chr1 = startDateStr.substring(5, 6)// '/'
					yrSt = startDateStr.substring(6, 10)//StartDate year
					//endDateStr to Date()
					dayEf= endDateStr.substring(0, 2)//EffectiveDate month
					chr0 = endDateStr.substring(2, 3)// '/'
					monEf = endDateStr.substring(3, 5)//EffectiveDate day
					chr1 = endDateStr.substring(5, 6)// '/'
					yrEf = endDateStr.substring(6, 10)//EffectiveDate year
			startDate = new Date(yrSt,monSt,daySt);
			startDate.setMonth(startDate.getMonth()-1);
			endDate = new Date(yrEf,monEf,dayEf);
			endDate.setMonth(endDate.getMonth()-1);
			validEffDate = new Date();
			validEffDate.setTime(startDate.getTime());
			validEffDate.setDate(validEffDate.getDate()+diffnum);			
			result = Date.parse(validEffDate)- Date.parse(endDate)
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
			       
 function validateForm()
        {
            var err=0, rryes=0, rcryes=0, toryes=0, abcyes=0, abcdyes=0, jepyes=0, crnyes=0, ap2yes=0, updyes=0, jep1yes=0, sCheck;
             formObj = document.FormP1app;
          
           
            
            if (formObj.AuthorizedRep.value == "") {
                alert("You have not filled in the Authorized Rep field. Please type in an Authorized Name and submit again");
                formObj.AuthorizedRep.focus();               
                return false;
            }
            if (formObj.AuthorizedRepTitle.value == "") {
                alert("You have not filled in the Authorized Rep Title field. Please type in an Authorized Name Title and submit again");
                formObj.AuthorizedRepTitle.focus();
                return false;
            }
            /*if (formObj.LATA.value == "") {
                alert("You have not filled in the LATA field. Please type in a 5 digit value and submit again");
                formObj.LATA.focus();               
                return false; 
            }*/
/*
Commented out by G. Brown on Sep 1/99 because the LATA is only a three digit number in Canada

             if (isNaN(formObj.LATA.value)){ 
                alert("The LATA field is not a number. Please type in a 4 digit value and submit again");
                formObj.LATA.focus();               
                return false;   
            }
            if ((formObj.LATA.value != "") && (formObj.LATA.value.length <4)) {
                alert("The LATA field must be 4 digits. Please type in a 4 digit value and submit again");
                formObj.LATA.focus();               
                return false; 
            }
            if (formObj.OCN.value == "") {
                alert("You have not filled in the OCN field. Please type in a 4 digit value and submit again");
                formObj.OCN.focus();
                return false;
            }

Changed by G. Brown on Sep 1/99 because the LATA is only a three digit number in Canada
*/
/* Commented out by G. Brown on Sept 26 2001 because Canadian LATA is 888
             if (isNaN(formObj.LATA.value)){ 
                alert("The LATA field is not a number. Please type in a 3 digit value and submit again");
                formObj.LATA.focus();               
                return false;   
            }
            if ((formObj.LATA.value != "") && (formObj.LATA.value.length <3)) {
                alert("The LATA field must be 3 digits. Please type in a 3 digit value and submit again");
                formObj.LATA.focus();               
                return false; 
            }
            if (formObj.OCN.value == "") {
                alert("You have not filled in the OCN field. Please type in a 3 digit value and submit again");
                formObj.OCN.focus();
                return false;
            }
*/
/*

End of Change

*/
/* Test

            if (isNaN(formObj.OCN.value)){ 
                alert("The OCN field is not a number. Please type in a 4 digit value and submit again");
                formObj.OCN.focus();               
                return false;   
            }

            if (formObj.OCN.value.length <4) {
                alert("The OCN field must be 4 digits. Please type in a 4 digit value and submit again");
                formObj.OCN.focus();               
                return false; 
            }
*/
            if (formObj.SwitchID.value == "") {
                alert("You have not filled in the Switch Identification field. Please type in an 11 character value and submit again");
                formObj.SwitchID.focus();
                return false;
            }
            if (formObj.SwitchID.value.length <11) {
                alert("The Switch Identification field must be 11 characters. Please type in an 11 character value and submit again");
                formObj.SwitchID.focus();               
                return false; 
            }
            if (formObj.WireCenter.value == "") {
            //|| formObj.WireCenter.value.length <10) {
                alert("The Wire Center field must be 1-40 characters. Please type in a 1-40 character value and submit again");
                formObj.WireCenter.focus();               
                return false; 
           }
     		if (formObj.RouteNPA.value != "") {
				if (isNaN(formObj.RouteNPA.value)){ 
                alert("The Route NPA field is not a number. Please type in a 3 digit value and submit again");
                formObj.RouteNPA.focus();               
                return false;   
				}
            	if ((formObj.RouteNPA.value <200 || formObj.RouteNPA.value>999)){
					alert("You must enter a value between 200-999 in the Route NPA field. Please retype and submit again");
					formObj.RouteNPA.focus();
                return false;
                }
           
		        if (isNaN(formObj.RouteNXX.value)){ 
                alert("The Route NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.RouteNXX.focus();               
                return false;   
				}
            	if ((formObj.RouteNXX.value <200 || formObj.RouteNXX.value>999)){
					alert("You must enter a value between 200-999 in the Route NXX field. Please retype and submit again");
					formObj.RouteNXX.focus();
                return false;
					}
                }
            if (formObj.RouteNPA.value == "" && formObj.RouteNXX.value != "") {
				if (isNaN(formObj.RouteNPA.value)){ 
                alert("The Route NPA field is required with Route NXX. Please type in a 3 digit value and submit again");
                formObj.RouteNPA.focus();               
                return false;   
				}
            }
            /*if (formObj.RateCenterAssignLookup.value == "") {
                alert("You must enter Rate Center field. Please enter appropriate values and submit again");
                formObj.RateCenterAssignLookup.focus();
                return false;
            }*/
            /*if (formObj.RateCenter.value == ""&& formObj.CenterNPA.value == "") {
                alert("You must enter either Rate Center or the Rate Center NPA/NXX field. Please enter appropriate values and submit again");
                formObj.RateCenter.focus();
                return false;
            }*/
            /*if (formObj.RateCenter.value != ""&& formObj.RateCenter.value.length <10) {
                alert("The Rate Center field must be 10 characters. Please type in an 10 character value and submit again");
                formObj.RateCenter.focus();               
                return false; 
           }*/
		    /*if (formObj.RateCenter.value != "" && formObj.CenterNPA.value != "") {
                alert("You may only enter either Rate Center or the Rate Center NPA/NXX field. Please enter appropriate values and submit again");
               formObj.RateCenter.focus();
                return false;
			}
			if (formObj.RateCenter.value != "" && formObj.CenterNXX.value != "") {
                alert("You may only enter either Rate Center or the Rate Center NPA/NXX field. Please enter appropriate values and submit again");
               formObj.RateCenter.focus();
                return false;
			}*/
			if (formObj.CenterNPA.value != "") {
				if (isNaN(formObj.CenterNPA.value)){ 
                alert("The Center NPA field is not a number. Please type in a 3 digit value and submit again");
                formObj.CenterNPA.focus();               
                return false;   
				}
            	if ((formObj.CenterNPA.value <200 || formObj.CenterNPA.value>999)){
					alert("You must enter a value between 200-999 in the Center NPA field. Please retype and submit again");
					formObj.CenterNPA.focus();
                return false;
                }
            	if (isNaN(formObj.CenterNXX.value)){ 
                alert("The Center NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.CenterNXX.focus();               
                return false;   
				}
            	if ((formObj.CenterNXX.value <200 || formObj.CenterNXX.value>999)){
					alert("You must enter a value between 200-999 in the Center NXX field. Please retype and submit again");
					formObj.CenterNXX.focus();
                return false;
						}
                }
          
            if (formObj.CenterNPA.value == "" && formObj.CenterNXX.value != "") {
				if (isNaN(formObj.CenterNPA.value)){ 
                alert("The Center NPA field is required with Center NXX. Please type in a 3 digit value and submit again");
                formObj.CenterNPA.focus();               
                return false;   
					}
				}
            if (formObj.ApplicationDate.value == "") {
                alert("You have not filled in the Application Date field. Please type in a valid date and submit again");
                formObj.ApplicationDate.focus();
                return false;
            }
            var result=checkdate(formObj.ApplicationDate.value) //this one             
            if (result==false)	{
				alert("The Application Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
                formObj.ApplicationDate.focus();
                return false;
            }
            
            if (formObj.RequestedEffDate.value == "") {
                alert("You have not filled in the Requested Effective Date field. Please type in a valid date and submit again");
                formObj.RequestedEffDate.focus();
                return false;
            }
            
            var result=checkdate(formObj.RequestedEffDate.value) //that one
            if (result==false)	{
				alert("The Requested Effective Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
                formObj.RequestedEffDate.focus();
                return false;
            }
       //     var valid=validDate(formObj.ApplicationDate.value,formObj.RequestedEffDate.value,formObj.Part1Days.value) //new
       //    if (valid == false){	//new
	//			alert("The Requested Effective field must be  greater than 45 days of the Application Date field. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
     //           formObj.RequestedEffDate.focus();
	//			return false; 
     //       }
            
            //check if RequestedEffDate >= 45 days past the ApplicationDate
			if (formObj.CarrierType.value== "o") {
				if (formObj.OtherCarrierType.value== "") {
                alert("You have selected 'other' carrier type.  Please enter your explanation and submit again");
                formObj.OtherCarrierType.focus();
                return false;
               }
             } 
            
            if (formObj.TypeOfService.value == "") {
                alert("You have not filled in the Type Of Service field. Please type in a valid data and submit again");
                formObj.TypeOfService.focus();
                return false;
            } 
            
			if (eval("formObj.CertificationRequired[0].checked") == true)
					 rryes++;
			if (eval("formObj.CertificationRequired[1].checked") == true)
					rryes++;
        
			if  (rryes==0){
    
                alert("You have not checked if certification or authorization is required in your geographical area. Please select one and submit again");
                 formObj.TypeOfService.focus();         
                return false;
				  }  
            if (formObj.CertificationRequired[1].checked) {
					if (formObj.CertificationNoExplained.value == ""){
                alert("You have stated that certification or authorization is not required in your geographical area.  Please enter your explanation and submit again");
                formObj.CertificationNoExplained.focus();
                return false;
                }
             }
             if (formObj.CertificationRequired[0].checked) {
				if (eval("formObj.RequiredCertificationReady[0].checked") == true)
					 rcryes++;
				if (eval("formObj.RequiredCertificationReady[1].checked") == true)
					rcryes++;
        
					if  (rcryes==0){
    
					alert("You have not stated if your company has certification or authorization. Please select one and submit again");
                       formObj.CertificationNoExplained.focus();      
					return false;
				  } 
				 if (formObj.RequiredCertificationReady[0].checked){
					if (formObj.RequiredYesExplanation.value == ""){
                alert("You have stated that you have certification or authorization. Please indicate type and certification/authorization date and submit again");
                formObj.RequiredYesExplanation.focus();
                return false;
                }
                } 
				if (formObj.RequiredCertificationReady[1].checked){
					if (formObj.RequiredNoExplanation.value == ""){
                alert("You have stated that you do not have certification or authorization. Please enter your explanation and submit again");
                formObj.RequiredNoExplanation.focus();
                return false;
                }
                }
            }
           
              
            
            if (eval("formObj.TypeOfRequest[0].checked") == true)
					 toryes++;
			if (eval("formObj.TypeOfRequest[1].checked") == true)
					toryes++;
			if (eval("formObj.TypeOfRequest[2].checked") == true)
					toryes++;
        
			if  (toryes==0){
    
                alert("You have not selected the Type of Request (Assignment, Update, or Reservation). Please select one and submit again");
                 formObj.NXX2A.focus();         
                return false;
				  }  
				  
		if (formObj.TypeOfRequest[0].checked){	  
		/*			
			if (formObj.NXX2A.value != "") {
				if (isNaN(formObj.NXX2A.value)){ 
                alert("The 1st Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX2A.focus();               
                return false;   
				}
                if ((formObj.NXX2A.value < 200) || (formObj.NXX2A.value > 999)){
					alert("You must enter a value between 200-999 in the 1st Secondary NXX field. Please retype and submit again");
					formObj.NXX2A.focus();
                return false;
                }
            }
            if (formObj.NXX3A.value != "") {
				if (isNaN(formObj.NXX3A.value)){ 
                alert("The 2nd Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX3A.focus();               
                return false; 
                 
				}
				if ((formObj.NXX3A.value <200 || formObj.NXX3A.value>999)){
                alert("You must enter a value between 200-999 in the 2nd Secondary NXX field. Please retype and submit again");
                formObj.NXX3A.focus();
                return false;
                }
            }
            /*if (formObj.NXX4A.value != "") {
				if (isNaN(formObj.NXX4A.value)){ 
                alert("The 3rd Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX4A.focus();               
                return false;   
				}
				if ((formObj.NXX4A.value <200 || formObj.NXX4A.value>999)){
                alert("You must enter a value between 200-999 in the 3rd Secondary NXX field. Please retype and submit again");
                formObj.NXX4A.focus();
                return false;
                }
            }
            if (formObj.NXX5A.value != "") {
				if (isNaN(formObj.NXX5A.value)){ 
                alert("The 4th Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX5A.focus();               
                return false;   
				}
				if ((formObj.NXX5A.value <200 || formObj.NXX5A.value>999)){
                alert("You must enter a value between 200-999 in the 4th Secondary NXX field. Please retype and submit again");
                formObj.NXX5A.focus();
                return false;
                }
            }*/
            if (formObj.NoNXX1A.value != "") {
				if (isNaN(formObj.NoNXX1A.value)){ 
                alert("The 1st Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX1A.focus();               
                return false;
				}
				if ((formObj.NoNXX1A.value <200 || formObj.NoNXX1A.value>999)){
                alert("You must enter a value between 200-999 in the 1st Undesired NXX field. Please retype and submit again");
                formObj.NoNXX1A.focus();
                return false;
                }
            }
            if (formObj.NoNXX2A.value != "") {  
				if (isNaN(formObj.NoNXX2A.value)){ 
                alert("The 2nd Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX2A.focus();               
                return false;   
				}
				if ((formObj.NoNXX2A.value <200 || formObj.NoNXX2A.value>999)){
                alert("You must enter a value between 200-999 in the 2nd Undesired NXX field. Please retype and submit again");
                formObj.NoNXX2A.focus();
                return false;
                }
            }
            if (formObj.NoNXX3A.value != "") {    
				if (isNaN(formObj.NoNXX3A.value)){ 
                alert("The 3rd Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX3A.focus();               
                return false;  
				}
				if ((formObj.NoNXX3A.value <200 || formObj.NoNXX3A.value>999)){
                alert("You must enter a value between 200-999 in the 3rd Undesired NXX field. Please retype and submit again");
                formObj.NoNXX3A.focus();
                return false;
                }
            }
            if (formObj.NoNXX4A.value != "") {    
				if (isNaN(formObj.NoNXX4A.value)){ 
                alert("The 4th Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX4A.focus();               
                return false;   
				}
				if ((formObj.NoNXX4A.value <200 || formObj.NoNXX4A.value>999)){
                alert("You must enter a value between 200-999 in the 4th Undesired NXX field. Please retype and submit again");
                formObj.NoNXX4A.focus();
                return false;
                }
            }
            if (formObj.NoNXX5A.value != "") {   
				if (isNaN(formObj.NoNXX5A.value)){ 
                alert("The 5th Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX5A.focus();               
                return false;   
				}
				if ((formObj.NoNXX5A.value <200 || formObj.NoNXX5A.value>999)){
                alert("You must enter a value between 200-999 in the 5th Undesired NXX field. Please retype and submit again");
                formObj.NoNXX5A.focus();
                return false;
                }
            }
            if (eval("formObj.ReasonForRequest[0].checked") == true)
						 abcyes++;
					if (eval("formObj.ReasonForRequest[1].checked") == true)
						abcyes++;
					if (eval("formObj.ReasonForRequest[2].checked") == true)
						abcyes++;
        
					if  (abcyes==0){
    
						alert("You have selected Code Assignment.  You must select either a), b), or c). Please select one and submit again");
                         formObj.NXX2A.focus();  
						return false;
						  }  
					  }
			if (formObj.TypeOfRequest[0].checked && formObj.ReasonForRequest[0].checked){	  
							 if (eval("formObj.AuthorizationPart2[0].checked") == true)
		 						    ap2yes++;
							  if (eval("formObj.AuthorizationPart2[1].checked") == true)
									ap2yes++;
							  if  (ap2yes==0){
    
									alert("You must fill out Section 1.8, submit again, and Complete Part 2");
                                    formObj.RequestNewOther.focus();
									return false;
							}
						}	  	
			if (formObj.TypeOfRequest[0].checked && formObj.ReasonForRequest[1].checked){	 	  
							if (eval("formObj.CodeRequestNew[0].checked") == true)
								 crnyes++;
							if (eval("formObj.CodeRequestNew[1].checked") == true)
								crnyes++;
					        
							if  (crnyes==0){
    
								alert("You must fill out Section 1.7 and submit again");
                                formObj.RequestNewNecessary.focus();
								return false;
								}	
					}  	
			if (formObj.TypeOfRequest[0].checked && formObj.ReasonForRequest[2].checked){	 	  
				  	if (eval("formObj.NPAinJeopardy[0].checked") == true)
						 jepyes++;
					if (eval("formObj.NPAinJeopardy[1].checked") == true)
						jepyes++;
					        
					if  (jepyes==0){
    
						alert("You must fill out Section 1.6 and submit again");
			            formObj.RequestNewNecessary.focus();
   						return false;
						}
				
					
				 }
	
	

		if (formObj.TypeOfRequest[1].checked){	 
				if (eval("formObj.ReasonForRequest[3].checked") == true)
						updyes++;
					        
					if  (updyes==0){
						alert("You have selected Code Update.  Please select the NXX update button and submit again");
			            formObj.NXXUpdate.focus();
   						return false;
						}
					
		}
						  
/*		if (formObj.TypeOfRequest[1].checked && formObj.ReasonForRequest[3].checked) {
			if (formObj.NXXUpdate == ""){
						alert("You have selected Code Update.  Please enter the NXX you are updating and submit again");
                formObj.NXXUpdate.focus();
                return false;
						}
				if (isNaN(formObj.NXXUpdate.value)){ 
                alert("The NXX Update field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXXUpdate.focus();               
                return false;   
				}
			
				if ((formObj.NXXUpdate.value <200 || formObj.NXXUpdate.value>999)){
                alert("You must enter a value between 200-999 in the NXX Update field. Please retype and submit again");
                formObj.NXXUpdate.focus();
                return false;
                }
		
		}*/
 			if (formObj.TypeOfRequest[2].checked){	
/*			
			 if (formObj.NXX2R.value != "") {
				if (isNaN(formObj.NXX2R.value)){ 
                alert("The 1st Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX2R.focus();               
                return false;   
				}
           	if ((formObj.NXX2R.value <200 || formObj.NXX2R.value>999)){
					alert("You must enter a value between 200-999 in the 1st Secondary NXX field. Please retype and submit again");
					formObj.NXX2R.focus();
                return false;
                }
             }
            if (formObj.NXX3R.value != "") {
				if (isNaN(formObj.NXX3R.value)){ 
                alert("The 2nd Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX3R.focus();               
                return false; 
                 
				}
				if ((formObj.NXX3R.value <200 || formObj.NXX3R.value>999)){
                alert("You must enter a value between 200-999 in the 2nd Secondary NXX field. Please retype and submit again");
                formObj.NXX3R.focus();
                return false;
                }
            }
            /*if (formObj.NXX4R.value != "") {
				if (isNaN(formObj.NXX4R.value)){ 
                alert("The 3rd Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX4R.focus();               
                return false;   
				}
				if ((formObj.NXX4R.value <200 || formObj.NXX4R.value>999)){
                alert("You must enter a value between 200-999 in the 3rd Secondary NXX field. Please retype and submit again");
                formObj.NXX4R.focus();
                return false;
                }
            }
            if (formObj.NXX5R.value != "") {
				if (isNaN(formObj.NXX5R.value)){ 
                alert("The 4th Secondary NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NXX5R.focus();               
                return false;   
				}
				if ((formObj.NXX5R.value <200 || formObj.NXX5R.value>999)){
                alert("You must enter a value between 200-999 in the 4th Secondary NXX field. Please retype and submit again");
                formObj.NXX5R.focus();
                return false;
                }
            }*/
            if (formObj.NoNXX1R.value != "") {
				if (isNaN(formObj.NoNXX1R.value)){ 
                alert("The 1st Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX1R.focus();               
                return false;
				}
				if ((formObj.NoNXX1R.value <200 || formObj.NoNXX1R.value>999)){
                alert("You must enter a value between 200-999 in the 1st Undesired NXX field. Please retype and submit again");
                formObj.NoNXX1R.focus();
                return false;
                }
            }
            if (formObj.NoNXX2R.value != "") {  
				if (isNaN(formObj.NoNXX2R.value)){ 
                alert("The 2nd Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX2R.focus();               
                return false;   
				}
				if ((formObj.NoNXX2R.value <200 || formObj.NoNXX2R.value>999)){
                alert("You must enter a value between 200-999 in the 2nd Undesired NXX field. Please retype and submit again");
                formObj.NoNXX2R.focus();
                return false;
                }
            }
            if (formObj.NoNXX3R.value != "") {    
				if (isNaN(formObj.NoNXX3R.value)){ 
                alert("The 3rd Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX3R.focus();               
                return false;  
				}
				if ((formObj.NoNXX3R.value <200 || formObj.NoNXX3R.value>999)){
                alert("You must enter a value between 200-999 in the 3rd Undesired NXX field. Please retype and submit again");
                formObj.NoNXX3R.focus();
                return false;
                }
            }
            if (formObj.NoNXX4R.value != "") {    
				if (isNaN(formObj.NoNXX4R.value)){ 
                alert("The 4th Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX4R.focus();               
                return false;   
				}
				if ((formObj.NoNXX4R.value <200 || formObj.NoNXX4R.value>999)){
                alert("You must enter a value between 200-999 in the 4th Undesired NXX field. Please retype and submit again");
                formObj.NoNXX4R.focus();
                return false;
                }
            }
            if (formObj.NoNXX5R.value != "") {   
				if (isNaN(formObj.NoNXX5R.value)){ 
                alert("The 5th Undesired NXX field is not a number. Please type in a 3 digit value and submit again");
                formObj.NoNXX5R.focus();               
                return false;   
				}
				if ((formObj.NoNXX5R.value <200 || formObj.NoNXX5R.value>999)){
                alert("You must enter a value between 200-999 in the 5th Undesired NXX field. Please retype and submit again");
                formObj.NoNXX5R.focus();
                return false;
                }
            }  
					if (eval("formObj.ReasonForRequest[4].checked") == true)
						 abcdyes++;
					if (eval("formObj.ReasonForRequest[5].checked") == true)
						abcdyes++;
					if (eval("formObj.ReasonForRequest[6].checked") == true)
						abcdyes++;
        
					if  (abcdyes==0){
    
						alert("You have selected Code Reservation.  You must select either a), b), or c). Please select one and submit again");
                        formObj.NXXUpdate.focus();
      					return false;
					}  
				  
				  }	
			if (formObj.TypeOfRequest[2].checked){	
					if (formObj.ReasonForRequest[5].checked){	 	  
							if (eval("formObj.CodeRequestNew[0].checked") == true)
								 crnyes++;
							if (eval("formObj.CodeRequestNew[1].checked") == true)
								crnyes++;
					        
							if  (crnyes==0){
    
								alert("You must fill out Section 1.7 and submit again");
								formObj.RequestNewNecessary.focus();
								return false;
								}
							}
							if (formObj.CodeRequestNew[0].checked){
								if (formObj.RequestNewNecessary.value == ""){
								    alert("You have selected that your CO Code is needed for distinct routing. Please enter your explanation and submit again");
							        formObj.RequestNewNecessary.focus();
							         return false;
								          }
								      }
							 if (formObj.CodeRequestNew[1].checked){
								if (formObj.RequestNewOther.value == ""){
							           alert("You have selected that your CO Code is needed for other reasons.  Please enter your explanation and submit again");
							          formObj.RequestNewOther.focus();
								       return false;
											 }	
										}  
						
				if (formObj.ReasonForRequest[6].checked){	 	  
				  		if (eval("formObj.NPAinJeopardy[0].checked") == true)
						 jep1yes++;
						if (eval("formObj.NPAinJeopardy[1].checked") == true)
						jep1yes++;
					        
						if  (jep1yes==0){
    
						alert("You must fill out Section 1.6 and submit again");
			            formObj.NXXGrowthCal.focus();
						return false;
						}
					
				}
            }
           if ((formObj.ReasonForRequest[0].checked)||(formObj.ReasonForRequest[4].checked)){
			jep1yes=0
			//alert("Reaseon 4 chk");
			if (eval("formObj.NPAinJeopardy[0].checked") == true){
				jep1yes++;
 				//(formObj.NPAinJeopardy[0].checked+"jep 0 ");
 			}
            if (eval("formObj.NPAinJeopardy[1].checked") == true){
				//alert(formObj.NPAinJeopardy[1].checked+"jep 1");
				jep1yes++;
			}
			if (eval("formObj.CodeRequestNew[0].checked") == true){
				//alert(formObj.CodeRequestNew[0].checked+"cr 0");
				jep1yes++;
			}
            if (eval("formObj.CodeRequestNew[1].checked") == true){
				//alert(formObj.CodeRequestNew[1].checked+"cr 0");
				jep1yes++;
			}
			if  (jep1yes!=0){
				formObj.NPAinJeopardy[0].checked = false;
				formObj.NPAinJeopardy[1].checked = false;
				formObj.CodeRequestNew[0].checked = false;
				formObj.CodeRequestNew[1].checked = false;
				alert("You have selected Intitial Code from 1.5.  Please leave the selections from 1.6 and 1.7 blank.");
				return false;
			}
           }
        if ((formObj.ReasonForRequest[1].checked)||(formObj.ReasonForRequest[5].checked)){
			jep1yes=0
			//alert("Reaseon 4 chk");
			if (eval("formObj.NPAinJeopardy[0].checked") == true){
				jep1yes++;
 				//alert(formObj.NPAinJeopardy[0].checked+"jep 0 ");
 			}
            if (eval("formObj.NPAinJeopardy[1].checked") == true){
				//alert(formObj.NPAinJeopardy[1].checked+"jep 1");
				jep1yes++;
			}
/*
Commented out by G. Brown on Aug 12/99 because this is not true.
Section 1.8 may still need to be completed by the applicant.

			if (eval("formObj.AuthorizationPart2[0].checked") == true){
				//alert(formObj.AuthorizationPart2[0].checked+"cr 0");
				jep1yes++;
			}
            if (eval("formObj.AuthorizationPart2[1].checked") == true){
				//alert(formObj.AuthorizationPart2[1].checked+"cr 0");
				jep1yes++;
			}
*/
			if  (jep1yes!=0){
				formObj.NPAinJeopardy[0].checked = false;
				formObj.NPAinJeopardy[1].checked = false;
/*
Commented out by G. Brown on Aug 12/99 because this is not true.
Section 1.8 may still need to be completed by the applicant.

				formObj.AuthorizationPart2[0].checked = false;
				formObj.AuthorizationPart2[1].checked = false;
				alert("You have selected Code Request from 1.5.  Please leave the selections from 1.6 and 1.8 blank.");
*/
				alert("You have selected Code Request from 1.5.  Please leave selection 1.6 blank.");

				return false;
			}
           }
		if ((formObj.ReasonForRequest[2].checked)||(formObj.ReasonForRequest[4].checked)){
			jep1yes=0
			//alert("Reaseon 4 chk");
			if (eval("formObj.CodeRequestNew[0].checked") == true){
				jep1yes++;
 				//alert(formObj.CodeRequestNew[0].checked+"jep 0 ");
 			}
            if (eval("formObj.CodeRequestNew[1].checked") == true){
				//alert(formObj.CodeRequestNew[1].checked+"jep 1");
				jep1yes++;
			}
/*
Commented out by G. Brown on Aug 12/99 because this is not true.
Section 1.8 may still need to be completed by the applicant.

			if (eval("formObj.AuthorizationPart2[0].checked") == true){
				//alert(formObj.AuthorizationPart2[0].checked+"cr 0");
				jep1yes++;
			}
            if (eval("formObj.AuthorizationPart2[1].checked") == true){
				//alert(formObj.AuthorizationPart2[1].checked+"cr 0");
				jep1yes++;
			}
*/
			if  (jep1yes!=0){
				formObj.CodeRequestNew[0].checked = false;
				formObj.CodeRequestNew[1].checked = false;
/*
Commented out by G. Brown on Aug 12/99 because this is not true.
Section 1.8 may still need to be completed by the applicant.

				formObj.AuthorizationPart2[0].checked = false;
				formObj.AuthorizationPart2[1].checked = false;
				alert("You have selected Additional Code Growth from 1.5.  Please leave the selections from 1.7 and 1.8 blank.");
*/
				alert("You have selected Additional Code Growth from 1.5.  Please leave selection 1.7.");

				return false;
			}
           }
       
        if ((formObj.NPAinJeopardy[0].checked)||(formObj.NPAinJeopardy[1].checked)){   
			return getValues_JS();
		}
		else
		{
		getValues_JS();
        }
		}
        // end hiding -->
// app-b    
function getValues_JS(){
var  formObj = document.FormP1app;
 jep1yes=0;
if (eval("formObj.NPAinJeopardy[0].checked") == true){jep1yes++;}
if (eval("formObj.NPAinJeopardy[1].checked") == true){jep1yes++;}
if  (jep1yes!=0)
{	//alert("You must fill out Section 1.6 and calculate again");	
	//return false;

//alert("getValues_JS");
if (document.FormP1app.NXXGrowthCal.value == "") {
    alert("You have not filled in the NXXs included field. Please type the NXXs included and calculate again");
	document.FormP1app.NXXGrowthCal.focus();
	return false;
}
if ((document.FormP1app.TNs.value == "")||(document.FormP1app.TNs.value == "0")) {
alert("You have not filled in the TNs available field. Please type the number of TNs included and calculate again");	
document.FormP1app.TNs.focus();
	return false; 
}
if (document.FormP1app.TNs.value != "") {
  if (isNaN(formObj.TNs.value)){
	alert("The TNs available field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.TNs.focus();
	return false;
  }
}

 
/*Check to see if the NPAinJeopardy radio button is set at all
 and if it is months 1- 12 must be filled with a valid number and with out skipping a month*/

 if ((document.FormP1app.Prev6Month1.value != "")||(document.FormP1app.Prev6Month1.value != "0")) {
  if (isNaN(formObj.Prev6Month1.value)){
	alert("This Previous 6 month Growth history field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.Prev6Month1.focus();
	return false;
  }
}
if ((document.FormP1app.Prev6Month2.value != "")||(document.FormP1app.Prev6Month2.value != "0")) {
  if (isNaN(formObj.Prev6Month2.value)){
	alert("This Previous 6 month Growth history field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.Prev6Month2.focus();
	return false;
  }
}
if ((document.FormP1app.Prev6Month3.value != "")||(document.FormP1app.Prev6Month3.value != "0")) {
  if (isNaN(formObj.Prev6Month3.value)){
	alert("This Previous 6 month Growth history field is not a number. Please type in a 1-9 digit value and calculate  again");	
	document.FormP1app.Prev6Month3.focus();
	return false;
  }
}
if ((document.FormP1app.Prev6Month4.value != "")||(document.FormP1app.Prev6Month4.value != "0")) {
  if (isNaN(formObj.Prev6Month4.value)){
	alert("This Previous 6 month Growth history field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.Prev6Month4.focus();
	return false;
  }
}
if ((document.FormP1app.Prev6Month5.value != "")||(document.FormP1app.Prev6Month5.value != "0")) {
  if (isNaN(formObj.Prev6Month5.value)){
	alert("This Previous 6 month Growth history field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.Prev6Month5.focus();
	return false;
  }
}
if ((document.FormP1app.Prev6Month6.value != "")||(document.FormP1app.Prev6Month6.value != "0")) {
  if (isNaN(formObj.Prev6Month6.value)){
	alert("This Previous 6 month Growth history field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.Prev6Month6.focus();
	return false;
  }
}
if (document.FormP1app.ProjGrowth16Month1.value != "") {
  if (isNaN(formObj.ProjGrowth16Month1.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth16Month1.focus();
	return false;
  }
}
/* The "0" was changed to "" in the section below because zero is a valid entry
This change occurred on Feb 2, 2000 by G. Brown
*/
if (((formObj.ProjGrowth16Month1.value == "")||(formObj.ProjGrowth16Month1.value == "0"))&&( formObj.ProjGrowth16Month2.value!= "0")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth16Month1.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth16Month2.value != "") {
  if (isNaN(formObj.ProjGrowth16Month2.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate  again");	
	document.FormP1app.ProjGrowth16Month2.focus();
	return false;
  }
}
if (((formObj.ProjGrowth16Month2.value == "")||(formObj.ProjGrowth16Month2.value == ""))&&( formObj.ProjGrowth16Month3.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth16Month2.focus();
	return false;

 }
 if (document.FormP1app.ProjGrowth16Month3.value != "") {
  if (isNaN(formObj.ProjGrowth16Month3.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth16Month3.focus();
	return false;
  }
}
if (((formObj.ProjGrowth16Month3.value == "")||(formObj.ProjGrowth16Month3.value == ""))&&( formObj.ProjGrowth16Month4.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth16Month3.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth16Month4.value != "") {
  if (isNaN(formObj.ProjGrowth16Month4.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth16Month4.focus();
	return false;
  }
}
if (((formObj.ProjGrowth16Month4.value == "")||(formObj.ProjGrowth16Month4.value == ""))&&( formObj.ProjGrowth16Month5.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth16Month4.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth16Month5.value != "") {
  if (isNaN(formObj.ProjGrowth16Month5.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth16Month5.focus();
	return false;
  }
}
if (((formObj.ProjGrowth16Month5.value == "")||(formObj.ProjGrowth16Month5.value == ""))&&( formObj.ProjGrowth16Month6.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth16Month5.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth16Month6.value != "") {
  if (isNaN(formObj.ProjGrowth16Month1.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth16Month1.focus();
	return false;
  }
}
if (((formObj.ProjGrowth16Month6.value == "")||(formObj.ProjGrowth16Month6.value == ""))&&( formObj.ProjGrowth712Month1.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth16Month6.focus();
	return false;
 }
 // for the months 7 - 12
 if (document.FormP1app.ProjGrowth712Month1.value != "") {
  if (isNaN(formObj.ProjGrowth712Month1.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth712Month1.focus();
	return false;
  }
}
if (((formObj.ProjGrowth712Month1.value == "") ||(formObj.ProjGrowth712Month1.value == ""))&&( formObj.ProjGrowth712Month2.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth712Month1.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth712Month2.value != "") {
  if (isNaN(formObj.ProjGrowth712Month2.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth712Month2.focus();
	return false;
  }
}
if (((formObj.ProjGrowth712Month2.value == "") ||(formObj.ProjGrowth712Month2.value == ""))&&( formObj.ProjGrowth712Month3.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth712Month2.focus();
	return false;
 }
if (document.FormP1app.ProjGrowth712Month3.value != "") {
  if (isNaN(formObj.ProjGrowth712Month3.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth712Month3.focus();
	return false;
  }
}
if (((formObj.ProjGrowth712Month3.value == "") ||(formObj.ProjGrowth712Month3.value == ""))&&( formObj.ProjGrowth712Month4.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth712Month3.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth712Month4.value != "") {
  if (isNaN(formObj.ProjGrowth712Month4.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth712Month4.focus();
	return false;
  }
}
 if (((formObj.ProjGrowth712Month4.value == "") ||(formObj.ProjGrowth712Month4.value == ""))&&( formObj.ProjGrowth712Month5.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth712Month4.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth712Month5.value != "") {
  if (isNaN(formObj.ProjGrowth712Month5.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth712Month5.focus();
	return false;
  }
}
 if (((formObj.ProjGrowth712Month5.value == "") ||(formObj.ProjGrowth712Month5.value == ""))&&( formObj.ProjGrowth712Month6.value!= "")){ 
	alert("Months 1-12 must be filled in order leaveing no blank fields between values");
	formObj.ProjGrowth712Month5.focus();
	return false;
 }
 if (document.FormP1app.ProjGrowth712Month6.value != "") {
  if (isNaN(formObj.ProjGrowth712Month6.value)){
	alert("This Projected Growth field is not a number. Please type in a 1-9 digit value and calculate again");	
	document.FormP1app.ProjGrowth712Month6.focus();
	return false;
  }
}
 /*Check to see if the NPAinJeopardy radio button is set to NON-Jeopardy
 and if it is months 1- 6 must be filled in */
  NPAinJ = eval("document.FormP1app.NPAinJeopardy[0].checked");

 if (((formObj.ProjGrowth16Month1.value == "") ||(formObj.ProjGrowth16Month1.value == ""))&&(NPAinJ==true)){ 
	alert("Months 1-7 must be filled for Non-Jeopardy.");
	formObj.ProjGrowth16Month1.focus();
	return false;
 }
 if ((( formObj.ProjGrowth16Month2.value== "") ||(formObj.ProjGrowth16Month2.value == ""))&&(NPAinJ==true)){
	alert("Months 1-7 must be filled for Non-Jeopardy.");
	formObj.ProjGrowth16Month2.focus();
	return false;
 }
 if (((formObj.ProjGrowth16Month3.value == "") ||(formObj.ProjGrowth16Month3.value == ""))&&(NPAinJ==true)){ 
	alert("Months 1-7 must be filled for Non-Jeopardy.");
	formObj.ProjGrowth16Month3.focus();
	return false;
 }
 if (((formObj.ProjGrowth16Month4.value == "") ||(formObj.ProjGrowth16Month4.value == ""))&&(NPAinJ==true)){ 
	alert("Months 1-7 must be filled for Non-Jeopardy.");
	formObj.ProjGrowth16Month4.focus();
	return false;
 }
 if (((formObj.ProjGrowth16Month5.value == "") ||(formObj.ProjGrowth16Month5.value == ""))&&(NPAinJ==true)){ 
	alert("Months 1-7 must be filled for Non-Jeopardy." );
	formObj.ProjGrowth16Month5.focus();
	return false;
 }
 if (((formObj.ProjGrowth16Month6.value == "") ||(formObj.ProjGrowth16Month6.value == ""))&&(NPAinJ==true)){ 
	alert("Months 1-7 must be filled for Non-Jeopardy." );
	formObj.ProjGrowth16Month6.focus();
	return false;
 }
 if (((formObj.ProjGrowth712Month1.value == "") ||(formObj.ProjGrowth712Month1.value == ""))&&(NPAinJ==true)){ 
	alert("Months 1-7 must be filled for Non-Jeopardy." );
	formObj.ProjGrowth712Month1.focus();
	return false;
}
//make all blank fields '0'
/*if (isNaN(formObj.ProjGrowth16Month1.value)){formObj.ProjGrowth16Month1.value="0";}
if (isNaN(formObj.ProjGrowth16Month2.value)){formObj.ProjGrowth16Month2.value="0";}
if (isNaN(formObj.ProjGrowth16Month3.value)){formObj.ProjGrowth16Month3.value="0";}
if (isNaN(formObj.ProjGrowth16Month4.value)){formObj.ProjGrowth16Month4.value="0";}
if (isNaN(formObj.ProjGrowth16Month5.value)){formObj.ProjGrowth16Month5.value="0";}
if (isNaN(formObj.ProjGrowth16Month6.value)){formObj.ProjGrowth16Month6.value="0";}
if (isNaN(formObj.ProjGrowth712Month1.value)){formObj.ProjGrowth712Month1.value="0";}
if (isNaN(formObj.ProjGrowth712Month2.value)){formObj.ProjGrowth712Month2.value="0";}
if (isNaN(formObj.ProjGrowth712Month3.value)){formObj.ProjGrowth712Month3.value="0";}
if (isNaN(formObj.ProjGrowth712Month4.value)){formObj.ProjGrowth712Month4.value="0";}
if (isNaN(formObj.ProjGrowth712Month5.value)){formObj.ProjGrowth712Month5.value="0";}
if (isNaN(formObj.ProjGrowth712Month6.value)){formObj.ProjGrowth712Month6.value="0";}
*/
 monValArray = new Array(11);
 m1 = Number (formObj.ProjGrowth16Month1.value);
 m2 = Number (formObj.ProjGrowth16Month2.value);
 m3 = Number (formObj.ProjGrowth16Month3.value);
 m4 = Number (formObj.ProjGrowth16Month4.value);
 m5 = Number (formObj.ProjGrowth16Month5.value);
 m6 = Number (formObj.ProjGrowth16Month6.value);
 m7 = Number (formObj.ProjGrowth712Month1.value);
 m8 = Number (formObj.ProjGrowth712Month2.value);
 m9 = Number (formObj.ProjGrowth712Month3.value);
 m10 = Number (formObj.ProjGrowth712Month4.value);
 m11 = Number (formObj.ProjGrowth712Month5.value);
 m12 = Number (formObj.ProjGrowth712Month6.value);
 
 monValArray[0] = m1;
 monValArray[1] = m2;
 monValArray[2] = m3;
 monValArray[3] = m4;
 monValArray[4] = m5;
 monValArray[5] = m6;
 monValArray[6] = m7;
 monValArray[7] = m8;
 monValArray[8] = m9;
 monValArray[9] = m10;
 monValArray[10] = m11;
 monValArray[11] = m12;
 var divisor = 0;
 var totVal = 0;
  for (cnt=0;cnt <=11; cnt++) {
	if (monValArray[cnt] != 0)  {		
		divisor = divisor +1;
		totVal = totVal + monValArray[cnt];
	}	
 }
/*
Add on Mar 13, 2000 by G. Brown to allow for zero growth projection in average calculation
*/
if (NPAinJ==true) divisor=12;
if (NPAinJ==false) divisor=6;
avgGrowthRate = Number(totVal / divisor);
/*
Add on Feb 2, 2000 by G. Brown to allow for zero growth projection
*/
if (totVal < 1) avgGrowthRate=0;
 document.FormP1app.AvgMonGrowthRate.value = avgGrowthRate;
 tns = Number(document.FormP1app.TNs.value);
 mte = Number(tns/avgGrowthRate);
/*
Add on Feb 2, 2000 by G. Brown to allow for zero growth projection
*/
if (avgGrowthRate < 1) mte=0;
 document.FormP1app.MonthsToExhaust.value = mte;
 return true;
}
else{
document.FormP1app.NXXGrowthCal.value =""
document.FormP1app.TNs.value = 0;
formObj.Prev6Month1.value=0;
formObj.Prev6Month2.value=0;
formObj.Prev6Month3.value=0;
formObj.Prev6Month4.value=0;
formObj.Prev6Month5.value=0;
formObj.Prev6Month6.value=0;
formObj.ProjGrowth16Month1.value=0;
formObj.ProjGrowth16Month2.value=0;
formObj.ProjGrowth16Month3.value=0;
formObj.ProjGrowth16Month4.value=0;
formObj.ProjGrowth16Month5.value=0;
formObj.ProjGrowth16Month6.value=0;
formObj.ProjGrowth712Month1.value=0;
formObj.ProjGrowth712Month2.value=0;
formObj.ProjGrowth712Month3.value=0;
formObj.ProjGrowth712Month4.value=0;
formObj.ProjGrowth712Month5.value=0;
formObj.ProjGrowth712Month6.value=0;
document.FormP1app.AvgMonGrowthRate.value =0;
document.FormP1app.MonthsToExhaust.value = 0;
   }
}

function disableFields() {
var  formObj = document.FormP1app;
document.FormP1app.ProjGrowth712Month1.value="0";
formObj.ProjGrowth712Month2.value="0";
formObj.ProjGrowth712Month3.value="0";
formObj.ProjGrowth712Month4.value="0";
formObj.ProjGrowth712Month5.value="0";
formObj.ProjGrowth712Month6.value="0";

document.FormP1app.ProjGrowth712Month1.readOnly=true;
formObj.ProjGrowth712Month2.readOnly=true;
formObj.ProjGrowth712Month3.readOnly=true;
formObj.ProjGrowth712Month4.readOnly=true;
formObj.ProjGrowth712Month5.readOnly=true;
formObj.ProjGrowth712Month6.readOnly=true;
}
function inableFields() {
var  formObj = document.FormP1app;

document.FormP1app.ProjGrowth712Month1.readOnly=false;
formObj.ProjGrowth712Month2.readOnly=false;
formObj.ProjGrowth712Month3.readOnly=false;
formObj.ProjGrowth712Month4.readOnly=false;
formObj.ProjGrowth712Month5.readOnly=false;
formObj.ProjGrowth712Month6.readOnly=false;
}
</script>
 <meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%
SelectedNPA=Request.Form("NPA")
LATA=888
session("SelNpa")=SelectedNPA
'Response.Write "SelectedNPA"
'Response.Write session("SelNpa")
Dim P1Days, P1getDays
uname = session("UserUserName")
Sub btnGoToMainFrm_onclick()
	Response.Redirect "xca_MenuSubPost.asp"
End Sub

'check to see where data coming from
 BlankP1=session("BlankP1")
 
Sub setApplicantDate()
	ApplicationDate.value= Date()
	ApplicationDate.disabled = true
	
End Sub

Select case BlankP1


case "Applicant"
	SelectedNPA=Request.Form("NPA")
	SelectedTypeOfReq=Request.Form("SelectedTypeOfReq")
	UserEntityID =int(session("UserEntityID"))
	session("P1UserEntityID")=UserEntityID
	setApplicantDate()

case "Admin"
	SelectedNPA=Request.Form("NPA")
	SelectedEntityName=Request.Form("EntityName")
	SelectedTypeOfReq=Request.Form("SelectedTypeOfReq")
	'Response.Write SelectedEntityName&"<BR>"
		If SelectedEntityName <> "" then
		sqlcheckEntity="Select * from xca_Entity,xca_User where xca_Entity.EntityName='"&SelectedEntityName&"' "
			GetUserEntityName.setSQLText(sqlcheckEntity)
			GetUserEntityName.Open
			checkEntityName= GetUserEntityName.fields.getValue("EntityName")
			UserEntityID= GetUserEntityName.fields.getValue("EntityID")
			'session("P1UserEntityID")=UserEntityID
			GetUserEntityName.close
		sqlcheckSelectedEntity="Select * from xca_Entity where EntityName='"&SelectedEntityName&"' "
			GetSelectedEntityName.setSQLText(sqlcheckSelectedEntity)
			GetSelectedEntityName.Open
'			checkEntityName= GetSelectedEntityName.fields.getValue("EntityName")
			SelectedEntityID= GetSelectedEntityName.fields.getValue("EntityID")
			UserEntityID=SelectedEntityID
			session("P1SelectedEntityID")=SelectedEntityID
'			Response.Write SelectedEntityID&"<BR>"
			GetUserEntityName.close
	if checkEntityName="" then	
							session("NoTixSent")="DidNotSend"
							Response.Redirect session("Here")
					end if
		else
			session("NoTixSent")="DidNotSend"
			Response.Redirect session("Here")			
		End if
case ""
	session("NoTixSent")="DidNotSend"
	Response.Redirect session("Here")	

End Select
session("P1UserEntityID")=UserEntityID
'Response.Write UserEntityID
	session("NoTixSent")=""

AdminData=session("ADMIN")


'get Admin info for top of form
sqlADMIN="Select * from xca_Entity where EntityName ='"&AdminData&"'"
	GetAdminEntityName.setSQLText(sqlADMIN)
	 GetAdminEntityName.Open

'get spare NXX from NPA selected	Rec1.fields.getValue("NewNPA")

' April 21 2006
'sql = "Select * from xca_Entity,xca_User where xca_Entity.EntityID = '"&UserEntityID&"' and xca_User.UserName= '"&uname&"' "
sql = "Select * From xca_Entity Inner Join xca_User On xca_Entity.EntityID=xca_User.EntityID where xca_Entity.EntityID = '"&UserEntityID&"' and xca_User.UserName= '"&uname&"' "
'sql = "Select * from xca_Entity,xca_User where xca_Entity.EntityID = '"&UserEntityID&"' and xca_User.UserName= '"&uname&"' "
   GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open


session("Here")="xca_Part1app.asp"
session("P1CONPA")=SelectedNPA
'
' Oct 1, 2001
'
' July 14, 2003
' Nov 10, 2005 add NPA 226 and 438
If SelectedNPA=778 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=604 Order by RateCenter ASC"
ElseIf SelectedNPA=289 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=905 Order by RateCenter ASC"
ElseIf SelectedNPA=226 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=519 Order by RateCenter ASC"
ElseIf SelectedNPA=438 Then
sqlRC="Select Distinct RateCenter From xca_COCode where NPA=514 Order by RateCenter ASC"
Else
sqlRC="Select Distinct RateCenter From xca_COCode where NPA='"&SelectedNPA&"' Order by RateCenter ASC"
end if
' End July 14, 2003
RateCenterAssignLookup.SetSQLText(sqlRC)
RateCenterAssignLookup.Open
'RateCenterAssignLookup.fields.getValue("RateCenter"),RateCenterAssignLookup.fields.getValue("Ra't'eCenter"),cntRC-1
'
' End Oct 1, 2001
'
'action = session("P1TypeOReq")
select case SelectedTypeOfReq 
case "A"	'Assign
sql1 = "Select  NXX from xca_COCode where (Status='S' and NPA='"&SelectedNPA&"') or (Status='R' and NPA='"&SelectedNPA&"' and  EntityID = '"&UserEntityID&"')ORDER BY NXX ASC"

   Part1NXXAssignLookup.setSQLText(sql1)
	 Part1NXXAssignLookup.Open
	 'Part1NXXAssignLookup.moveFirst
'	const NxxCnt = cint(Part1NXXAssignLookup.getCount)
'	 dim rowNPAVal(NxxCnt)
'	for cnt = 1 to NxxCnt
'		if not Part1NXXAssignLookup.EOF	then
'		rowNPAVal(cnt)=Part1NXXAssignLookup.fields.getValue("NXX")
'		Part1NXXAssignLookup.moveNext		
'		response.write rowNPAVal(cnt)&"<br>"
'		end if
'	 next
'Session("NPAArray") = rowNPAVal


'	  for cnt = 1 to 20
	'	while not Part1NXXAssignLookup.EOF		
'		NXXAssign_.addItem Part1NXXAssignLookup.fields.getValue("NXX"),Part1NXXAssignLookup.fields.getValue("NXX"),cnt-1
'		Part1NXXAssignLookup.moveNext
'		txt = txt& cnt & " "
'		NXXAssign.value = txt
	'	wend
'	 next
	 
	  
	 'NXXAssign_.setValue Test,0
	 

	 	 	 
	 readOnlyA="" 
	 readOnlyU="disabled"
	 readOnlyR="disabled"
	 
	 checkedA="checked"
	 checkedU=""
	 checkedR=""
case "U"	'Update
'get updateable NXXs
sqlUpdate = "Select  NXX from xca_COCode where Status='I' and  NPA='"&SelectedNPA&"' and EntityID = '"&UserEntityID&"' ORDER BY NXX ASC"

	Part1NXXUpdateLook.setSQLText(sqlUpdate)
	 Part1NXXUpdateLook.Open
	
	
	 readOnlyA="disabled" 
	 readOnlyU=""
	 readOnlyR="disabled"
	 
	 checkedA=""
	 checkedU="checked"
	 checkedR=""
	 
case "R"	'Reserve
sqlReserve = "Select  NXX from xca_COCode where Status='S' and NPA='"&SelectedNPA&"' ORDER BY NXX ASC"

	 Part1NXXReserveLook.setSQLText(sqlReserve)
	 Part1NXXReserveLook.Open 
	 
	 readOnlyA="disabled" 
	 readOnlyU="disabled"
	 readOnlyR=""
	 
	 checkedA=""
	 checkedU=""
	 checkedR="checked"
	 
case else
end select
	
sqlParm = "Select * from xca_Parms where name='P1DAYS'"

	P1Parms.setSQLText(sqlParm)
	P1Parms.Open
	P1getDays= P1Parms.fields.getValue("Value")
	Part1Days.setCaption(P1getDays)
	P1Parms.close
	
	
sub fill() 
	sql1 = "Select  NXX from xca_COCode where (Status='S' and NPA='"&SelectedNPA&"') or (Status='R' and NPA='"&SelectedNPA&"' and  EntityID = '"&UserEntityID&"')ORDER BY NXX ASC"

	 Part1NXXAssignLookup.setSQLText(sql1)
	 Part1NXXAssignLookup.Open
'	 NxxCnt = cint(Part1NXXAssignLookup.getCount)
	 for cnt = 1 to NxxCnt
		'while not Part1NXXAssignLookup.EOF		
		txtVal = cstr(Part1NXXAssignLookup.fields.getValue("NXX"))
		NXXAssign_.addItem txtVal,txtVal,cnt-1
		Part1NXXAssignLookup.moveNext
		
		'NXXAssign.value = txt
	next

	
end sub
'call fill() 


 %>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() {
getValues_JS();

}

function NPAinJeopardy_y_onclick() {
disableFields();
}
function NPAinJeopardy_n_onclick() {
inableFields();
}
function button2_onclick() {
//window_onload();
window.open("xca_NXXLookUp.asp",null,"fullscreen=NO,resizable=yes,scrollbars= yes ,width=150,height=400,top=20,left=5")
}

function window_onload() {
//FormP1app.target="_blank";
//FormP1app.action ="xca_NXXLookUp.asp" ;
//FormP1app.method="post";
//FormP1app.submit;
//window.open("xca_NXXLookUp.asp",null,"fullscreen=NO,resizable=yes,scrollbars= yes ,width=150,height=400,top=20,left=5")
}

//-->
</SCRIPT>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript >
Sub FillNXXList
	optOutput = ""
	for cnt = 1 to 20
		'while not Part1NXXAssignLookup.EOF		
		'txtVal = cstr(rowNPAVal(cnt))
		'document.FormP1app.NXXList
		
		
		optOutput = optOutput & "<option value = "&cnt&">"&cnt&"<\option>"
		
		txt = txt& cnt & " "
		
	next


'locArr = session("NPAArray") 
document.FormP1app.NXXAssign.value = "locArr(3)"
'document.FormP1app.NXXAssign.value = optOutput
End Sub
Sub AssignListButton_onclick()
'call fill()
'FillNXXList()
'xca_Part1app.navigate.fill


End Sub


</SCRIPT>
</head>
<FORM>
<body leftmargin="20" rightmargin="20" bgColor="#d7c7a4" text="black" LANGUAGE=javascript onload="return window_onload()">

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=P1Parms 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sValue\sfrom\sxca_Parms\swhere\sName=?\q,TCControlID_Unmatched=\qP1Parms\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sValue\sfrom\sxca_Parms\swhere\sName=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q25\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersP1Parms()
{
}
function _initP1Parms()
{
	P1Parms.advise(RS_ONBEFOREOPEN, _setParametersP1Parms);
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
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserEntityName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetSelectedEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetSelectedEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetSelectedEntityName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetSelectedEntityName.setRecordSource(rsTmp);
}
function _GetSelectedEntityName_ctor()
{
	CreateRecordset('GetSelectedEntityName', _initGetSelectedEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetAdminEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminEntityName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
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
<!-- -->
<!-- Oct 1 -->
<!-- -->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RateCenterAssignLookup 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sRateCenter\sfrom\sxca_COCode\swhere\sNPA=?\q,TCControlID_Unmatched=\qRateCenterAssignLookup\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sRateCenter\sfrom\sxca_COCode\swhere\sNPA=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRateCenterAssignLookup()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select Distinct RateCenter from xca_COCode where NPA=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RateCenterAssignLookup.setRecordSource(rsTmp);
	if (thisPage.getState('pb_RateCenterAssignLookup') != null)
RateCenterAssignLookup.setBookmark(thisPage.getState('pb_RateCenterAssignLookup'));
}
function _RateCenterAssignLookup_ctor()
{
	CreateRecordset('RateCenterAssignLookup', _initRateCenterAssignLookup, null);
}
function _RateCenterAssignLookup_dtor()
{
	RateCenterAssignLookup._preserveState();
	thisPage.setState('pb_RateCenterAssignLookup', RateCenterAssignLookup.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<!-- -->
<!-- End Oct 1 -->
<!-- -->


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Part1NXXAssignLookup 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sNXX\sfrom\sxca_COCode\swhere\sStatus='S'\sand\sNPA=?\q,TCControlID_Unmatched=\qPart1NXXAssignLookup\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sNXX\sfrom\sxca_COCode\swhere\sStatus='S'\sand\sNPA=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initPart1NXXAssignLookup()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select Distinct NXX from xca_COCode where Status=\'S\' and NPA=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Part1NXXAssignLookup.setRecordSource(rsTmp);
	if (thisPage.getState('pb_Part1NXXAssignLookup') != null)
		Part1NXXAssignLookup.setBookmark(thisPage.getState('pb_Part1NXXAssignLookup'));
}
function _Part1NXXAssignLookup_ctor()
{
	CreateRecordset('Part1NXXAssignLookup', _initPart1NXXAssignLookup, null);
}
function _Part1NXXAssignLookup_dtor()
{
	Part1NXXAssignLookup._preserveState();
	thisPage.setState('pb_Part1NXXAssignLookup', Part1NXXAssignLookup.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Part1NXXUpdateLook 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sNXX\sfrom\sxca_COCode\swhere\sStatus='S'\sand\sNPA=?\q,TCControlID_Unmatched=\qPart1NXXUpdateLook\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sNXX\sfrom\sxca_COCode\swhere\sStatus='S'\sand\sNPA=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initPart1NXXUpdateLook()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select Distinct NXX from xca_COCode where Status=\'S\' and NPA=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Part1NXXUpdateLook.setRecordSource(rsTmp);
	if (thisPage.getState('pb_Part1NXXUpdateLook') != null)
		Part1NXXUpdateLook.setBookmark(thisPage.getState('pb_Part1NXXUpdateLook'));
}
function _Part1NXXUpdateLook_ctor()
{
	CreateRecordset('Part1NXXUpdateLook', _initPart1NXXUpdateLook, null);
}
function _Part1NXXUpdateLook_dtor()
{
	Part1NXXUpdateLook._preserveState();
	thisPage.setState('pb_Part1NXXUpdateLook', Part1NXXUpdateLook.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Part1NXXReserveLook 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sNXX\sfrom\sxca_COCode\swhere\sStatus='S'\sand\sNPA=?\q,TCControlID_Unmatched=\qPart1NXXReserveLook\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sNXX\sfrom\sxca_COCode\swhere\sStatus='S'\sand\sNPA=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initPart1NXXReserveLook()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select Distinct NXX from xca_COCode where Status=\'S\' and NPA=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Part1NXXReserveLook.setRecordSource(rsTmp);
	if (thisPage.getState('pb_Part1NXXReserveLook') != null)
		Part1NXXReserveLook.setBookmark(thisPage.getState('pb_Part1NXXReserveLook'));
}
function _Part1NXXReserveLook_ctor()
{
	CreateRecordset('Part1NXXReserveLook', _initPart1NXXReserveLook, null);
}
function _Part1NXXReserveLook_dtor()
{
	Part1NXXReserveLook._preserveState();
	thisPage.setState('pb_Part1NXXReserveLook', Part1NXXReserveLook.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart1Data style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxcaPart1.Tix\s=?\q,TCControlID_Unmatched=\qGetPart1Data\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxcaPart1.Tix\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart1Data()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
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


<table border="0" cellpadding="0"><tr>
	<td wrap><font color="maroon" face="Arial Black" size="4"><strong>
Part 1 - 
            Canadian Central Office Code (NXX) Assignment Request 
            Form</strong></font>
            </td></tr>
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
<p>The Code Applicant and the Code Administrator acknowledge 
that the information contained on this request form is sensitive and will be 
treated as confidential. Prior to confirmation the information in this form will 
only be shared with the appropriate administrator and/or regulators. Information 
requested for RDBS and BRIDS will become available to the public upon input into 
those systems.</p>
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
<td wrap>
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
<td align="right" wrap><label><font face="arial" size="2"><strong>Authorized Representative 
            Name:&nbsp;&nbsp;</strong></font></label></td>
<td align="left" wrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=AuthorizedRep 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="AuthorizedRep">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AuthorizedRep">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRep()
{
	AuthorizedRep.setStyle(TXT_TEXTBOX);
	AuthorizedRep.setDataSource(GetPart1Data);
	AuthorizedRep.setDataField('AuthorizedRep');
	AuthorizedRep.setMaxLength(35);
	AuthorizedRep.setColumnCount(35);
}
function _AuthorizedRep_ctor()
{
	CreateTextbox('AuthorizedRep', _initAuthorizedRep, null);
}
</script>
<% AuthorizedRep.display %>

<!--METADATA TYPE="DesignerControl" endspan-->          
</td></tr>
<tr>
<td align="right" wrap><label><font face="arial" size="2"><strong>Title:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=AuthorizedRepTitle 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="AuthorizedRepTitle">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AuthorizedRepTitle">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRepTitle()
{
	AuthorizedRepTitle.setStyle(TXT_TEXTBOX);
	AuthorizedRepTitle.setDataSource(GetPart1Data);
	AuthorizedRepTitle.setDataField('AuthorizedRepTitle');
	AuthorizedRepTitle.setMaxLength(35);
	AuthorizedRepTitle.setColumnCount(35);
}
function _AuthorizedRepTitle_ctor()
{
	CreateTextbox('AuthorizedRepTitle', _initAuthorizedRepTitle, null);
}
</script>
<% AuthorizedRepTitle.display %>

<!--METADATA TYPE="DesignerControl" endspan-->        
</td></tr>
<tr>
<td align="right" wrap><label><font face="arial" size="2"><strong>Date of 
            Receipt:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <%
Response.Write date()
%>
            

</td></tr>
</table>
<br><br>
<br><br>
<strong><center><font size="4" face="arial" color="#993300">General Information</font></strong></CENTER>
<table align="left" border="0" cellPadding="0" cellSpacing="1">
<tr>
        <td wrap style="FONT-WEIGHT: bold"><label><strong><font size="3" face="arial" color="#993300">1.1 Contact 
            Information:</font></strong></label> 
 
 </td></tr>
 
 </table>
 <br>
 <br>


<table align="center" border="0" cellPadding="1" cellSpacing="1">
    <tbody>
    
    <tr>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code Applicant 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left" wrap><font face="Arial"> </font>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Entity 
            Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityname 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityname">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Entity Name 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityName 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityName">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityName()
{
	AdminEntityName.setDataSource(GetAdminEntityName);
	AdminEntityName.setDataField('EntityName');
}
function _AdminEntityName_ctor()
{
	CreateLabel('AdminEntityName', _initAdminEntityName, null);
}
</script>
<% AdminEntityName.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityContact 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityContact">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityContact 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 76px" width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityAddress 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityAddress()
{
	AppEntityAddress.setDataSource(GetUserEntityName);
	AppEntityAddress.setDataField('UserAddress');
}
function _AppEntityAddress_ctor()
{
	CreateLabel('AppEntityAddress', _initAppEntityAddress, null);
}
</script>
<% AppEntityAddress.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityAddress 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityCity 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityCity">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityCity()
{
	AppEntityCity.setDataSource(GetUserEntityName);
	AppEntityCity.setDataField('UserCity');
}
function _AppEntityCity_ctor()
{
	CreateLabel('AppEntityCity', _initAppEntityCity, null);
}
</script>
<% AppEntityCity.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityCity 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityProvince 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityProvince()
{
	AppEntityProvince.setDataSource(GetUserEntityName);
	AppEntityProvince.setDataField('UserProvince');
}
function _AppEntityProvince_ctor()
{
	CreateLabel('AppEntityProvince', _initAppEntityProvince, null);
}
</script>
<% AppEntityProvince.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityProvince 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
         
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal 
            Code</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityPostalCode 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityPostalCode()
{
	AppEntityPostalCode.setDataSource(GetUserEntityName);
	AppEntityPostalCode.setDataField('UserPostalCode');
}
function _AppEntityPostalCode_ctor()
{
	CreateLabel('AppEntityPostalCode', _initAppEntityPostalCode, null);
}
</script>
<% AppEntityPostalCode.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal Code 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityPostalCode 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
           
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail Address 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityEmail 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 59px" width=59>
	<PARAM NAME="_ExtentX" VALUE="1561">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityEmail()
{
	AppEntityEmail.setDataSource(GetUserEntityName);
	AppEntityEmail.setDataField('UserEmail');
}
function _AppEntityEmail_ctor()
{
	CreateLabel('AppEntityEmail', _initAppEntityEmail, null);
}
</script>
<% AppEntityEmail.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityEmail 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>    
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityFax 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 48px" width=48>
	<PARAM NAME="_ExtentX" VALUE="1270">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityFax">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityFax()
{
	AppEntityFax.setDataSource(GetUserEntityName);
	AppEntityFax.setDataField('UserFax');
}
function _AppEntityFax_ctor()
{
	CreateLabel('AppEntityFax', _initAppEntityFax, null);
}
</script>
<% AppEntityFax.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityFax 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 52px" width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
          
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityTelephone 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityTelephone()
{
	AppEntityTelephone.setDataSource(GetUserEntityName);
	AppEntityTelephone.setDataField('UserTelephone');
}
function _AppEntityTelephone_ctor()
{
	CreateLabel('AppEntityTelephone', _initAppEntityTelephone, null);
}
</script>
<% AppEntityTelephone.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityTelephone 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityExtension 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 84px" width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityExtension()
{
	AppEntityExtension.setDataSource(GetUserEntityName);
	AppEntityExtension.setDataField('UserExtension');
}
function _AppEntityExtension_ctor()
{
	CreateLabel('AppEntityExtension', _initAppEntityExtension, null);
}
</script>
<% AppEntityExtension.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityExtension 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 88px" width=88>
	<PARAM NAME="_ExtentX" VALUE="2328">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
           
</td></tr></tbody>

</table>


<br><br>

<table align="left" border="0" cellPadding="0" cellSpacing="0">
    
    <tr>
        <td align="left" colSpan="8"><strong><font face="arial" color="#993300" size="3">
	1.2 
            CO Code Information:</font></strong>
    <tr>
        <td align="right" colSpan="8">
            <div align="left">&nbsp; </div>
    <tr>
        <td align="left" colSpan="2" width="100"><strong><font face="arial" size="2">&nbsp;NPA:&nbsp; 
           <font color="blue">
            <%Response.Write(session("P1CONPA"))%></font></strong></FONT>
<!-- G. Brown Add in section for LATA Sept 26 2001 -->
            <td align="left" colSpan="2" width="100"><strong><font face="arial" size="2">&nbsp;LATA:&nbsp; 
           <font color="blue">
            <%Response.Write(LATA)%></font></strong></FONT>
<!-- End of new section for LATA-->
<!-- The old section has been removed. --> 
<!-- End of section commented out -->
        <td align="left" colSpan="4" width="100"><strong><font face="arial" size="2">&nbsp;OCN:&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=OCN style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 24px" 
	width=24>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="OCN">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="OCN">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="4">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOCN()
{
	OCN.setStyle(TXT_TEXTBOX);
	OCN.setDataSource(GetPart1Data);
	OCN.setDataField('OCN');
	OCN.setMaxLength(4);
	OCN.setColumnCount(4);
}
function _OCN_ctor()
{
	CreateTextbox('OCN', _initOCN, null);
}
</script>
<% OCN.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">Switch 
            Identification (Switching Entity / POI):&nbsp;&nbsp;</strong></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=SwitchID 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 66px" width=66>
	<PARAM NAME="_ExtentX" VALUE="1746">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="SwitchID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="11">
	<PARAM NAME="DisplayWidth" VALUE="18">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initSwitchID()
{
	SwitchID.setStyle(TXT_TEXTBOX);
	SwitchID.setDataSource(GetPart1Data);
	SwitchID.setDataField('SwitchID');
	SwitchID.setMaxLength(11);
	SwitchID.setColumnCount(18);
}
function _SwitchID_ctor()
{
	CreateTextbox('SwitchID', _initSwitchID, null);
}
</script>
<% SwitchID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="5">
        <td align="left" colSpan="2"><font face="Arial" size="2">This is an eleven-character descriptor of the 
            switch provided by the owning entity for the purpose of routing 
            calls. This is the 11 character COMMON LANGUAGE Location 
            Identification - (CLLI) of the switch or POI.</font>
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">
	City or Wire 
            Center:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=WireCenter 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 240px" width=240>
	<PARAM NAME="_ExtentX" VALUE="6350">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="WireCenter">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="WireCenter">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="40">
	<PARAM NAME="DisplayWidth" VALUE="40">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initWireCenter()
{
	WireCenter.setStyle(TXT_TEXTBOX);
	WireCenter.setDataSource(GetPart1Data);
	WireCenter.setDataField('WireCenter');
	WireCenter.setMaxLength(40);
	WireCenter.setColumnCount(40);
}
function _WireCenter_ctor()
{
	CreateTextbox('WireCenter', _initWireCenter, null);
}
</script>
<% WireCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!-- -->
<!-- Old RateCenter Section has been removed for RateCenter drop down menu -->
<!-- -->
<tr>
<td align="left" colSpan="7"><strong><font face="arial" size="2">Rate Center:&nbsp;&nbsp;</font></strong>
<!-- -->
<!-- Oct 1 -->
<!-- -->
<%
RateCenterAssignLookup.moveFirst
Response.Write"<SELECT id=RateCenterAssignLookup name=RateCenterAssignLookup>"
while (Not RateCenterAssignLookup.EOF) 
Response.Write"<OPTION>"  
Response.Write RateCenterAssignLookup.fields.getValue("RateCenter") 
Response.Write"</OPTION>"
RateCenterAssignLookup.moveNext
wend
Response.Write"</SELECT>"
%>
</SELECT>
<!-- -->
<!-- End Oct 1 -->
<!-- -->
<font face="Arial" size="2">Rate Center Name must be a tariffed Rate Center 
            associated with toll billing.</font>
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="2">Route Same 
            as<strong><font face="arial" size="2">&nbsp;NPA:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RouteNPA 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RouteNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RouteNPA">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNPA()
{
	RouteNPA.setStyle(TXT_TEXTBOX);
	RouteNPA.setDataSource(GetPart1Data);
	RouteNPA.setDataField('RouteNPA');
	RouteNPA.setMaxLength(3);
	RouteNPA.setColumnCount(3);
}
function _RouteNPA_ctor()
{
	CreateTextbox('RouteNPA', _initRouteNPA, null);
}
</script>
<% RouteNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <% 
RouteNPA.setDataField"222"
%>            
<strong><font face="Arial" size="2">&nbsp; NXX:&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RouteNXX 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RouteNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RouteNXX">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNXX()
{
	RouteNXX.setStyle(TXT_TEXTBOX);
	RouteNXX.setDataSource(GetPart1Data);
	RouteNXX.setDataField('RouteNXX');
	RouteNXX.setMaxLength(3);
	RouteNXX.setColumnCount(3);
}
function _RouteNXX_ctor()
{
	CreateTextbox('RouteNXX', _initRouteNXX, null);
}
</script>
<% RouteNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;<strong><font face="Arial" size="2">Use 
            Same Rate Center as<strong><font face="Arial" size="2">&nbsp;NPA:&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=CenterNPA 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="CenterNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CenterNPA">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNPA()
{
	CenterNPA.setStyle(TXT_TEXTBOX);
	CenterNPA.setDataSource(GetPart1Data);
	CenterNPA.setDataField('CenterNPA');
	CenterNPA.setMaxLength(3);
	CenterNPA.setColumnCount(3);
}
function _CenterNPA_ctor()
{
	CreateTextbox('CenterNPA', _initCenterNPA, null);
}
</script>
<% CenterNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<strong><font face="Arial" size="2">&nbsp; NXX:&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=CenterNXX 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="CenterNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CenterNXX">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNXX()
{
	CenterNXX.setStyle(TXT_TEXTBOX);
	CenterNXX.setDataSource(GetPart1Data);
	CenterNXX.setDataField('CenterNXX');
	CenterNXX.setMaxLength(3);
	CenterNXX.setColumnCount(3);
}
function _CenterNXX_ctor()
{
	CreateTextbox('CenterNXX', _initCenterNXX, null);
}
</script>
<% CenterNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></strong></font></strong></font></strong></font></strong></font> </strong>
    <tr>
        <td align="left" colSpan="7">&nbsp;&nbsp;
    <tr>
        <td align="left" colSpan="7"><strong><font face="arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
1.3 Dates:</font></strong>
    <tr>
        <td align="left" colSpan="7">&nbsp;
    <tr>
        <td align="left" colSpan="7"><strong><font face="Arial" size="2">Application 
Date:&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=ApplicationDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="ApplicationDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initApplicationDate()
{
	ApplicationDate.setStyle(TXT_TEXTBOX);
	ApplicationDate.setDataSource(GetPart1Data);
	ApplicationDate.setDataField('ApplicationDate');
	ApplicationDate.setMaxLength(10);
	ApplicationDate.setColumnCount(10);
}
function _ApplicationDate_ctor()
{
	CreateTextbox('ApplicationDate', _initApplicationDate, null);
}
</script>
<% ApplicationDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<font face="arial" size="1">dd/mm/ccyy</font></font></strong> 
    <tr>
        <td align="left" colSpan="7"><strong><font face="Arial" size="2"><strong>Requested Effective Date:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RequestedEffDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RequestedEffDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestedEffDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedEffDate()
{
	RequestedEffDate.setStyle(TXT_TEXTBOX);
	RequestedEffDate.setDataSource(GetPart1Data);
	RequestedEffDate.setDataField('RequestedEffDate');
	RequestedEffDate.setMaxLength(10);
	RequestedEffDate.setColumnCount(10);
}
function _RequestedEffDate_ctor()
{
	CreateTextbox('RequestedEffDate', _initRequestedEffDate, null);
}
</script>
<% RequestedEffDate.display %>

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
calendar days 
            from the submission of this form. It should be noted that 
            interconnection arrangements and facilities need to be in place 
            prior to activation of a code. Such arrangements are outside the 
            scope of these guidelines.</font></p>
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
	<td align="left" wrap colSpan="7">
</td></tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<br><br>
<table align="left" background ="" border="0" cellPadding="0" cellSpacing="0">
    <TBODY>
    
    <tr>
        <td align="left" colSpan="3"><strong><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
	1.4 Type of Entity Requesting the 
            Code:</font></strong> 
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;
<tr>
<td wrap align="left" colSpan="3"><strong><font face="Arial" size="2"> A)&nbsp;&nbsp;
<select Name="CarrierType">
	<option Value="l" selected>Local 
                    Exchange Carrier
	<option Value="w">Wireless Service 
Provider
	<option Value="o">Other (Specify)
</select></font></strong>&nbsp; 
<strong><font face="Arial" size="2">&nbsp; Other Explained:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=OtherCarrierType 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="OtherCarrierType">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="OtherCarrierType">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="30">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOtherCarrierType()
{
	OtherCarrierType.setStyle(TXT_TEXTBOX);
	OtherCarrierType.setDataSource(GetPart1Data);
	OtherCarrierType.setDataField('OtherCarrierType');
	OtherCarrierType.setMaxLength(30);
	OtherCarrierType.setColumnCount(30);
}
function _OtherCarrierType_ctor()
{
	CreateTextbox('OtherCarrierType', _initOtherCarrierType, null);
}
</script>
<% OtherCarrierType.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></strong>     

</td><tr></tr>
    <tr>
        <td align="left" colSpan="3" vAlign="top">&nbsp;


<tr>
        <td align="left" colSpan="3" vAlign="top"><font face="arial" size="2"><strong>B)&nbsp; Type of Service for which code is being 
            requested:</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=TypeOfService 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="TypeOfService">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="TypeOfService">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="100">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTypeOfService()
{
	TypeOfService.setStyle(TXT_TEXTBOX);
	TypeOfService.setDataSource(GetPart1Data);
	TypeOfService.setDataField('TypeOfService');
	TypeOfService.setMaxLength(100);
	TypeOfService.setColumnCount(50);
}
function _TypeOfService_ctor()
{
	CreateTextbox('TypeOfService', _initTypeOfService, null);
}
</script>
<% TypeOfService.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td align="left" colSpan="3">&nbsp;


<tr>
<td wrap align="left" colSpan="3"><strong><font face="Arial" size="2">C)&nbsp; Is certification or authorization required to provide 
            this type of service in the relevant geographic 
            area?</strong></FONT></td>
           
            </tr>
    <tr>
        <td width="25">
        <td colSpan="2">
			
                 <input type="radio" name="CertificationRequired" value="Y" CHECKED style="LEFT: -1px; TOP: 0px"><strong><font face="Arial"></strong> Yes</FONT>
                 <input type="radio" name="CertificationRequired" value="N"><strong><font face="Arial"></strong> No</FONT>
<tr>
<td wrap width="25"></td>
        <td colSpan="2"><strong><font face="Arial" size="2">(1)&nbsp; If no, 
            explain:</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=CertificationNoExplained 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="CertificationNoExplained">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CertificationNoExplained">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="100">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCertificationNoExplained()
{
	CertificationNoExplained.setStyle(TXT_TEXTBOX);
	CertificationNoExplained.setDataSource(GetPart1Data);
	CertificationNoExplained.setDataField('CertificationNoExplained');
	CertificationNoExplained.setMaxLength(100);
	CertificationNoExplained.setColumnCount(50);
}
function _CertificationNoExplained_ctor()
{
	CreateTextbox('CertificationNoExplained', _initCertificationNoExplained, null);
}
</script>
<% CertificationNoExplained.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 </td></tr>

<tr>

<td align="left" wrap><font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;</strong></font></td>
        <td align="left" colSpan="2"><font face="Arial" size="2"><strong>(2)&nbsp; If yes, 
            does your company have such certification or 
            authorization?</strong></font></td>
    <tr>
        <td align="left"></td>
        <td align="left" colSpan="2">
			     <input type="radio" name="RequiredCertificationReady" value="Y" CHECKED><strong><font face="Arial"></strong> Yes</FONT>
                 <input type="radio" name="RequiredCertificationReady" value="N"><strong><font face="Arial"></strong> No</FONT>

<tr>
<td align="left" wrap>&nbsp;</td>
        <td align="left" width="35"></td>
        <td align="left"><strong><font face="Arial" size="2">(i)&nbsp;&nbsp;If yes, 
            indicate type and date of certification or authorization(e.g. letter 
            of authorization, license, Certificate of Public Convenience &amp; 
            Necessity (CPCN), tarriff, etc.):
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RequiredYesExplanation 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RequiredYesExplanation">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequiredYesExplanation">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="100">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequiredYesExplanation()
{
	RequiredYesExplanation.setStyle(TXT_TEXTBOX);
	RequiredYesExplanation.setDataSource(GetPart1Data);
	RequiredYesExplanation.setDataField('RequiredYesExplanation');
	RequiredYesExplanation.setMaxLength(100);
	RequiredYesExplanation.setColumnCount(50);
}
function _RequiredYesExplanation_ctor()
{
	CreateTextbox('RequiredYesExplanation', _initRequiredYesExplanation, null);
}
</script>
<% RequiredYesExplanation.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 </font></strong>
            

<tr>
<td align="left" wrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        <td align="left">
</td>
        <td align="left"><font face="Arial" size="2"><strong>(ii)&nbsp; If no, 
            explain:</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RequiredNoExplanation 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RequiredNoExplanation">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequiredNoExplanation">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="100">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequiredNoExplanation()
{
	RequiredNoExplanation.setStyle(TXT_TEXTBOX);
	RequiredNoExplanation.setDataSource(GetPart1Data);
	RequiredNoExplanation.setDataField('RequiredNoExplanation');
	RequiredNoExplanation.setMaxLength(100);
	RequiredNoExplanation.setColumnCount(50);
}
function _RequiredNoExplanation_ctor()
{
	CreateTextbox('RequiredNoExplanation', _initRequiredNoExplanation, null);
}
</script>
<% RequiredNoExplanation.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="3">&nbsp; 
    <tr>
        <td align="left" colSpan="3">&nbsp;&nbsp;&nbsp; 
    <tr>
        <td align="left" colSpan="3"><strong><font face="Arial" size="3" color="#993300">1.5&nbsp; Type of Request: 
    
	</font></strong>
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3"><font face="Arial" size="2"><strong>
			<input type="radio" name="TypeOfRequest"   value="A" <%=readOnlyA %> <%=checkedA%>>1)&nbsp; 
            Code Assignment - Requested NXX:</strong></font>
            
<%

Part1NXXAssignLookup.moveFirst
Response.Write"<SELECT id=NXXAssign name=NXXAssign>"
while (Not Part1NXXAssignLookup.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXAssignLookup.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXAssignLookup.moveNext
wend
Response.Write"</SELECT>"

%>
</SELECT>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <p><font face="Arial" size="2"><strong>Secondary NXXs if requested becomes 
            unavailable (optional, you can identify 2 
            NXXs):</strong></font></FONT></p>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
<%

Part1NXXAssignLookup.moveFirst
Response.Write"<SELECT id=NXX2A name=NXX2A>"
while (Not Part1NXXAssignLookup.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXAssignLookup.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXAssignLookup.moveNext
wend
Response.Write"</SELECT>"

Part1NXXAssignLookup.moveFirst
Response.Write"<SELECT id=NXX3A name=NXX3A>"
while (Not Part1NXXAssignLookup.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXAssignLookup.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXAssignLookup.moveNext
wend
Response.Write"</SELECT>"

%>

    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2"><strong>Undesirable NXXs 
            (optional, you can identify 5 NXXs):</strong></font> 
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX1A style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX1A">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1A()
{
	NoNXX1A.setStyle(TXT_TEXTBOX);
	NoNXX1A.setDataSource(GetPart1Data);
	NoNXX1A.setDataField('NoNXX1');
	NoNXX1A.setMaxLength(3);
	NoNXX1A.setColumnCount(3);
}
function _NoNXX1A_ctor()
{
	CreateTextbox('NoNXX1A', _initNoNXX1A, null);
}
</script>
<% NoNXX1A.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX2A style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX2A">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2A()
{
	NoNXX2A.setStyle(TXT_TEXTBOX);
	NoNXX2A.setDataSource(GetPart1Data);
	NoNXX2A.setDataField('NoNXX2');
	NoNXX2A.setMaxLength(3);
	NoNXX2A.setColumnCount(3);
}
function _NoNXX2A_ctor()
{
	CreateTextbox('NoNXX2A', _initNoNXX2A, null);
}
</script>
<% NoNXX2A.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX3A style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX3A">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3A()
{
	NoNXX3A.setStyle(TXT_TEXTBOX);
	NoNXX3A.setDataSource(GetPart1Data);
	NoNXX3A.setDataField('NoNXX3');
	NoNXX3A.setMaxLength(3);
	NoNXX3A.setColumnCount(3);
}
function _NoNXX3A_ctor()
{
	CreateTextbox('NoNXX3A', _initNoNXX3A, null);
}
</script>
<% NoNXX3A.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX4A style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX4A">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4A()
{
	NoNXX4A.setStyle(TXT_TEXTBOX);
	NoNXX4A.setDataSource(GetPart1Data);
	NoNXX4A.setDataField('NoNXX4');
	NoNXX4A.setMaxLength(3);
	NoNXX4A.setColumnCount(3);
}
function _NoNXX4A_ctor()
{
	CreateTextbox('NoNXX4A', _initNoNXX4A, null);
}
</script>
<% NoNXX4A.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX5A style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX5A">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5A()
{
	NoNXX5A.setStyle(TXT_TEXTBOX);
	NoNXX5A.setDataSource(GetPart1Data);
	NoNXX5A.setDataField('NoNXX5');
	NoNXX5A.setMaxLength(3);
	NoNXX5A.setColumnCount(3);
}
function _NoNXX5A_ctor()
{
	CreateTextbox('NoNXX5A', _initNoNXX5A, null);
}
</script>
<% NoNXX5A.display %>

<!--METADATA TYPE="DesignerControl" endspan--> 
        <tr>
        <td align="left" colSpan="2">
        <td align="left"> <font face="arial" size="2">
                 <input type="radio" name="ReasonForRequest" value="aic" <%=readOnlyA %> ><font face="Arial">a) Initial Code for new Switching Entity or new Point of 
            Interconnection (Code Applicant must complete Section 1.8 and Part 
            2)</font><br>
                 <input type="radio" name="ReasonForRequest" value="aau" <%=readOnlyA %> ><font face="Arial">b) Code request for New Application for existing 
            switching entity or point of interconnection (Code Applicant must 
            complete Section 1.7)</font><br>
                 <input type="radio" name="ReasonForRequest" value="aag" <%=readOnlyA %> ><font face="Arial">c) Additional Code for Growth (Code 
            Applicant must complete Section 1.6) </font>
                
      </font>
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3"><strong>
			                 <input type="radio" name="TypeOfRequest" value="U" <%=readOnlyU%> <%=checkedU%>>2)&nbsp; <font face="Arial" size="2" >Update 
            Information (Complete Part 2)&nbsp; <input type="radio" name="ReasonForRequest" value="upd" <%=readOnlyU%> ><font face="Arial"> NXX requiring update:&nbsp;</strong></FONT>
            
<%Part1NXXUpdateLook.moveFirst
Response.Write"<SELECT id=NXXUpdate name=NXXUpdate>"
while (Not Part1NXXUpdateLook.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXUpdateLook.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXUpdateLook.moveNext
wend
Response.Write"</SELECT>"%>

</FONT>
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="3">
            <p><font face="Arial" size="2"><strong><input type="radio" name="TypeOfRequest" value="R" <%=readOnlyR %> <%=checkedR%>>3)&nbsp; Code Reservation 
            Only - Requested NXX:&nbsp;</strong></font></STRONG></FONT>
            
<%Part1NXXReserveLook.moveFirst
Response.Write"<SELECT id=NXXReserve name=NXXReserve>"
while (Not Part1NXXReserveLook.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXReserveLook.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXReserveLook.moveNext
wend
Response.Write"</SELECT>"%>
</p>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <p><font face="Arial" size="2"><strong>Secondary NXXs if requested becomes 
            unavailable (optional, you can identify 2 
            NXXs):</strong></font></FONT></p>
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            
<%Part1NXXReserveLook.moveFirst
Response.Write"<SELECT id=NXX2R name=NXX2R>"
while (Not Part1NXXReserveLook.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXReserveLook.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXReserveLook.moveNext
wend
Response.Write"</SELECT>"%>
            
<% Part1NXXReserveLook.moveFirst
Response.Write"<SELECT id=NXX3R name=NXX3R>"
while (Not Part1NXXReserveLook.EOF) 
Response.Write"<OPTION>"  
Response.Write Part1NXXReserveLook.fields.getValue("NXX") 
Response.Write"</OPTION>"
Part1NXXReserveLook.moveNext
wend
Response.Write"</SELECT>"%>
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2"><strong>Undesirable NXXs 
            (optional, you can identify 5 NXXs):</strong></font> 
    <tr>
        <td align="left" colSpan="2">
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX1R style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX1R">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1R()
{
	NoNXX1R.setStyle(TXT_TEXTBOX);
	NoNXX1R.setDataSource(GetPart1Data);
	NoNXX1R.setDataField('NXX2');
	NoNXX1R.setMaxLength(3);
	NoNXX1R.setColumnCount(3);
}
function _NoNXX1R_ctor()
{
	CreateTextbox('NoNXX1R', _initNoNXX1R, null);
}
</script>
<% NoNXX1R.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX2R style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX2R">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2R()
{
	NoNXX2R.setStyle(TXT_TEXTBOX);
	NoNXX2R.setDataSource(GetPart1Data);
	NoNXX2R.setDataField('NXX2');
	NoNXX2R.setMaxLength(3);
	NoNXX2R.setColumnCount(3);
}
function _NoNXX2R_ctor()
{
	CreateTextbox('NoNXX2R', _initNoNXX2R, null);
}
</script>
<% NoNXX2R.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX3R style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX3R">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3R()
{
	NoNXX3R.setStyle(TXT_TEXTBOX);
	NoNXX3R.setDataSource(GetPart1Data);
	NoNXX3R.setDataField('NXX2');
	NoNXX3R.setMaxLength(3);
	NoNXX3R.setColumnCount(3);
}
function _NoNXX3R_ctor()
{
	CreateTextbox('NoNXX3R', _initNoNXX3R, null);
}
</script>
<% NoNXX3R.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX4R style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX4R">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4R()
{
	NoNXX4R.setStyle(TXT_TEXTBOX);
	NoNXX4R.setDataSource(GetPart1Data);
	NoNXX4R.setDataField('NXX2');
	NoNXX4R.setMaxLength(3);
	NoNXX4R.setColumnCount(3);
}
function _NoNXX4R_ctor()
{
	CreateTextbox('NoNXX4R', _initNoNXX4R, null);
}
</script>
<% NoNXX4R.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=NoNXX5R style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="NoNXX5R">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5R()
{
	NoNXX5R.setStyle(TXT_TEXTBOX);
	NoNXX5R.setDataSource(GetPart1Data);
	NoNXX5R.setDataField('NXX2');
	NoNXX5R.setMaxLength(3);
	NoNXX5R.setColumnCount(3);
}
function _NoNXX5R_ctor()
{
	CreateTextbox('NoNXX5R', _initNoNXX5R, null);
}
</script>
<% NoNXX5R.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="arial" size="2">
                 <input type="radio" name="ReasonForRequest" value="ric" <%=readOnlyR %> ><font face="Arial"> 
            a) Initial Code</font><br>
                 <input type="radio" name="ReasonForRequest" value="rau" <%=readOnlyR %> ><font face="Arial"> b) New Application (Complete Section 
            1.7)</font><br>
                 <input type="radio" name="ReasonForRequest" value="rag" <%=readOnlyR %> ><font face="Arial"> c) Growth (Complete Section 1.6) 
            </font>
                
      </font>
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
        <td align="left" colSpan="3"><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
	<strong>1.6 Additional Code Request For 
            Growth:</strong></font> 
    <tr>
        <td align="left" colSpan="3">&nbsp;
    <tr>
        <td align="left" colSpan="2">
<p>&nbsp;</p>
        <td align="left">
<p><font face="Arial" size="2">Basis of eligibility for an additional code for growth assigned 
            to the switching entity/POI assumes the following: the initial code 
            or the code previously assigned to a new application meets the 
            exhaust criteria, as specified in the Central Office Code (NXX) 
            Assignment Guidelines, depending on whether the NPA is in a 
            non-jeopardy situation as described in Section 7.3 of the 
            guidelines. The appropriate situation shall be indicated below 
            (select one).</font></p>
    <tr>
        <td align="left" colSpan="3">
			
                 <input type="radio" name="NPAinJeopardy" value="n" LANGUAGE=javascript onclick="NPAinJeopardy_n_onclick()"><font face="Arial" size="2"><strong>Non-Jeopardy NPA Situation</strong></font> 
          
    <tr>
        <td align="left" colSpan="2">
        <td align="left"><font face="Arial" size="2">I hereby certify that the existing CO Code(s) 
            (NXX) at this Switching Entity/POI is/(are) projected to exhaust 
            within 12 months of the date of this application. This fact is 
            documented on Appendix B and will be supplied to an auditor when 
            requested to do so per Appendix A of the Guidelines.</font>
    <tr>
        <td align="left" colSpan="3">

		  <input type="radio" name="NPAinJeopardy" value="y" LANGUAGE=javascript onclick="NPAinJeopardy_y_onclick()"><font face="Arial" size="2"><strong>Jeopardy NPA Situation (see Section 7.4(c) of 
            the Guidelines) 
            </strong>
            </font>
        <tr>
        <td align="left" colSpan="2"><font face="Arial"></font>
        <td align="left"><p><font face="Arial" size="2">I 
            hereby certify that the existing CO Code(s) (NXX) at this Switching 
            Entity/POI is/(are) projected to exhaust within 6 months of the date 
            of this application. This fact is documented on Appendix B and will 
            be supplied to an auditor when requested to do so per Appendix A of 
            the Guidelines.</font></p><font face ="" size="2"></font>
    <TR>
        <TD align=left colSpan=3>
    <TR>
        <TD align=left colSpan=3>
    <TR>
        <TD align=left colSpan=3>
    <tr>
        <TD align=left colSpan=3>
    <tr>
        <TD align=left colSpan=3>
<P>&nbsp;<P>
<table border=0 background="">
    <tr>
        <td align="left" colSpan="12"><STRONG><FONT face=Arial size=3 color="#993300">APPENDIX B:</FONT></STRONG>
    <TR>
        <TD align=left colSpan=12>
    <TR>
        <TD align=left colSpan=12><FONT face=Arial size=2><STRONG>NXXs included in growth 
                        calculation:</STRONG></FONT>
                        
<INPUT id=NXXGrowthCal name=NXXGrowthCal size=50 maxlength=100 >
    <TR>
        <TD align=left colSpan=12><STRONG><FONT face=Arial 
            size=2>A.&nbsp; Telephone Numbers (TNs) Available for 
                        Assignment (See Glossary):</FONT></STRONG>&nbsp;
                        
<INPUT id=TNs name=TNs size=9  maxlength=9 value=0 
                       >
    <TR>
        <TD align=left colSpan=12><FONT face=Arial 
            size=2>Definitions of 
                        terms may be found in the Glossary section of the 
                        Central Office Code (NXX) Assignment Guidelines.</FONT>
    <TR>
        <TD align=left colSpan=6></td>
		<td>
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
</td>
		<td>
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
</td>
		<td>
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
</td>
		<td>
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
</td>
		<TD>
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
</td>
		<TD>
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
		</TD></TR>
    <TR>
        <TD align=left colSpan=6><STRONG><FONT face=Arial 
            size=2>B.&nbsp; Previous 6-month growth 
                        history:</FONT></STRONG></TD>
		<td>
                        
<INPUT id=Prev6Month1 name=Prev6Month1 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=Prev6Month2 name=Prev6Month2 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=Prev6Month3 name=Prev6Month3 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=Prev6Month4 name=Prev6Month4 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=Prev6Month5 name=Prev6Month5 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=Prev6Month6 name=Prev6Month6 size=9 maxlength=9 value=0 > 
		</td></TR>
    <TR>
        <TD align=left colSpan=12><FONT face=Arial size=2>Telephone Numbers 
                        (TNs) assigned in each previous month, starting with the 
                        most distant month as Month #1, and Month #6 as the 
                        current month.</FONT></TD></TR>
    <TR>
        <TD align=left colSpan=6><STRONG><FONT face=Arial size=2>C.&nbsp; Projected growth - Months&nbsp;&nbsp; 
                        1-6:</FONT></STRONG></TD>
		<td>
                        
<INPUT id=ProjGrowth16Month1 name=ProjGrowth16Month1 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=ProjGrowth16Month2 name=ProjGrowth16Month2 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=ProjGrowth16Month3 name=ProjGrowth16Month3 maxlength=9 size=9 value=0 >
		</td><td>
                        
<INPUT id=ProjGrowth16Month4 name=ProjGrowth16Month4 size=9 maxlength=9 value=0 >
		</td><TD>
                        
<INPUT id=ProjGrowth16Month5 name=ProjGrowth16Month5 size=9 maxlength=9 value=0 >
		</td><TD>
                        
<INPUT id=ProjGrowth16Month6 name=ProjGrowth16Month6 size=9 maxlength=9 value=0 > 
		</TD></TR>
    <TR>
        <TD align=left colSpan=6>&nbsp;&nbsp;&nbsp;&nbsp; 
            <STRONG><FONT face=Arial size=2>Projected growth - Months&nbsp; 
                        7-12:</FONT></STRONG></TD>
		<td>
                        
<INPUT id=ProjGrowth712Month1 name=ProjGrowth712Month1 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=ProjGrowth712Month2 name=ProjGrowth712Month2 size=9 maxlength=9   
                        value=0>
		</td><td>
                        
<INPUT id=ProjGrowth712Month3 name=ProjGrowth712Month3 size=9 maxlength=9 value=0 >
		</td><td>
                        
<INPUT id=ProjGrowth712Month4 name=ProjGrowth712Month4 size=9 maxlength=9 value=0 >
		</td><TD>
                        
<INPUT id=ProjGrowth712Month5 name=ProjGrowth712Month5 size=9 maxlength=9 value=0 >
		</td><TD>
                        
<INPUT id=ProjGrowth712Month6 name=ProjGrowth712Month6 size=9 maxlength=9 value=0 > 
		</TD></TR>
    <TR>
        <TD align=left colSpan=12><FONT face=Arial size=2>TNs assigned in 
                        each following month, starting with the most recent 
                        month as Month #1.&nbsp; In a jeopardy situation, only 6 
                        months growth projection is required.</FONT></TD></TR>
    <TR>
        <TD align=left colSpan=12><STRONG><FONT face=Arial size=2>D.&nbsp; Average Monthly Growth Rate (From Part C 
                        above):
                        
<INPUT id=AvgMonGrowthRate name=AvgMonGrowthRate readonly size =9 maxlength=9 value=0> 
</FONT></STRONG>
                        

                        
                        
<INPUT type="button" value="Calculate" id=button1 name=button1 LANGUAGE=javascript onclick="return button1_onclick()"> 
                        		</TD>
	</TR>
    <TR>
        <TD align=left colSpan=12><STRONG><FONT face=Arial size=2>E.&nbsp; Months to Exhaust = TNs Available for 
                        Assignment (A) / Average Monthly Growth Rate (D) 
                        =</STRONG>
                        
<INPUT id=MonthsToExhaust name=MonthsToExhaust readonly size =9 maxlength=9 value=0> </FONT></TD></TR>
    <TR>
        <TD align=left colSpan=12><FONT face=Arial size=2>To be assigned an 
                        additional CO Code for growth, &quot;Months to 
                        Exhaust&quot; must be less than or equal to 12 month for 
                        a non -jeopardy NPA (See Section 4.2.1 of the 
                        Guidelines), or less than or equal to 6 months for a 
                        jeopardy NPA (See Section 8.4(c) of the 
                        Guidelines).</FONT></TD></TR>
    <TR>
        <TD align=left colSpan=12><STRONG><FONT face=Arial size=2>Explanation:&nbsp;</FONT></STRONG>
                        
<INPUT id=AppendixBExplanation name=AppendixBExplanation size=75 maxlength=100 
                       > 
                
		</TD>
	</TR>
</table>
<P>&nbsp;<P></P>
		</TD>
	</tr>
    <TR>
        <TD align = left colSpan = 3>
    <TR>
        <TD align = left colSpan = 3>
    <TR>
        <TD align=left colSpan=3>
    <tr>
        <td align = "left" colSpan = "3">
    <tr>
        <td align = "left" colSpan = "3"><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
		<strong>1.7 Code Request for New 
            Application(see Section 4.2 of the Guidelines)</strong></font>
    <tr>
        <td align = "left" colSpan = "3">&nbsp;&nbsp;
    <tr>
        <td align = "left" colSpan = "2">
        <td align = "left"><font face="Arial" size="2">Basis of eligibility for an additional code 
            means that there has not been a code assigned to this switching 
            entity/point of interconnection for this purpose. (Check the 
            applicable space and, if applicable, provide the requested 
            information). If eligibility is based on a category that requires 
            additional explanation or documentation and the code administrator 
            denies a request, the applicant has the option to pursue an appeals 
            process.</font>
    <tr>
        <td align = "left" colSpan = "3">
			 <dd>
    <input type="radio" name="CodeRequestNew" value="c"><strong><font face="Arial" size="2">Code is necessary for distinct 
            routing, rating or billing purposes.<font face="Arial" Size="2"><strong>Any additional 
            information that can be provided by the Code Applicant may 
            facilitate the processing of that 
            application.</strong></font></strong> </FONT></dd>
    <tr>
        <td align = "left" colSpan = "2">
        <td align = "left">
            <strong><font face="Arial" size="2">Description:</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RequestNewNecessary 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RequestNewNecessary">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestNewNecessary">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="100">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestNewNecessary()
{
	RequestNewNecessary.setStyle(TXT_TEXTBOX);
	RequestNewNecessary.setDataSource(GetPart1Data);
	RequestNewNecessary.setDataField('RequestNewNecessary');
	RequestNewNecessary.setMaxLength(100);
	RequestNewNecessary.setColumnCount(50);
}
function _RequestNewNecessary_ctor()
{
	CreateTextbox('RequestNewNecessary', _initRequestNewNecessary, null);
}
</script>
<% RequestNewNecessary.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align = "left" colSpan = "3">
      <dd>
      <input type="radio" name="CodeRequestNew" value="o"><font face="Arial" size="2">
      <strong>Other <font size="2">The Code Applicant must provide an explanation of why existing 
            resources assigned to that entity cannot satisfy this 
            requirement.</strong></font> </FONT></dd>
    <tr>
        <td align = "left" colSpan = "2">
        <td align = "left">
            <font face="Arial" size="2"><strong>Description:</font></STRONG>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=RequestNewOther 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="RequestNewOther">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestNewOther">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="100">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestNewOther()
{
	RequestNewOther.setStyle(TXT_TEXTBOX);
	RequestNewOther.setDataSource(GetPart1Data);
	RequestNewOther.setDataField('RequestNewOther');
	RequestNewOther.setMaxLength(100);
	RequestNewOther.setColumnCount(50);
}
function _RequestNewOther_ctor()
{
	CreateTextbox('RequestNewOther', _initRequestNewOther, null);
}
</script>
<% RequestNewOther.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align = "left" colSpan = "3">
    <TR>
        <TD align=left colSpan=3>
    <tr>
        <td align = "left" colSpan = "3">&nbsp;&nbsp;
    <tr>
        <td align = "left" colSpan = "3"><strong><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
<P>&nbsp;<P>
	1.8 Authorization for entry of Part 2 
            Information into Bellcore databases (Check applicable 
            space):</font></strong></p>
    <tr>
        <td align = "left" colSpan = "3">&nbsp;&nbsp;
    <tr>
        <td align = "left" colSpan = "2">
        <td align = "left"><strong><font face="Arial" size="2"><input type="radio" name="AuthorizationPart2" value="y">Yes-<font size="2"></strong> I have attached a completed Part 2 of this form. 
            This is the Code Administrator's authorization to input/revise the indicated RDBS and/or BRIDS data. Further, I understand that the Code Administrator may not be the authorized party to input the        data. The authorization and/or data input responsibilities are    determined on an Operating Company Number level. If the Code Administrator advises me that said Code Administrator does not have Administrative Operating Company Number (AOCN) responsibility for my data inputs, I will contact Bellcore-TRA to determine the correct AOCN company. Upon that determination, I will submit Part 2 directly to the AOCN company for input to RDBS and BRIDS.</FONT></FONT></STRONG> 
    <tr>
        <td align = "left" colSpan = "2"></td>
        <td align = "left"><input type="radio" name="AuthorizationPart2" value="n"><font face="arial" size="2"><strong>No - <font face="arial" size="2"></strong>Part 2 of this form is not attached. RDBS and BRIDS input will be the responsibility of the Code Applicant. The 66 calendar day nation-wide minimum interval cut-over for RDBS and BRIDS will not begin until input into RDBS and BRIDS has been completed.</font> </FONT><tr>
		<td align = "left" colSpan = "3">&nbsp;&nbsp;&nbsp;</td><tr>
		<td align = "left" colSpan = "3"> <input type="submit" value="Submit" name="submit">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoToMainFrm 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoToMainFrm">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoToMainFrm()
{
	btnGoToMainFrm.value = 'Return';
	btnGoToMainFrm.setStyle(0);
}
function _btnGoToMainFrm_ctor()
{
	CreateButton('btnGoToMainFrm', _initbtnGoToMainFrm, null);
}
</script>
<% btnGoToMainFrm.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<tr>
<td align = "left" colSpan = "3" wrap>
	</td></tr></TBODY></TABLE></FORM>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>
