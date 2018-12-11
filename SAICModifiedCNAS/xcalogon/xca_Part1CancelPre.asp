<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<!--#include file="xca_CNASLib.inc"-->
<form action="xca_Part1Cancel.asp" method="post" id="Part1CancelPre" name="Part1CancelPre" onSubmit="return validateForm()">

<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
       
function validateForm()
        {
            formObj = document.Part1CancelPre;
            
            if (formObj.P1CancelTix.value == "") {
                alert("You have not filled in a Ticket #.  Please enter a number and submit again");
                formObj.P1CancelTix.focus();               
                return false;
            }
            if (isNaN(formObj.P1CancelTix.value)){ 
                alert("The Ticket is not a number. Please enter a valid ticket number and submit again");
                formObj.P1CancelTix.focus();               
                return false;   
				}
		}
        // end hiding -->
    

</SCRIPT>


<%
if session("NoTixSent")<>""	then	
%>
<script LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
alert("That Ticket does not exist for your Entity, or cannot be cancelled because it has been processed, or it is an Update Request.  Please enter another value.....");


// end hiding -->
</script>
<%
session("NoTixSent")=""
end if

%>
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
session("Here")="xca_Part1CancelPre.asp"
%>
<P><center><font face="Arial Black" color=maroon size=5><strong>Cancel Part 1</strong></font></center></P>
<P>&nbsp;<P>
<font face="Arial" size="3"><font face="Arial"><strong><em>Please enter the 
Part 1 Ticket Number to CANCEL.....</font>
<br><br></EM></STRONG>

<p>&nbsp; 

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="50%">
    <tr>
        <td>
            <DIV align=right><font face="Arial"><STRONG>Ticket #:&nbsp;<EM> </EM></STRONG> 
            </font></DIV><strong><em>
        <td>

<input id="Tix1" name="P1CancelTix" Size="9" Maxlength="9">
        <td>
<input type="submit" value="Go" id="button1" name="submit">
    <tr>
        <td></td>
        <td></td>
        <td></td></tr>
    <tr>
        <td>
        <td>

        <td></td></tr></table> 
<p>&nbsp;</p>
<hr align="left">
</FORM></EM></STRONG></font>
</body>
</HTML>
