<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<form action="xca_Part3.asp" method="post" id="Part3Pre" name="form44" onSubmit="return validateForm()">
<!--#include file="xca_CNASLib.inc"-->
<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
       
function validateForm()
        {
            formObj = document.Part3Pre;
            
            if (formObj.Tix.value == "") {
                alert("You have not filled in a Ticket #.  Please enter a number and submit again");
                formObj.Tix.focus();               
                return false;
            }
            if (isNaN(formObj.Tix.value)){ 
                alert("The Ticket is not a number. Please enter a valid Ticket number and submit again");
                formObj.Tix.focus();               
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
alert("That Ticket does not exist or is closed.  Please try again.....");


// end hiding -->
</script>
<%
session("NoTixSent")=""
end if

%>
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
session("Tix")=Tix
session("Here")="xca_Part3Pre.asp"
%>
<P><center><font face="Arial Black" color=maroon size=5><strong>Input Part 3</strong></font></center></P>
<P>&nbsp;<P>
<P><font face="Arial" size="3"><font face="Arial"><strong><em>Please enter the 
Part 3 Ticket Number to INPUT.....
<br><br>
You will first see the Part 1 Information. Scroll down to see the Part 3 
form.</em></strong></font></P>&nbsp; 
</font>
<p>&nbsp; 

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="50%">
    <tr>
        <td>
            <DIV align=right><font face="Arial"><STRONG>Ticket #:&nbsp; </STRONG> 
            </font></DIV>
        <td>

<input id="Tix1" name="Tix" Size="9" Maxlength="9">
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
</FORM>

</body>
</html>
