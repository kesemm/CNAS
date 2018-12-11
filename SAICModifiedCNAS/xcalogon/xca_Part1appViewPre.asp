<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<!--#include file="xca_CNASLib.inc"-->

<form action="xca_Part1appView.asp" method="post" id="form444" name="P1ViewPre" onSubmit="return validateForm()">

<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
       
function validateForm()
        {
            formObj = document.P1ViewPre;
            
            if (formObj.P1ViewTix.value == "") {
                alert("You have not filled in a Ticket #.  Please enter a number and submit again");
                formObj.P1ViewTix.focus();               
                return false;
            }
            if (isNaN(formObj.P1ViewTix.value)){ 
                alert("The Ticket is not a number. Please type in a valid Ticket number and submit again");
                formObj.P1ViewTix.focus();               
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
alert("That Ticket does not exist or does not belong to your Entity.  Please try again.....");


// end hiding -->
</script>
<%
session("NoTixSent")=""
end if
%>
</HEAD>
<%
session("Here")= "xca_Part1appViewPre.asp"
session("ViewP1") = "ViewP1"
%>

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<P><center><font face="Arial Black" color=maroon size=5><strong>View Part 1</strong></font></center></P>
<P>&nbsp;<P>
<P><font face="Arial" size="3"><font face="Arial"><strong><em>Please enter the 
Part 1 Ticket Number to VIEW.....</em></strong></font></P>&nbsp; 
</font>
<form action="xca_Part1appView.asp" method="post" id="form444" name="form444">
<p>&nbsp; 

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="50%">
    <tr>
        <td>
            <DIV align=right><font face="Arial"><STRONG>Ticket #:&nbsp; </STRONG> 
            </font></DIV>
        <td>

<input id="P1ViewTix" name="P1ViewTix" Size="9" Maxlength="9">
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
</form>

</body>
</HTML>
