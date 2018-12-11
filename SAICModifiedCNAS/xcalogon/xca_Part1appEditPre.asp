<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<form action="xca_Part1appEdit.asp" method="post" id="form444" name="P1EditPre" onSubmit="return validateForm()">
<!--#include file="xca_CNASLib.inc"-->
<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
       

function validateForm()
        {
            formObj = document.P1EditPre;
            
            if (formObj.P1EditTix.value == "") {
                alert("You have not filled in a Ticket #.  Please enter a number and submit again");
                formObj.P1EditTix.focus();               
                return false;
            }
            if (isNaN(formObj.P1EditTix.value)){ 
                alert("The Ticket is not a number. Please enter a valid Ticket number and submit again");
                formObj.P1EditTix.focus();               
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
alert("That Ticket has been processed, does not exist,is an Update Request, or does not belong to your Entity.  Please try again.....");


// end hiding -->
</script>
<%
session("NoTixSent")=""
end if

%>
</HEAD>
<%
session("Here")= "xca_Part1appEditPre.asp"
session("BlankP1") = "Edit"
%>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<P><center><font face="Arial Black" color=maroon size=5><strong>Edit Part 1</strong></font></center></P>
<P>&nbsp;<P>
<P><font face="Arial" size="3"><font face="Arial"><strong><em>Please enter the Part 1 
Ticket Number to EDIT.....</em></strong></font></P> 
</font>
<p><STRONG>&nbsp;</STRONG> 

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="50%">
    <tr>
        <td>
            <DIV align=right><font face="Arial"><STRONG>Ticket 
            #:&nbsp; </STRONG> 
            </font></DIV>
        <td>

<input id="P1EditTix" name="P1EditTix" Size="9" Maxlength="9">
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
</HTML>
