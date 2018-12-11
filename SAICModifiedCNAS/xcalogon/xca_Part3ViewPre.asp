<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<form action="xca_Part3View.asp" method="post" id="Part3ViewPre" name="Part3ViewPre" onSubmit="return validateForm()">
<!--#include file="xca_CNASLib.inc"-->
<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
       
function validateForm()
        {
            formObj = document.Part3ViewPre;
            
            if (formObj.P3ViewTix.value == "") {
                alert("You have not filled in a Ticket #.  Please enter a number and submit again");
                formObj.P3Tix.focus();               
                return false;
            }
            if (isNaN(formObj.P3ViewTix.value)){ 
                alert("The Ticket is not a number. Please enter a valid Ticket number and submit again");
                formObj.P3Tix.focus();               
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
alert("That Ticket does not exist.  Please try again.....");


// end hiding -->
</script>
<%
session("NoTixSent")=""
end if
%>
</head>
<%
'session("Tix")=Tix
session("Here")="xca_Part3ViewPre.asp"
session("ViewP3") = "ViewP3"

%>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<P><center><font face="Arial Black" color=maroon size=5><strong>View Part 3</strong></font></center></P>
<P>&nbsp;<P>
<font face="Arial" size="3"><font face="Arial"><strong><em>Please enter the 
Part 3 Ticket Number to VIEW.....</em></strong></font></P>&nbsp; 
</font>
<p>&nbsp; 

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="50%">
    <tr>
        <td><font face="Arial">Ticket #:&nbsp;&nbsp;&nbsp; 
            </font>
        <td>

<input id="P3ViewTix" name="P3ViewTix" Size="9" Maxlength="9">
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
