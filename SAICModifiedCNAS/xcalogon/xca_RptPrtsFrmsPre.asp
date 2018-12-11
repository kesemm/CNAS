<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<!--#include file="xca_CNASLib.inc"-->

<form action="xca_RptPrtsFrms.asp" method="post" id="RptPrtsFrms" name="RptPrtsFrms" onSubmit="return validateForm()">

<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
       
function validateForm()
        {
            formObj = document.RptPrtsFrms;
            
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
alert("That Ticket does not exist.  Please try again.....");


// end hiding -->
</script>
<%
session("NoTixSent")=""
end if
%>
</head>
<%
session("Tix")=Tix
'session("abTix")=Tix
session("Here")="xca_RptPrtsFrmsPre.asp"
session("ViewP134") = "ViewP134"
%>

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">

<table align=center border="0" cellpadding="2">
<tr>
  <td align="center" nowrap><font color="maroon" face="Arial Black" size="5"><strong>
Request Forms Report: Part1, Part3, Part4</strong></font></td>
</tr>
</table>

<p>&nbsp;</p>

<table WIDTH="36.76%" BGCOLOR="#d7c7a4" BORDERCOLOR="maroon" ALIGN="center" BORDER="0" CELLSPACING="1" CELLPADDING="2" height="17" style="HEIGHT: 17px; WIDTH: 311px">
	<tr>
		<td BORDERCOLOR="maroon" NOWRAP align="right" vAlign="top"><strong><em><font color="black" face="Arial" size="4">TICKET#:</font></em></strong></td>
		<td BGCOLOR="#d7c7a4" BORDERCOLOR="maroon" NOWRAP align="left" vAlign="top">

<input id="P134Tix" name="Tix" Size="9" Maxlength="9">
           
<font color="black" face="Arial" size="2"><strong>(<em>Please enter the 
            Ticket Number to the corresponding Part 1, 3, &amp; 4 Forms 
            </em>)</strong> </font> 
</td>
	</tr>
    <tr>
    <td></td>
        <td noWrap align=left>
            
<input type="submit" value="GO" id="button1" name="submit"> 
            

</td>
       </tr>
       </table>
<br><br>
<br><br>
</form>
<P>&nbsp;</P>

</body>
</HTML>
