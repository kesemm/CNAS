<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>NPA 403 JCP MAX IN SERVICE DATES</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPA403MaxInservice_Result.asp,v $
'* Commit Date:   $Date: 2006/11/06 16:25:27 $ (UTC)
'* Committed by:  $Author: browng $
'* CVS Revision:  $Revision: 1.1 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>NPA 403 JCP MAX IN SERVICE DATES</BIG></B><BR></P>
<BR>
<%
' SET UP THE CONNECTION AND RECORDSET

SQLQueryText ="select xca_COCode.Tix,xca_COCode.NPA,NXX,EntityName,xca_COCode.RateCenter,Convert(varchar(20),ApplicationDate,101) As ApplicationDate, Convert(varchar(20),DateAdd(mm,6,ApplicationDate),101) As MaxInserviceDate,IsNull(CNARemarks,'') As CNARemarks " & _
"from xca_COCode " & _
"Left Join xca_Entity " & _
"On xca_COCode.EntityID=xca_Entity.EntityID " & _
"Left Join xca_Part1 " & _
"On xca_COCode.Tix=xca_Part1.Tix " & _
"where xca_COCode.NPA=403 And Status='A' And ApplicationDate <= '01-11-2006' " & _
"Union " & _
"select xca_COCode.Tix,xca_COCode.NPA,NXX,EntityName,xca_COCode.RateCenter,Convert(varchar(20),ApplicationDate,101) As ApplicationDate, Convert(varchar(20),DateAdd(mm,4,ApplicationDate),101) As MaxInserviceDate,IsNull(CNARemarks,'') As CNARemarks " & _
"from xca_COCode " & _
"Left Join xca_Entity " & _
"On xca_COCode.EntityID=xca_Entity.EntityID " & _
"Left Join xca_Part1 " & _
"On xca_COCode.Tix=xca_Part1.Tix " & _
"where xca_COCode.NPA=403 And Status='A' And ApplicationDate > '01-11-2006' " & _
"Order by xca_COCode.Tix "
SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstNPA403Codes = server.createObject("ADODB.recordset")
rstNPA403Codes.activeConnection = objConnection
rstNPA403Codes.CursorLocation = adUseServer
rstNPA403Codes.CursorType = adOpenStatic
rstNPA403Codes.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS TO DETERMINE IF NPA ENTERED IS VALID

if rstNPA403Codes.RecordCount>0 then

%>
<TABLE align="center" border="1" width="97%">
<TR>
<TD align="center"><B>TIX</B></TD>
<TD align="center"><B>NPA</B></TD>
<TD align="center"><B>NXX</B></TD>
<TD align="center"><B>Company</B></TD>
<TD align="center"><B>RateCenter</B></TD>
<TD align="center"><B>Application Date<BR>(mm/dd/yyyy)</B></TD>
<TD align="center"><B>Max In Service Date<BR>(mm/dd/yyyy)</B></TD>
<TD align="center"><B>CNA Remarks</B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstNPA403Codes.EOF
%>
<TR>
<TD align="center"><%=rstNPA403Codes("TIX")%></TD>
<TD align="center"><%=rstNPA403Codes("NPA")%></TD>
<TD align="center"><%=rstNPA403Codes("NXX")%>&nbsp;</TD>
<TD align="center"><%=rstNPA403Codes("EntityName")%>&nbsp;</TD>
<TD align="center"><%=rstNPA403Codes("RateCenter")%>&nbsp;</TD>
<TD align="center"><%=rstNPA403Codes("ApplicationDate")%>&nbsp;</TD>
<TD align="center"><%=rstNPA403Codes("MaxInserviceDate")%>&nbsp;</TD>
<TD align="center"><%=rstNPA403Codes("CNARemarks")%>&nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstNPA403Codes.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>No Special Codes in NPA: or the NPA is Invalid</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstNPA403Codes.Close
objConnection.close
%><BR>
<%
' THIS IS THE VERSION CONTROL INFORMATION BLOCK
' ---------------------------------------------
'
' Subdued input text box, that when clicked will make an alert with CVS Info
'
%>
<br><br>
<INPUT TYPE="TEXT" 
       STYLE="border: none; background-color: #D7C7A4; font: 7pt Arial; color: gray; width: 200px" 
       ONCLICK="VerInfo()" VALUE="CNAS Version Control Information"
       READONLY>
<SCRIPT language="JavaScript">
function VerInfo()
{
var strAlertText
strAlertText="SAIC Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: NPA403MaxInservice_Result.asp,v $\n"
+"$Revision: 1.1 $\n"
+"$Date: 2006/11/06 16:25:27 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>
