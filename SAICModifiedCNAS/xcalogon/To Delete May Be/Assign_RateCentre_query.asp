<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>RateCentre Query</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" -->
</form>
<%
session("selectedNPA") = request.querystring("NPA")
session("selectedNXX") = request.querystring("NXX")
aNPA=session("selectedNPA")
aNXX=session("selectedNXX")
'
' July 14, 2003
'
If aNPA=778 Then
aNPA=604
'Changed by KT on 20041203
ElseIF aNPA="289" Then
aNPA="905"
'Changed by KT on 20041203
Else
SQLNPARateCentre = "SELECT DISTINCT [RC_NAME10], [RC_NAME],[RC_ST], [MAJOR_VERT], [MAJOR_HORZ] FROM [LERG6] Where [NPA]='" & aNPA & "' ORDER BY [RC_NAME10];"
End IF
'
' End July 14, 2003
'
SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstNPARateCentre = server.createObject("ADODB.recordset")
rstNPARateCentre.activeConnection = objConnection
rstNPARateCentre.CursorLocation = adUseServer
rstNPARateCentre.CursorType = adOpenStatic
rstNPARateCentre.open SQLNPARateCentre
'SET rstNPARateCentre = objConnection.execute(SQLNPARateCentre)
%>
<% If aNPA=604 or aNPA=778 Then %>

<b>
<p align="center"><a target="blank" href="NPA604-778Exchanges.asp">Check NPA 604-778 Exchanges</large></a></p>
<p align="center">Assign NPA-Rate Centre query</b><br></p>
<p align="center">Click on the Rate Centre</p>
<br>
</p>

<table align="center" BORDER="1">
  <tr>
    <td>Rate Centres in NPA: <%= aNPA %></td><td>Province</td>
</td>
  </tr>
<%
do until rstNPARateCentre.EOF
%>
  <tr>
    <td align="center"><a HREF="CNAS_Assign_NPA_NXX.asp?NPA=<% =request.querystring("NPA") %>&RC=<%=rstNPARateCentre("RC_NAME10")%>&FRC=<%=rstNPARateCentre("RC_NAME")%>&PR=<%=rstNPARateCentre("RC_ST")%>&MV=<%=rstNPARateCentre("MAJOR_VERT") %> &MH=<%=rstNPARateCentre("MAJOR_HORZ") %> "><%=rstNPARateCentre("RC_NAME") %></a> </td><td><%=rstNPARateCentre("RC_ST") %></td>
     </tr>
<%
  rstNPARateCentre.movenext
loop
%>
<% Response.Write ("<Script Language='JavaScript'>")
Response.write("window.open ('NPA604-778Exchanges.asp')")
Response.Write ("</Script>")
Response.Flush
 %>

</table>

<% ElseIf aNPA=613 Then %>

<b>
<p align="center"><a target="blank" href="http://www.cnac.ca/SCOCAP%20(NPAs%20613&819).htm">Check Assignment Pool </large></a></p>
<p align="center">Assign NPA-Rate Centre query</b><br></p>
<p align="center">Click on the Rate Centre</p>
<br>
</p>

<table align="center" BORDER="1">
  <tr>
    <td>Rate Centres in NPA: <%= aNPA %></td><td>Province</td>
</td>
  </tr>
<%
do until rstNPARateCentre.EOF
%>
  <tr>
    <td align="center"><a HREF="CNAS_Assign_NPA_NXX.asp?NPA=<% =request.querystring("NPA") %>&RC=<%=rstNPARateCentre("RC_NAME10")%>&FRC=<%=rstNPARateCentre("RC_NAME")%>&PR=<%=rstNPARateCentre("RC_ST")%>&MV=<%=rstNPARateCentre("MAJOR_VERT") %> &MH=<%=rstNPARateCentre("MAJOR_HORZ") %> "><%=rstNPARateCentre("RC_NAME") %></a> </td><td><%=rstNPARateCentre("RC_ST") %></td>
     </tr>
<%
  rstNPARateCentre.movenext
loop
%>
<% Response.Write ("<Script Language='JavaScript'>")
Response.write("window.open ('http://www.cnac.ca/SCOCAP%20(NPAs%20613&819).htm')")
Response.Write ("</Script>")
Response.Flush
 %>

</table>

<% ElseIf aNPA=819 Then %>

<b>
<p align="center"><a target="blank" href="http://www.cnac.ca/SCOCAP%20(NPAs%20613&819).htm">Check Assignment Pool </a></p>
<p align="center">Assign NPA-Rate Centre query</b><br></p>
<p align="center">Click on the Rate Centre</p>
<br>
</p>

<table align="center" BORDER="1">
  <tr>
    <td>Rate Centres in NPA: <%= aNPA %></td><td>Province</td>
</td>
  </tr>
<%
do until rstNPARateCentre.EOF
%>
  <tr>
    <td align="center"><a HREF="CNAS_Assign_NPA_NXX.asp?NPA=<% =request.querystring("NPA") %>&RC=<%=rstNPARateCentre("RC_NAME10")%>&FRC=<%=rstNPARateCentre("RC_NAME")%>&PR=<%=rstNPARateCentre("RC_ST")%>&MV=<%=rstNPARateCentre("MAJOR_VERT") %> &MH=<%=rstNPARateCentre("MAJOR_HORZ") %> "><%=rstNPARateCentre("RC_NAME") %></a> </td><td><%=rstNPARateCentre("RC_ST") %></td>
     </tr>
<%
  rstNPARateCentre.movenext
loop
%>
</table>
<% Response.Write ("<Script Language='JavaScript'>")
Response.write("window.open ('http://www.cnac.ca/SCOCAP%20(NPAs%20613&819).htm')")
Response.Write ("</Script>")
Response.Flush
 %>

<% Else %>
<b>

<p align="center">Assign NPA-Rate Centre query</b><br></p>
<p align="center">Click on the Rate Centre</p>
<br>
</p>

<table align="center" BORDER="1">
  <tr>
    <td>Rate Centres in NPA: <%= aNPA %></td><td>Province</td>
</td>
  </tr>
<%
do until rstNPARateCentre.EOF
%>
  <tr>
    <td align="center"><a HREF="CNAS_Assign_NPA_NXX.asp?NPA=<% =request.querystring("NPA") %>&RC=<%=rstNPARateCentre("RC_NAME10")%>&FRC=<%=rstNPARateCentre("RC_NAME")%>&PR=<%=rstNPARateCentre("RC_ST")%>&MV=<%=rstNPARateCentre("MAJOR_VERT") %> &MH=<%=rstNPARateCentre("MAJOR_HORZ") %> "><%=rstNPARateCentre("RC_NAME") %></a> </td><td><%=rstNPARateCentre("RC_ST") %></td>
     </tr>
<%
  rstNPARateCentre.movenext
loop
%>
<% End IF %>
</table>



</body>
</html>
