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
<title>Local Exchange and Routing Guide</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>
<%
SET objConnectionLERG1 = server.createobject("ADODB.connection")
SET rstLergDateLERG1 = server.createobject("ADODB.recordset")
objConnectionLERG1.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLergDateLERG1 = "SELECT * FROM LERG1DATE"
SET rstLergDateLERG1 = objConnectionLERG1.execute(SQLLergDateLERG1)

SET objConnectionLERG6 = server.createobject("ADODB.connection")
SET rstLergDateLERG6 = server.createobject("ADODB.recordset")
objConnectionLERG6.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLergDateLERG6 = "SELECT * FROM LERG6DATE"
SET rstLergDateLERG6 = objConnectionLERG6.execute(SQLLergDateLERG6)

SET objConnectionLERG7 = server.createobject("ADODB.connection")
SET rstLergDateLERG7 = server.createobject("ADODB.recordset")
objConnectionLERG7.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLergDateLERG7 = "SELECT * FROM LERG7DATE"
SET rstLergDateLERG7 = objConnectionLERG7.execute(SQLLergDateLERG7)

SET objConnectionLERG12 = server.createobject("ADODB.connection")
SET rstLergDateLERG12 = server.createobject("ADODB.recordset")
objConnectionLERG12.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SQLLergDateLERG12 = "SELECT * FROM LERG12DATE"
SET rstLergDateLERG12 = objConnectionLERG12.execute(SQLLergDateLERG12)

%>
<b>
<TD><IMG src="undercon.gif" width="35" height="36"></TD>
Underconstruction : queries may not work until further notice (April 27, 2005)
</p>
<p>LERG 1 - Canadian OCN Query </b></b> based on <%=rstlergDateLERG1("LERG1DATE") %> data</p>
<form ACTION="LERG_OCN.asp" METHOD="GET">
  <p>Get a list of Canadian OCNs<br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
</form>
<p><b>LERG 6 - Canadian NXX Query </b></b> based on <%=rstlergDateLERG6("LERG6DATE") %> data</p>
<form ACTION="NPA_NXX_Result.asp" METHOD="get">
  <p>Enter a NPA:<input TYPE="text" NAME="NPA" SIZE="3" MAXLENGTH="3"><br>
  Enter a NXX:<input TYPE="text" NAME="NXX" SIZE="3" MAXLENGTH="3"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
</form>

<form ACTION="LERG_RateCentre_Query.asp" METHOD="GET">
  <p>Get a list of rate centres and count of COCodes for each rate Centre within a NPA<br>
  Enter a NPA:<input TYPE="text" Name="NPA" SIZE="3" MAXLENGHT="3"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
</form>
<p><b>LERG 7 - Canadian Switch Query </b></b> based on <%=rstlergDateLERG7("LERG7DATE") %> data</p>
<form ACTION="LERG_Switch.asp" METHOD="get">
</form>
<table ALIGN="CENTER" BORDER="1" CELLPADING="3" CELLSPACING="3" WIDTH="100%">
<tr>
<td><p>Enter a RateCenter:</td>
<td><input TYPE="text" NAME="RATECENTER" SIZE="4" MAXLENGTH="4"></td>
</tr><br>
  Enter a Province:<input TYPE="text" NAME="PROVINCE" SIZE="2" MAXLENGTH="2"><br>
Enter a Building:<input TYPE="text" NAME="BUILDING" SIZE="2" MAXLENGHT="2"><br>
Enter an Equipment Type:<input TYPE="text" NAME="EQUIPMENT" SIZE="3" MAXLENGHT="3"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
Note: Use % as a wildcard for each character.
  <br>
  </p>
<p><b>LERG 12 - Canadian LRN Query </b></b> based on <%=rstlergDateLERG12("LERG12DATE") %> data</p>
<form ACTION="LERG_LRN.asp" METHOD="get">
  <p>Enter a NPA:<input TYPE="text" NAME="NPA" SIZE="3" MAXLENGTH="3"><br>
  Enter a NXX:<input TYPE="text" NAME="NXX" SIZE="3" MAXLENGTH="3"><br>
  <input TYPE="submit"><input TYPE="reset"><br>
  <br>
  </p>
</form>


</body>
</html>
