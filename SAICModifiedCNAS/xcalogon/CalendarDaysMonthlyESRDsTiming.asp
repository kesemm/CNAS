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
<title>Processing Timing For ESRD Requests (in Calendar Days)</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<%
' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT
sqlQry = "SELECT [Month], [Year], IsNull(CAST([Min] as VarChar),'&nbsp;') as [Min], IsNull(CAST([Max] as VarChar),'&nbsp;') as [Max], IsNull(CAST(Round([Avg],2) as VarChar),'&nbsp;') as [Avg], IsNull(CAST(Round([STDEV],2) as VarChar),'&nbsp;') As [STDEV], [Sum] from CalendarDaysMonthlyESRDs Where [Year]=2018 order by [Year] desc, [Order]"
 %>



<%
SET objConnection = server.createobject("ADODB.connection")
SET rstQry = server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstQry = objConnection.execute(sqlQry)
%>

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>
<p align="center"><strong>Processing Timing For ESRD Requests (in Calendar Days)</strong></p>
<p>  <br>
</p>
<p>  <br>
</p>
<table ALIGN="CENTER" BORDER="1" CELLPADING="3" CELLSPACING="3" WIDTH="100%">
<tr ALIGN="CENTER">
<th ALIGN="CENTER">Month</th>
<th ALIGN="CENTER">Year</th>
<th ALIGN="CENTER">Min</th>
<th ALIGN="CENTER">Max</th>
<th ALIGN="CENTER">Avg</th>
<th ALIGN="CENTER">STDEV</th>
<th ALIGN="CENTER">Monthly Total</th>
</tr>
<tr>
</tr>



<% Do Until rstQry.EOF %>
<tr>
<td ALIGN="CENTER"><%=rstQry("Month")%></td>
<td ALIGN="CENTER"><%=rstQry("Year")%></td>
<td ALIGN="CENTER"><%=rstQry("Min")%></td>
<td ALIGN="CENTER"><%=rstQry("Max")%></td>
<td ALIGN="CENTER"><%=rstQry("Avg")%></td>
<td ALIGN="CENTER"><%=rstQry("STDEV")%></td>
<td ALIGN="CENTER"><%=rstQry("Sum")%></td>
</tr>
<% rstQry.moveNext
loop %>



</table>
<h5>
</body>
</html>
