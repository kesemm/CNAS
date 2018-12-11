<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  This is file is based on, but has undergone extensive modifications:
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  CVS File:      $RCSfile:$
  Commit Date:   $Date:$ (UTC)
  Committed by:  $Author:$
  CVS Revision:  $Revision:$
  Checkout Tag:  $Name(Version/Build)

  Purpose:  This is the main page.
  -->

  <!-- 	#include file = "public.asp" -->
  <% 
    Dim cnnDB, sid, fname
    Set cnnDB = CreateCon
    sid = Cint(Session("lhd_sid"))
    Session("logon_fname")=Usr(cnnDB, sid, "fname")
  %>

<head>
<title> <%=lang(cnnDB, "HelpDesk")%></title>
<link rel="stylesheet" type="text/css" href="default.css">
</head>
<body>
<%
	Call DisplayHeader(cnnDB, sid)
%>
<div align="center">
  <table Class="Normal">
    <tr class="Head1">
      <td>
        <div align="center">
          <%=lang(cnnDB, "HelpDesk")%>
        </div>
      </td>
    </tr>
<table Class="Normal">
<tr class="Body2">
<td valign="center" align="center">Submit A New Task:
</td>
</tr>

    <tr class="Body1">
    <td align="center">
      <form method="post" action="new.asp">
    <input type="submit" value="<%=lang(cnnDB, "SubmitNewProblem")%>" Class="button"></form>
    </td>
    </tr>
</table>
<table Class="Normal">
<tr class="Body2">
<td valign="center" align="center">Query Tasks Based On User:
</td>
</tr>

<table Class="Normal">
<tr class="Body1">
<td valign="center" align="center">
<form method="post" action="view.asp">
<%=lang(cnnDB, "Viewproblemsfor")%>: <SELECT NAME="rep_id">
<%
' Display a list of users to the rep
' can view their problems.
Dim replRes
Set replRes = SQLQuery(cnnDB, "SELECT * From tblUsers ORDER BY username ASC")
If Not replRes.EOF Then
Do While Not replRes.EOF
If replRes("sid") = sid Then
%>
<OPTION VALUE="<% = replRes("sid")%>" SELECTED>
<% = replRes("username") %></OPTION>
<% Else %>
<OPTION VALUE="<% = replRes("sid")%>">
<% = replRes("username") %></OPTION>
<% End If
replRes.MoveNext
Loop
End If
%>
</SELECT>
&nbsp;&nbsp;
<input type="submit" value="<%=lang(cnnDB, "View")%>"></form>
</td>
</tr>
</table>
<table Class="Normal">
<tr class="Body2">
<td valign="center" align="center">Query Tasks Based On Status:
</td>
</tr>
<tr class="Body1">
<td valign="center" align="center">
<br>
<form method="post" action="viewstatus.asp">
<%=lang(cnnDB, "Viewproblemsfor")%>: <SELECT NAME="status_id">
<%
' Display a list of status that can be viewed.
Dim rstStatus
Set rstStatus = SQLQuery(cnnDB, "SELECT * From status ORDER BY status_id ASC")
If Not rstStatus.EOF Then
Do While Not rstStatus.EOF
%>                 
<OPTION VALUE="<% = rstStatus("status_id")%>">
<% = rstStatus("sname") %></OPTION>
<% 	
rstStatus.MoveNext
Loop
End If
%>
</SELECT>&nbsp;&nbsp;<input type="submit" value="<%=lang(cnnDB, "View")%>"></form>
</td>
</tr>
</table>
<table Class="Normal">
<tr class="Body2">
<td valign="center" align="center">Query Tasks Based On Category Type:
</td>
</tr>
<tr class="Body1">
<td valign="center" align="center">          <br />
<form method="post" action="viewcategory.asp">
<%=lang(cnnDB, "Viewproblemsfor")%>: <SELECT NAME="category_id">
<%
' Display a list of category the rep
' can view.
' Dim replRes
Set replRes = SQLQuery(cnnDB, "SELECT * From categories ORDER BY cname ASC")
If Not replRes.EOF Then
Do While Not replRes.EOF
%>
<OPTION VALUE="<% = replRes("category_id")%>" SELECTED>
<% = replRes("cname") %></OPTION>
<%    replRes.MoveNext
Loop
End If
%>
</SELECT>&nbsp;&nbsp;<input type="submit" value="<%=lang(cnnDB, "View")%>"></form>
</td>
</tr>
<table Class="Normal">
<tr class="Body2">
<td valign="center" align="center">Query Tasks Based On ID:
</td>
</tr>

      <tr class="Body1">
        <td valign="center">
          <div align="center">
            <br />
            <form method="POST" action="details.asp">
              <input type="text" size="6" name="id"> <input type="submit" value="<%=lang(cnnDB, "LookupbyID")%>">
            </form>
          </div>
        </td>
      </tr>
</table>
</div>

<%
	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>
</body>
</html>
