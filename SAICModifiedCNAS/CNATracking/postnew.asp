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

  Purpose:  Takes the input from new.asp, checks for errors and enters the
  problem into the database.
  -->

  <!-- 	#include file = "public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = Cint(Session("lhd_sid"))
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemSubmitted")%></title>
    <link rel="stylesheet" type="text/css" href="default.css">
  </head>
  <body>

    <%
      ' Get the information from the form fields
      Dim entered_by_id, category, title, description, assignedto_id, duedate, priority,quantity,startdate
      Dim kb
      
      entered_by_id = Request.Form("entered_by_id")
      category = Cint(Request.Form("category"))
      priority = Cint(Request.Form("priority"))
      title = Request.Form("title")
      description = Request.Form("description")
      assignedto_id = Cint(Request.Form("assignedto"))
      quantity = Cint(Request.Form("quantity"))
	startdate = SQLDate(Request.Form("startdate"),lhdAddSQLDelim)
	duedate = SQLDate(Request.Form("duedate"),lhdAddSQLDelim)

    ' Check for required fields
    if IsDate(Request.Form("duedate")) = False Then
    Call DisplayError(3, lang(cnnDB, "InvalidDate"))
    end if

    if IsDate(Request.Form("startdate")) = False Then
    Call DisplayError(3, lang(cnnDB, "InvalidDate"))
    end if

      if category = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Category"))
      End if

      if priority = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Priority"))
      End if

      if assignedto_id = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "AssignedTo"))
      End if

      if Len(title)=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Title"))
      Elseif Len(title) > 50 Then
        title = Trim(title)
        title = Left(title, 50)
      End if

      if Len(description)=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Description"))
      End if

    ' Get missing variables to enter problem
      Dim id, status, rep, time_spent
      status = Cfg(cnnDB, "DefaultStatus")
      time_spent = 10
'
' I cannot get the DateTime formatting to work with the FormatDateTime
' on the SQLDate format so I have created a display start date
'
      ' Get the category name by querying on category_id
      Dim cname, catRes
      Set catRes = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & Request.Form("category"))
 '     rep = catRes("rep_id")
      cname = catRes("cname")

      ' Get the category name by querying on category_id
      Dim fname, assRes
      Set assRes = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE sid=" & Request.Form("assignedto"))
      fname = assRes("fname")
	


    ' Get the problem ID number then immediately update it
      id = GetUnique(cnnDB, "problems")

    ' Clean up variables
      title = Replace(title,"'","''")
      description = Replace(description,"'","''")

    ' All data is present
    ' Write problem into database
      Dim strProblemQry, rstProbInsert
      strProblemQry = "INSERT INTO problems (id, entered_by_id,assigned_to,category, title, description, priority," & _
	"status, start_date, time_spent, due_date,quantity) " & _
      "VALUES (" & id & "," & entered_by_id & "," & assignedto_id & "," & category & ",'" & title & "','" & description & "'," & priority & "," & status & "," & startdate & "," & time_spent &", " & duedate & ", " & quantity &")"
      Set rstProbInsert = SQLQuery(cnnDB, strProblemQry)

    Call eMessage(cnnDB, "repnew", id, Usr(cnnDB, assignedto_id, "email"),Usr(cnnDB,sid,"email"))

    ' Convert the strings back to display them
      title = Replace(title,"''","'")
      description = Replace(description,"''","'")
    %>

    <div align="center">
      <table class="Wide">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "Problem")%>&nbsp;<% = id %>&nbsp;<%=lang(cnnDB, "Submitted")%>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <table class="Wide">
              <tr>
                <td width="125">
                  <b><%=lang(cnnDB, "ProblemID")%>:</b>
                </td>
                <td>
                  <% = id %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "EnteredBy")%>:</b>
                </td>
                <td>
                  <% = Usr(cnnDB,entered_by_id,"fname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "StartDate")%>:</b>
                </td>
                <td>
          <% = FormatDateTime(Request.Form("startdate"),1) %>
                </td>
              </tr>
              <tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "DueDate")%>:</b>
                </td>
                <td>
          <% = FormatDateTime(Request.Form("duedate"),1) %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Category")%>:</b>
                </td>
                <td>
                  <% = catRes("cname") %>
                </td>
              </tr>
<tr>
        <td>
          <b><%=lang(cnnDB, "Quantity")%>:</b>
        </td>
        <td>
          <input type="text" size="4" name="quantity" value="<% = quantity %>")
        </td>
      </tr>

              <tr>
                <td>
                  <b><%=lang(cnnDB, "AssignedTo")%>:</b>
                </td>
                <td>
                  <% = fname %></a>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Title")%>:</b>
                </td>
                <td>
                  <% = title %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <%=lang(cnnDB, "Description")%>:
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <center><form><textarea name="display_desc" rows="10" cols="80"><% = description %></textarea></form></center>
          </td>
        </tr>
      </table>
    </div>

    <%
      ' Close records
      catRes.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>

  </body>
</html>
