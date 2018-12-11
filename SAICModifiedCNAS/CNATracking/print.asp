<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Response.Expires = -1
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

  Purpose:  Displays a printer friendly version of the problem details.
  -->

  <!-- 	#include file = "public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = Cint(Session("lhd_sid"))
  
  ' Look for the problem id in the query string, if it's not
  ' there display an error.

  If Len(Request.QueryString("id")) = 0 Then
    Response.Write("<title>Error</title></head><body>")
    Call DisplayError(3, "No valid problem ID was entered.")
  End If
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "Problem")%> <% = Request.QueryString("id") %> <%=lang(cnnDB, "Details")%></title>
    <link rel="stylesheet" type="text/css" href="default.css">
  </head>
  <body>

    <%

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))

      ' Generate a query, making sure to search for id
      Dim queryStr, probRes, rstNotes

      queryStr = _
      "SELECT p.id, t.fname as [Entered_By], r.fname as [Assigned_To], p.start_date, p.status, s.sname, p.due_date," & _
      "p.close_date, c.cname, p.title, p.solution, p.description, " & _
      "pri.pname FROM (((((problems AS p " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN tblUsers AS r ON p.assigned_to = r.sid) " & _
      "INNER JOIN tblUsers As t on p.entered_by_id=t.sid)" & _
      "INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id )" & _
      " WHERE p.id=" & id
      Set probRes = SQLQuery(cnnDB, queryStr)
      Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")

      ' If no results are returned, display an error
      If probRes.EOF Then
        Call DisplayError(3, lang(cnnDB, "ProblemID") & " " & id & " " & lang(cnnDB, "wasfoundinthedatabase") & ".")
      End If

      Dim description, solution
      description = Replace(probRes("description"), vbNewLine, "<br>")
      description = Replace(description, "[", "<b>[")
      description = Replace(description, "]", "]</b>")

      ' If it is a closed problem, get the solution
      If probRes("status") = Cfg(cnnDB, "CloseStatus") Then
        Dim solRes
        Set solRes = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)
        solution = Replace(solRes("solution"), vbNewLine, "<br>")
        solution = Replace(solution, "[", "<b>[")
        solution = Replace(solution, "]", "]</b>")
      End If



      ' Display The problem info, and if OPEN allow some updates
    %>

    <div align="center">
      <table class="Wide">
        <tr class="Head1">
          <td colspan="2">
            <b><%=lang(cnnDB, "DetailsforProblem")%>&nbsp;<% = id %></b>
          </td>
        </tr>
        <tr class="Body1" >
          <td>
            <table class="Wide">
              <tr>
                <td width="45%">
                  <b><%=lang(cnnDB, "ProblemID")%>:</b>
                </td>
                <td>
                  <% = id %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Title")%>:</b>
                </td>
                <td>
                  <% = probRes("title") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "EnteredBy")%>:</b>
                </td>
                <td>
                  <% = probRes("Entered_By") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "StartDate")%>:</b>
                </td>
                <td>
                  <% = FormatDateTime(probRes("start_date"), 1) %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "DueDate")%>:</b>
                </td>
                <td>
                  <% = FormatDateTime(probRes("due_date"), 1) %>
                </td>
              </tr>

      <tr>
        <td>
          <b><%=lang(cnnDB, "CloseDate")%>:</b>
        </td>
        <td>
<% If Len(Trim(probRes("close_date"))) > 0 Then %>
<% =FormatDateTime(probRes("close_date"),1) %>
<% Else %>
<% =DisplayDate(probRes("close_date"),0) %>
<% end if %>
        </td>
      </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "AssignedTo")%>:</b>
                </td>
                <td>
                  <% = probRes("Assigned_To") %>
                </td>
              </tr>
           
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Category")%>:</b>
                </td>
                <td>
                  <% = probRes("cname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Status")%>:</b>
                </td>
                <td>
                  <% = probRes("sname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Priority")%>:</b>
                </td>
                <td>
                  <% = probRes("pname") %>
                </td>
              </tr>
            </table>
          </td>
        <tr class="Head2">
          <td colspan="2">
            <%=lang(cnnDB, "Description")%>:
          </td>
        </tr>
        <tr class="Body1">
          <td colspan="2">
            <% = description %>
          </td>
        </tr>
        <tr class="Head2">
          <td colspan="2">
            <%=lang(cnnDB, "Notes")%>:
          </td>
        </tr>          
        <tr class="Body1">
          <td colspan="2">
            <% If rstNotes.EOF Then %>
               <%=lang(cnnDB, "NoAvailableNotes")%>
             <% 
              Else
              Do While Not rstNotes.EOF
              
                Response.Write("<b>[")
                Response.Write(DisplayDate(rstNotes("addDate"), lhdDateTime) & " - " & rstNotes("uid") & "]")
               Response.Write("</b><br />" & vbNewLine)
                Response.Write(Replace(rstNotes("note"), vbNewLine, "<br />"))
                Response.Write("<p>" & vbNewLine)

                rstNotes.MoveNext
              Loop
              End If %>
          </td>
        </tr>
        <% If probRes("status") = Cfg(cnnDB, "CloseStatus") Then %>
          <tr class="Head2">
            <td colspan="2">
              <%=lang(cnnDB, "Solution")%>:
            </td>
          </tr>
          <tr class="Body1">
            <td colspan="2">
              <% = solution %>
            </td>
          </tr>
        <% SolRes.Close 
           End If %>
      </table>
      <form>
        <input type=button value="<%=lang(cnnDB, "CloseThisWindow")%>" onClick="javascript:window.close();"> 
      </form>
    </div>


    <%
      ' Close Results
      probRes.Close
      cnnDB.Close
    %>
  </body>
</html>
