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
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: adduser.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Form to add new users.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "AddNewUsers")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      Call CheckAdmin


      Dim success
      success = FALSE

      If Request.Form("save") = 1 Then
Dim uid, email, fname, firstname, lastname, statusdate, intNewSid
intNewSid = GetUnique(cnnDB, "users")
uid = Left(Lcase(Trim(Request.Form("uid"))), 50)
email = Left(Lcase(Trim(Request.Form("email"))), 50)
email = Replace(email, "'", "''")
firstname = Left(Trim(Replace(Request.Form("firstname"), "'", "''")), 25)
lastname = Left(Trim(Replace(Request.Form("lastname"), "'", "''")), 24)
fname = firstname & " " & lastname
statusdate = SQLDate(Now, lhdAddSQLDelim)



        If CBool(InStr(uid, "'")) Then
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
        End If
        If Len(email) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Emailaddress") & " " & lang(cnnDB, "isarequiredfield") & ".")
        End If
        If Len(firstname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "FirstName") & " " & lang(cnnDB, "isarequiredfield") & ".")
        End IF
        If Len(lastname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "LastName") & " " & lang(cnnDB, "isarequiredfield") & ".")
        End IF

        
        Dim sqlString, updRes

        Dim newpassword
newpassword = Left(Trim(Request.Form("newpassword")), 50)
newpassword = Replace(newpassword, "'", "''")
sqlString = "INSERT INTO tblUsers (sid, uid, email, fname, firstname, lastname, statusdate,[password])" & _
            " VALUES (" & intNewSid & ", '" & uid & "', '" & email & "', '" & fname & "', '" & _
          firstname & "', '" & lastname & "', " & statusdate & ", '" & newpassword & "')"

          Set updRes = SQLQuery(cnnDB, sqlString)
          success = True
      End If
    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "AddNewUsers")%>
          </td>
        </tr>
        <% If success Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "AccountCreated")%>: '<% = uid %>'
              </div>
            </td>
          </tr>
        <% End If %>
        <tr class="Body1">
          <td>
            <form name="upduser" action="adduser.asp" method="POST">
              <input type="hidden" name="save" value="1">
              <p>
              <table class="Normal">
              <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "Username")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="uid" size="30">
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "FirstName")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="firstname" size="30">
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "LastName")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="lastname" size="30">
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "EmailAddress")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="email" size="30">
                  </td>
                </tr>

                <tr class="Head2">
                  <td colspan="2">
                    <%=lang(cnnDB, "Password")%>:
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "NewPassword")%>: </b>
                  </td>
                  <td>
                    <input type="password" name="newpassword" size="30">
                  </td>
                </tr>
              </table>
              <p>
              <div align="center">
                <p>
                <input type="submit" value="<%=lang(cnnDB, "Save")%>">
              </div>
            </form>
          </td>
        </tr>
      </table>
      <p>
      <a href="viewusers.asp"><%=lang(cnnDB, "ManageUsers")%></a><br>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
    </div>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
