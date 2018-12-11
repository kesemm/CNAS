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
  Purpose:  Logs off the user by removing session variables.

  -->
  <!-- 	#include file = "public.asp" -->
  
<%
      Session("lhd_LanguageID") = Empty
      Session("lhd_IsAdmin") = False
      Session("lhd_sid") = 0

Response.Write ("<Script Language='JavaScript'>")
Response.write("window.close()")
Response.Write ("</Script>")
Response.Flush
 %>
  </body>
</html>
