<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#include file="xca_CNASLib.inc"-->
<HTML>
<HEAD>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>CNAS Assign Database Query</title>

<%
UserEntityType=session("UserEntityType")
 %>
</head>
<body text="black" bgProperties="fixed" bgColor="#d7c7a4">
<p><%
aNPA = session("selectedNPA")
session("selectedNPA")=session("selectedNPA")
aNXX = session("selectedNXX")
session("selectedNXX")=session("selectedNXX")
aRC = request.querystring("RC")
aFRC = request.querystring("FRC")
aPR = request.querystring("PR")
session("selectedRC")=aRC
aMV = request.querystring("MV")
aMH = request.querystring("MH")
aNPANXXQry=""
aLERGQry=""
aTitle=""
aHeader=""
SET objConnection = server.createobject("ADODB.connection")
SET objConnectionLERG = server.createobject("ADODB.connection")
SET rstNPANXXQry =server.createobject("ADODB.recordset")
SET rstLERGQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
objConnectionLERG.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
If aNPA=204 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=204 or NPA=306 or NPA=807 or NPA=867)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='204' or NPA='306' or NPA='807' or NPA='867')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=250 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=250 or NPA=403 or NPA=604 or NPA=778 or NPA=780 or NPA=867)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='250' or NPA='403' or NPA='604' or NPA='778' or NPA='780' or NPA='867')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=289 or aNPA= 905 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=289 or NPA=416 or NPA=519 or NPA=613 or NPA=647 or NPA=705 or NPA=905)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='289' or NPA='416' or NPA='519' or NPA='613' or NPA='647' or NPA='705' or NPA='905')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=306 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=306 or NPA=204 or NPA=403 or NPA=780 or NPA=867)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='306' or NPA='204' or NPA='403' or NPA='780' or NPA='867')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=403 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=403 or NPA=250 or NPA=306 or NPA=780)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='403' or NPA='250' or NPA='306' or NPA='780')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=416 or aNPA=647 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=416 or NPA=289 or NPA=647 or NPA=905)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='416' or NPA='289' or NPA='647' or NPA='905')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=418 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=418 or NPA=506 or NPA=709 or NPA=819)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='418' or NPA='506' or NPA='709' or NPA='819')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=450 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=450 or NPA=514 or NPA=613 or NPA=819 or NPA=418)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='450' or NPA='514' or NPA='613' or NPA='819' or NPA='418')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=506 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=506 or NPA=418 or NPA=902)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='506' or NPA='418' or NPA='902')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=514 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=514 or NPA=450 or NPA=418 or NPA=819)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='514' or NPA='450' or NPA='418' or NPA='819')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=519 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=519 or NPA=289 or NPA=705 or NPA=905)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='519' or NPA='289' or NPA='705' or NPA='905')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=604 or aNPA=778 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=604 or NPA=250 or NPA=778)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='604' or NPA='250' or NPA='778')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=613 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=613 or NPA=289 or NPA=450 or NPA=705 or NPA=819 or NPA=905)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='613' or NPA='289' or NPA='450' or NPA='705' or NPA='819' or NPA='905')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=705 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=705 or NPA=519 or NPA=289 or NPA=613 or NPA=807 or NPA=819 or NPA=905)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='705' or NPA='289' or NPA='519' or NPA='613' or NPA='807' or NPA='819' or NPA='905')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=709 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=709 or NPA=418 or NPA=819)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='709' or NPA='418' or NPA='819')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=780 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=780 or NPA=250 or NPA=306 or NPA=403 or NPA=867)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='780' or NPA='250' or NPA='306' or NPA='403' or NPA='867')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=807 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=807 or NPA=204 or NPA=705)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='807' or NPA='204' or NPA='705')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=819 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=819 or NPA=418 or NPA=450 or NPA=613 or NPA=705 or NPA=709)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='819' or NPA='418' or NPA='450' or NPA='613' or NPA='705' or NPA='709')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=867 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=867 or NPA=204 or NPA=250 or NPA=306 or NPA=780)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='867' or NPA='204' or NPA='250' or NPA='306' or NPA='780')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
Elseif aNPA=902 Then
SQLNPANXXQry = "SELECT NPA,NXX,COStatusDescription,xca_COCode.OCN,EntityName,RateCenter,PublicRemarks,CNARemarks From xca_COCode Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus Left Join xca_Entity On xca_COCode.EntityID=xca_Entity.EntityID Where (NPA=902 or NPA=506)  AND NXX='" & aNXX & "' Order by NPA,NXX;"
SQLLERGQry="SELECT Distinct NPA,NXX,OCN,[RCFULLNAME],[RCSTATE], [MAJOR_V],[MAJOR_H] FROM [LERG6] INNER JOIN [LERG8] ON [LERG6].[RCABBRE]+[LERG6].[LOCSTATE]=[LERG8].[RCABBRNAME]+[LERG8].[RCSTATE] Where  (NPA='902' or NPA='506')  AND NXX='" & aNXX & "' Order by NPA,NXX;"
End if 
SET rstNPANXXQry = objConnection.execute(SQLNPANXXQry)
SET rstLERGQry=objConnectionLERG.execute(SQLLERGQry)
%> </p>
<small>
<p align="center"><strong>CNAS Assign NPA-NXX Query </strong></p>
<b>
<p><br>
<b></p>

<p align="center">Listing of Neighbouring NXXs around NPA <%=aNPA%> and NXX <%=aNXX%> for the RateCentre of <%=aRC%> </b></p>
<p align="center">From CNA Records</b></p>
</small>
<table align="center" BORDER="1">
  <tr>
    <th align="center"><small>NPA</small></th>
    <th align="center"><small>NXX</small></th>
    <th align="center"><small>Status</small></th>
    <th align="center"><small>Company</small></th>
    <th align="center"><small>OCN</small></th>
    <th align="center"><small>Rate Centre</small></th>
    <th align="center"><small>Public Remarks</small></th>
    <th align="center"><small>CNA Remarks</small></th>
    
<% Do Until rstNPANXXQry.EOF %> 
  </tr>
  <tr align="center">
    <td><small><%= rstNPANXXQry("NPA") %></small>
</td>
    <td><small><%= rstNPANXXQry("NXX") %></small>
</td>
    <td nowrap><small><%= rstNPANXXQry("COStatusDescription") %></small>
</td>
    <td nowrap><small><%= rstNPANXXQry("EntityName") %></small>
</td>
    <td><small><%= rstNPANXXQry("OCN") %></small>
</td>
    <td nowrap><small><%= rstNPANXXQry("RateCenter") %></small>
</td>
    <td><small><%= rstNPANXXQry("PublicRemarks") %></small>
</td>
    <td><small><%= rstNPANXXQry("CNARemarks") %></small>
</td>
  </tr>
<%
If IsNull(rstNPANXXQry("COStatusDescription")) Then
aCOStatusDescription="-"
Else
aCOStatusDescription=rstNPANXXQry("COStatusDescription")
End IF
If IsNull(rstNPANXXQry("EntityName")) Then
aEntityName="-"
Else
aEntityName=rstNPANXXQry("EntityName")
End IF
If IsNull(rstNPANXXQry("OCN")) Then
aOCN="-"
Else
aOCN=Trim(rstNPANXXQry("OCN"))
End IF
If IsNull(rstNPANXXQry("RateCenter")) Then
aRateCenter="-"
Else
aRateCenter=rstNPANXXQry("RateCenter")
End IF
If IsNull(rstNPANXXQry("PublicRemarks")) Then
aPublicRemarks="-"
Else
aPublicRemarks=Replace(rstNPANXXQry("PublicRemarks"),",",";")
End IF
If IsNull(rstNPANXXQry("CNARemarks")) Then
aCNARemarks="-"
Else
aCNARemarks=Replace(rstNPANXXQry("CNARemarks"),",",";")
End IF
%>
<% aNPANXXQry=aNPANXXQry+Trim(rstNPANXXQry("NPA"))+chr(44)+ Trim(rstNPANXXQry("NXX"))+chr(44)+ Trim(aCOStatusDescription) + chr(44) + Trim(aEntityName) + chr(44)+ Trim(aOCN) + chr(44)+ Trim(aRateCenter) + chr(44)+ Trim(aPublicRemarks) + chr(44)+ Trim(aCNARemarks)+ chr(44)+chr(13)%>
<% rstNPANXXQry.moveNext
 loop %>
</table>
<%objConnection.close %>
</b></b>
<br>
<small>
<p align="center">From the LERG</b></p>
<p align="center">The requested Full RateCentre is <%=aFRC%> (<%=aRC%>) with Major V of <%=aMV%> and Major H of <%=aMH%> </b></p>
<table align="center" BORDER="1">
  <tr>
    <th align="center"><small>NPA</small></th>
    <th align="center"><small>NXX</small></th>
    <th align="center"><small>OCN</small></th>
    <th align="center"><small>Rate Centre</small></th>
    <th align="center"><small>MAJOR V</small></th>
    <th align="center"><small>MAJOR H</small></th>
    <th align="center"><small>Distance (km)</small></th>
    <th align="center"><small>Warning</small></th>
    
<% Do Until rstLERGQry.EOF %>   
  </tr>
<%
aDistance=round(1.61*sqr((aMV-rstLERGQry("MAJOR_V"))*(aMV-rstLERGQry("MAJOR_V"))+(aMH-rstLERGQry("MAJOR_H"))*(aMH-rstLERGQry("MAJOR_H")))/3,1)
If aDistance < 60 Then
aWarning="High"
Elseif aDistance <100 Then
aWarning="Medium"
Else
aWarning="Low"
End If
%>
  <tr align="center">
    <td nowrap><small><%= rstLERGQry("NPA") %></small>
</td>
    <td nowrap><small><%= rstLERGQry("NXX") %></small>
</td>
    <td nowrap><small><%= rstLERGQry("OCN") %></small>
</td>
    <td nowrap><small><%= rstLERGQry("RCFULLNAME") %>, <%= rstLERGQry("RCSTATE") %></small>
</td>
    <td nowrap><small><%= rstLERGQry("MAJOR_V") %></small>
</td>
    <td nowrap><small><%= rstLERGQry("MAJOR_H") %></small>
</td>
    <td nowrap><small><%=aDistance%></small>
</td>
    <td nowrap><small><%=aWarning%></small>
</td>

  </tr>
<% aLERGQry=aLERGQry+Trim(rstLERGQry("NPA"))+chr(44)+ Trim(rstLERGQry("NXX"))+chr(44)+ Trim(rstLERGQry("OCN"))+chr(44)+ Trim(rstLERGQry("RCFULLNAME"))+  chr(44) + Trim(rstLERGQry("RCSTATE")) + chr(44)+  Trim(rstLERGQry("MAJOR_V")) + chr(44)+ Trim(rstLERGQry("MAJOR_H")) + chr(44)+  CSTR(aDistance) + chr(44)+ aWarning + chr(44)+ chr(13)%>
<% rstLERGQry.moveNext
 loop %>
</table>
<INPUT TYPE="BUTTON" VALUE="Print Page" onClick="window.print()">
<%PrintDate = "Produced on "+Cstr(WeekdayName(Weekday(Date()),False))+" "+Cstr((MonthName(month(date()),False)))+"-"+Cstr(day(date()))+"-" +Cstr(year(date()))+" at " +Cstr((FormatDateTime(now(),vblongtime)))%>
<% = PrintDate %>

<%objConnectionLERG.close %>

</BODY>
</FORM>
</HTML>
