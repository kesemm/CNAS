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
<title>Processed Table</title>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>
<p align="center"><strong>Processing Timing For ESRD Requests (in working days)</strong></p>
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

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">3.57</td>
<td ALIGN="CENTER">1.80</td>
<td ALIGN="CENTER">21</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">7.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">2.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">3</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">6.50</td>
<td ALIGN="CENTER">2.37</td>
<td ALIGN="CENTER">16</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">2.09</td>
<td ALIGN="CENTER">.30</td>
<td ALIGN="CENTER">11</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">2.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">4</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">3.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">5.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">11.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">30</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">10.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">3</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">9.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">8</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">6.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">4.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">6.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">4</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">6.90</td>
<td ALIGN="CENTER">.32</td>
<td ALIGN="CENTER">10</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">7.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">2.86</td>
<td ALIGN="CENTER">2.27</td>
<td ALIGN="CENTER">7</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">6.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">6.91</td>
<td ALIGN="CENTER">1.56</td>
<td ALIGN="CENTER">23</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">5.60</td>
<td ALIGN="CENTER">.55</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">5.29</td>
<td ALIGN="CENTER">3.17</td>
<td ALIGN="CENTER">35</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">11.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

</table>
<h5>
</body>
</html>
