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
<p align="center"><strong>Processing Timing For Recover CO Codes Requests (in working days)</strong></p>
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
<td ALIGN="CENTER">29</td>
<td ALIGN="CENTER">46</td>
<td ALIGN="CENTER">37.50</td>
<td ALIGN="CENTER">12.02</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">5.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">60</td>
<td ALIGN="CENTER">60</td>
<td ALIGN="CENTER">60.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">4</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">77</td>
<td ALIGN="CENTER">125</td>
<td ALIGN="CENTER">92.00</td>
<td ALIGN="CENTER">19.21</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">56</td>
<td ALIGN="CENTER">145</td>
<td ALIGN="CENTER">97.00</td>
<td ALIGN="CENTER">44.91</td>
<td ALIGN="CENTER">3</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">4.24</td>
<td ALIGN="CENTER">.97</td>
<td ALIGN="CENTER">33</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">112</td>
<td ALIGN="CENTER">45.20</td>
<td ALIGN="CENTER">57.00</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">99</td>
<td ALIGN="CENTER">89.30</td>
<td ALIGN="CENTER">30.67</td>
<td ALIGN="CENTER">10</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">33</td>
<td ALIGN="CENTER">100</td>
<td ALIGN="CENTER">95.33</td>
<td ALIGN="CENTER">15.57</td>
<td ALIGN="CENTER">18</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">105</td>
<td ALIGN="CENTER">19.88</td>
<td ALIGN="CENTER">38.32</td>
<td ALIGN="CENTER">26</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">93</td>
<td ALIGN="CENTER">103</td>
<td ALIGN="CENTER">98.00</td>
<td ALIGN="CENTER">7.07</td>
<td ALIGN="CENTER">2</td>
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
<td ALIGN="CENTER">99</td>
<td ALIGN="CENTER">99</td>
<td ALIGN="CENTER">99.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">86</td>
<td ALIGN="CENTER">91</td>
<td ALIGN="CENTER">87.40</td>
<td ALIGN="CENTER">2.07</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">87</td>
<td ALIGN="CENTER">88</td>
<td ALIGN="CENTER">87.50</td>
<td ALIGN="CENTER">.71</td>
<td ALIGN="CENTER">2</td>
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
<td ALIGN="CENTER">80</td>
<td ALIGN="CENTER">80</td>
<td ALIGN="CENTER">80.00</td>
<td ALIGN="CENTER">.00</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">83</td>
<td ALIGN="CENTER">98</td>
<td ALIGN="CENTER">90.50</td>
<td ALIGN="CENTER">10.61</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">79</td>
<td ALIGN="CENTER">40.00</td>
<td ALIGN="CENTER">55.15</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">46</td>
<td ALIGN="CENTER">47</td>
<td ALIGN="CENTER">46.14</td>
<td ALIGN="CENTER">.38</td>
<td ALIGN="CENTER">7</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">21</td>
<td ALIGN="CENTER">111</td>
<td ALIGN="CENTER">66.00</td>
<td ALIGN="CENTER">63.64</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">106</td>
<td ALIGN="CENTER">53.50</td>
<td ALIGN="CENTER">74.25</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">119</td>
<td ALIGN="CENTER">60.80</td>
<td ALIGN="CENTER">54.26</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">47</td>
<td ALIGN="CENTER">3.88</td>
<td ALIGN="CENTER">11.50</td>
<td ALIGN="CENTER">16</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">364</td>
<td ALIGN="CENTER">237.40</td>
<td ALIGN="CENTER">171.97</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">84</td>
<td ALIGN="CENTER">228</td>
<td ALIGN="CENTER">114.67</td>
<td ALIGN="CENTER">55.67</td>
<td ALIGN="CENTER">6</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">71</td>
<td ALIGN="CENTER">36.00</td>
<td ALIGN="CENTER">49.50</td>
<td ALIGN="CENTER">2</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">94</td>
<td ALIGN="CENTER">94</td>
<td ALIGN="CENTER">94.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
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
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">56</td>
<td ALIGN="CENTER">142</td>
<td ALIGN="CENTER">131.38</td>
<td ALIGN="CENTER">26.83</td>
<td ALIGN="CENTER">13</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">78</td>
<td ALIGN="CENTER">117</td>
<td ALIGN="CENTER">98.82</td>
<td ALIGN="CENTER">19.93</td>
<td ALIGN="CENTER">11</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">84</td>
<td ALIGN="CENTER">84</td>
<td ALIGN="CENTER">84.00</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">1</td>
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
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">n/a</td>
<td ALIGN="CENTER">0</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">71</td>
<td ALIGN="CENTER">118</td>
<td ALIGN="CENTER">106.71</td>
<td ALIGN="CENTER">19.75</td>
<td ALIGN="CENTER">7</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">107</td>
<td ALIGN="CENTER">48.30</td>
<td ALIGN="CENTER">51.00</td>
<td ALIGN="CENTER">10</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">91</td>
<td ALIGN="CENTER">96</td>
<td ALIGN="CENTER">93.50</td>
<td ALIGN="CENTER">3.54</td>
<td ALIGN="CENTER">2</td>
</tr>

</table>
<h5>
</body>
</html>
