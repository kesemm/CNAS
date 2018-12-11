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
<p align="center"><strong>Processing Timing For Update CO Codes Requests (in working days)</strong></p>
<p align="center">(includes Misc and Bulk Codes) </p>
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
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">4.02</td>
<td ALIGN="CENTER">2.50</td>
<td ALIGN="CENTER">49</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.49</td>
<td ALIGN="CENTER">2.20</td>
<td ALIGN="CENTER">68</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.92</td>
<td ALIGN="CENTER">2.77</td>
<td ALIGN="CENTER">53</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">4.21</td>
<td ALIGN="CENTER">2.42</td>
<td ALIGN="CENTER">14</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">4.96</td>
<td ALIGN="CENTER">1.77</td>
<td ALIGN="CENTER">23</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.47</td>
<td ALIGN="CENTER">3.75</td>
<td ALIGN="CENTER">36</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">3.05</td>
<td ALIGN="CENTER">2.06</td>
<td ALIGN="CENTER">20</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">4.27</td>
<td ALIGN="CENTER">2.64</td>
<td ALIGN="CENTER">33</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">6.69</td>
<td ALIGN="CENTER">2.09</td>
<td ALIGN="CENTER">16</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.30</td>
<td ALIGN="CENTER">3.16</td>
<td ALIGN="CENTER">10</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">3.61</td>
<td ALIGN="CENTER">3.03</td>
<td ALIGN="CENTER">23</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">5.33</td>
<td ALIGN="CENTER">1.44</td>
<td ALIGN="CENTER">12</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">3.29</td>
<td ALIGN="CENTER">2.71</td>
<td ALIGN="CENTER">17</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.47</td>
<td ALIGN="CENTER">2.80</td>
<td ALIGN="CENTER">153</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.34</td>
<td ALIGN="CENTER">1.73</td>
<td ALIGN="CENTER">58</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.41</td>
<td ALIGN="CENTER">2.14</td>
<td ALIGN="CENTER">103</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">4.36</td>
<td ALIGN="CENTER">2.93</td>
<td ALIGN="CENTER">75</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">4.32</td>
<td ALIGN="CENTER">2.19</td>
<td ALIGN="CENTER">22</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">3.71</td>
<td ALIGN="CENTER">1.88</td>
<td ALIGN="CENTER">62</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">3.67</td>
<td ALIGN="CENTER">2.24</td>
<td ALIGN="CENTER">21</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">5.09</td>
<td ALIGN="CENTER">1.74</td>
<td ALIGN="CENTER">100</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">3.80</td>
<td ALIGN="CENTER">2.44</td>
<td ALIGN="CENTER">44</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">4.81</td>
<td ALIGN="CENTER">1.86</td>
<td ALIGN="CENTER">53</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">5.66</td>
<td ALIGN="CENTER">4.25</td>
<td ALIGN="CENTER">77</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">3.59</td>
<td ALIGN="CENTER">1.91</td>
<td ALIGN="CENTER">17</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">7.44</td>
<td ALIGN="CENTER">2.78</td>
<td ALIGN="CENTER">27</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">5.81</td>
<td ALIGN="CENTER">1.97</td>
<td ALIGN="CENTER">32</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">6.55</td>
<td ALIGN="CENTER">1.42</td>
<td ALIGN="CENTER">33</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">6.20</td>
<td ALIGN="CENTER">2.86</td>
<td ALIGN="CENTER">5</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">3.92</td>
<td ALIGN="CENTER">2.52</td>
<td ALIGN="CENTER">53</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">4.68</td>
<td ALIGN="CENTER">2.33</td>
<td ALIGN="CENTER">19</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">4.41</td>
<td ALIGN="CENTER">2.35</td>
<td ALIGN="CENTER">54</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">7.66</td>
<td ALIGN="CENTER">2.15</td>
<td ALIGN="CENTER">41</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">5.00</td>
<td ALIGN="CENTER">2.96</td>
<td ALIGN="CENTER">25</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">3.16</td>
<td ALIGN="CENTER">1.00</td>
<td ALIGN="CENTER">82</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">5.35</td>
<td ALIGN="CENTER">1.11</td>
<td ALIGN="CENTER">17</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">7.00</td>
<td ALIGN="CENTER">1.57</td>
<td ALIGN="CENTER">22</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">7.96</td>
<td ALIGN="CENTER">.54</td>
<td ALIGN="CENTER">25</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.52</td>
<td ALIGN="CENTER">1.68</td>
<td ALIGN="CENTER">44</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">9.38</td>
<td ALIGN="CENTER">2.05</td>
<td ALIGN="CENTER">42</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.71</td>
<td ALIGN="CENTER">.85</td>
<td ALIGN="CENTER">17</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.10</td>
<td ALIGN="CENTER">2.28</td>
<td ALIGN="CENTER">10</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">6.38</td>
<td ALIGN="CENTER">.97</td>
<td ALIGN="CENTER">21</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">5.92</td>
<td ALIGN="CENTER">1.38</td>
<td ALIGN="CENTER">13</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">7.49</td>
<td ALIGN="CENTER">.77</td>
<td ALIGN="CENTER">59</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">6.55</td>
<td ALIGN="CENTER">1.61</td>
<td ALIGN="CENTER">20</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">6.42</td>
<td ALIGN="CENTER">1.32</td>
<td ALIGN="CENTER">36</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.57</td>
<td ALIGN="CENTER">1.75</td>
<td ALIGN="CENTER">46</td>
</tr>

</table>
<h5>
</body>
</html>
