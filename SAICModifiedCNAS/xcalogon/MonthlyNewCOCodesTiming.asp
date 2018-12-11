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
<p align="center"><strong>Processing Timing For New CO Codes Requests (in working days)</strong></p>
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
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">7.59</td>
<td ALIGN="CENTER">2.09</td>
<td ALIGN="CENTER">69</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">6.76</td>
<td ALIGN="CENTER">1.62</td>
<td ALIGN="CENTER">46</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">6.80</td>
<td ALIGN="CENTER">1.99</td>
<td ALIGN="CENTER">76</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.54</td>
<td ALIGN="CENTER">1.48</td>
<td ALIGN="CENTER">39</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.31</td>
<td ALIGN="CENTER">1.69</td>
<td ALIGN="CENTER">93</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.01</td>
<td ALIGN="CENTER">1.99</td>
<td ALIGN="CENTER">101</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">4.95</td>
<td ALIGN="CENTER">1.39</td>
<td ALIGN="CENTER">66</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">13</td>
<td ALIGN="CENTER">7.57</td>
<td ALIGN="CENTER">1.68</td>
<td ALIGN="CENTER">137</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">5.81</td>
<td ALIGN="CENTER">1.57</td>
<td ALIGN="CENTER">97</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.66</td>
<td ALIGN="CENTER">1.89</td>
<td ALIGN="CENTER">59</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.29</td>
<td ALIGN="CENTER">1.34</td>
<td ALIGN="CENTER">125</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2007</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">17</td>
<td ALIGN="CENTER">7.63</td>
<td ALIGN="CENTER">2.16</td>
<td ALIGN="CENTER">120</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">5.96</td>
<td ALIGN="CENTER">1.22</td>
<td ALIGN="CENTER">51</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">1</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">6.83</td>
<td ALIGN="CENTER">2.23</td>
<td ALIGN="CENTER">80</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.58</td>
<td ALIGN="CENTER">2.10</td>
<td ALIGN="CENTER">111</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">14</td>
<td ALIGN="CENTER">6.86</td>
<td ALIGN="CENTER">1.99</td>
<td ALIGN="CENTER">56</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.61</td>
<td ALIGN="CENTER">2.08</td>
<td ALIGN="CENTER">61</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">8.68</td>
<td ALIGN="CENTER">1.20</td>
<td ALIGN="CENTER">71</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">5.66</td>
<td ALIGN="CENTER">2.11</td>
<td ALIGN="CENTER">74</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">5.46</td>
<td ALIGN="CENTER">1.91</td>
<td ALIGN="CENTER">112</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.08</td>
<td ALIGN="CENTER">1.60</td>
<td ALIGN="CENTER">142</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">6.42</td>
<td ALIGN="CENTER">2.13</td>
<td ALIGN="CENTER">67</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">16</td>
<td ALIGN="CENTER">6.88</td>
<td ALIGN="CENTER">2.37</td>
<td ALIGN="CENTER">52</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2006</td>
<td ALIGN="CENTER">2</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">4.71</td>
<td ALIGN="CENTER">1.45</td>
<td ALIGN="CENTER">56</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">8.42</td>
<td ALIGN="CENTER">1.78</td>
<td ALIGN="CENTER">55</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.57</td>
<td ALIGN="CENTER">1.66</td>
<td ALIGN="CENTER">69</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.48</td>
<td ALIGN="CENTER">1.49</td>
<td ALIGN="CENTER">48</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.79</td>
<td ALIGN="CENTER">1.14</td>
<td ALIGN="CENTER">47</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">13</td>
<td ALIGN="CENTER">7.85</td>
<td ALIGN="CENTER">1.66</td>
<td ALIGN="CENTER">39</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.40</td>
<td ALIGN="CENTER">1.77</td>
<td ALIGN="CENTER">78</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">6.91</td>
<td ALIGN="CENTER">1.70</td>
<td ALIGN="CENTER">70</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">8.39</td>
<td ALIGN="CENTER">1.98</td>
<td ALIGN="CENTER">51</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">3</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">5.77</td>
<td ALIGN="CENTER">2.29</td>
<td ALIGN="CENTER">44</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">6</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">8.20</td>
<td ALIGN="CENTER">1.31</td>
<td ALIGN="CENTER">60</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">13</td>
<td ALIGN="CENTER">7.82</td>
<td ALIGN="CENTER">2.60</td>
<td ALIGN="CENTER">49</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2005</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">7.03</td>
<td ALIGN="CENTER">1.88</td>
<td ALIGN="CENTER">39</td>
</tr>

<tr>
<td ALIGN="CENTER">Jan</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.33</td>
<td ALIGN="CENTER">1.99</td>
<td ALIGN="CENTER">15</td>
</tr>

<tr>
<td ALIGN="CENTER">Feb</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.70</td>
<td ALIGN="CENTER">1.61</td>
<td ALIGN="CENTER">33</td>
</tr>

<tr>
<td ALIGN="CENTER">Mar</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">12</td>
<td ALIGN="CENTER">8.93</td>
<td ALIGN="CENTER">1.57</td>
<td ALIGN="CENTER">42</td>
</tr>

<tr>
<td ALIGN="CENTER">Apr</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.80</td>
<td ALIGN="CENTER">1.72</td>
<td ALIGN="CENTER">54</td>
</tr>

<tr>
<td ALIGN="CENTER">May</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.07</td>
<td ALIGN="CENTER">1.51</td>
<td ALIGN="CENTER">29</td>
</tr>

<tr>
<td ALIGN="CENTER">Jun</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">8.86</td>
<td ALIGN="CENTER">1.83</td>
<td ALIGN="CENTER">14</td>
</tr>

<tr>
<td ALIGN="CENTER">Jul</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">8</td>
<td ALIGN="CENTER">5.18</td>
<td ALIGN="CENTER">1.22</td>
<td ALIGN="CENTER">50</td>
</tr>

<tr>
<td ALIGN="CENTER">Aug</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">7</td>
<td ALIGN="CENTER">5.92</td>
<td ALIGN="CENTER">1.08</td>
<td ALIGN="CENTER">25</td>
</tr>

<tr>
<td ALIGN="CENTER">Sep</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">5</td>
<td ALIGN="CENTER">9</td>
<td ALIGN="CENTER">6.57</td>
<td ALIGN="CENTER">1.17</td>
<td ALIGN="CENTER">30</td>
</tr>

<tr>
<td ALIGN="CENTER">Oct</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">6.60</td>
<td ALIGN="CENTER">2.06</td>
<td ALIGN="CENTER">35</td>
</tr>

<tr>
<td ALIGN="CENTER">Nov</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">10</td>
<td ALIGN="CENTER">6.79</td>
<td ALIGN="CENTER">2.11</td>
<td ALIGN="CENTER">63</td>
</tr>

<tr>
<td ALIGN="CENTER">Dec</td>
<td ALIGN="CENTER">2004</td>
<td ALIGN="CENTER">4</td>
<td ALIGN="CENTER">11</td>
<td ALIGN="CENTER">7.13</td>
<td ALIGN="CENTER">2.20</td>
<td ALIGN="CENTER">63</td>
</tr>

</table>
<h5>
</body>
</html>
