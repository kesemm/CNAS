<%@ Page MaintainScrollPositionOnPostback="true" Language="vb" AutoEventWireup="false" CodeBehind="NonGeo_FormA_FillForm.aspx.vb" Inherits="CNAS_Supp.NonGeo_FormA_FillForm" %>
<!DOCTYPE html>
<%
    '****************************************************************************************
    '* Created by:    Kelly T. Walsh (Leidos Canada)
    '* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
    '* Purpose:       ASP.Net Page - NonGeo_FormA_FillForm.aspx
    '*                This page is an application form for Non-Geographic code assignments. It will
    '*                default to the next available code, but allow manual entry, then collect all
    '*                the application information and enter it into the CNAS database. If OK, it
    '*                will send the user to the result page that will show the tix and info entered.
    '* CVS File:      NonGeo_FormA_FillForm.aspx,v
    '* Commit Date:   2018/01/25 16:43:49 (UTC)
    '* Committed by:  saic-ottawa\walshkel
    '* CVS Revision:  1.1
    '* Checkout Tag:  $Name$ (Version/Build)
    '**************************************************************************************** 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Non-Geographic Code - Form A</title>
    <style type="text/css">
        .auto-style1
        {
            height: 19px;
            width: 104px;
        }
        .auto-style4
        {
            height: 19px;
            width: 134px;
        }
        .auto-style6
        {
            height: 1px;
            font-size: xx-small;
            text-align: center;
        }
        .auto-style7
        {
            height: 19px;
            width: 65px;
        }
        .auto-style10
        {
            height: 19px;
            width: 193px;
        }
        #Table1
        {
            height: 67px;
        }
        .auto-style11
        {
            width: 80%;
        }
        .auto-style13
        {
            text-align: right;
        }
        .auto-style14
        {
            height: 0px;
        }
        .auto-style16
        {
            width: 134px;
            text-align: right;
            font-size: small;
        }
        .auto-style17
        {
            width: 125px;
            text-align: right;
            font-size: small;
        }
        .auto-style18
        {
            text-align: center;
        }
        .auto-style20
        {
            margin-bottom: 0px;
        }
        .auto-style21
        {
            width: 802px;
            text-align: left;
        }
        .auto-style22
        {
            width: 80%;
            font-weight: bold;
            font-style: italic;
        }
        .auto-style23
        {
            font-size: small;
        }
        .auto-style24
        {
            height: 22px;
            font-size: small;
            text-align: right;
            font-weight: bold;
        }
        .auto-style25
        {
            text-align: center;
            height: 22px;
            font-weight: bold;
        }
        .auto-style26
        {
            font-style: italic;
            font-weight: bold;
        }
        .auto-style27
        {
            height: 50px;
        }
        .auto-style28
        {
            width: 994px;
            font-family: arial;
            font-weight: bold;
            font-style: italic;
            color: #990033;
        }
        .auto-style29
        {
            color: #990033;
        }
        .auto-style31
        {
            font-size: small;
            text-align: right;
            font-weight: bold;
        }
        .auto-style32
        {
            text-align: center;
            font-weight: bold;
        }
        .auto-style33
        {
            width: 63%;
        }
        .auto-style34
        {
            font-size: x-large;
            font-weight: bold;
        }
        .auto-style35
        {
            width: 125px;
            text-align: right;
            height: 22px;
            font-size: small;
        }
        .auto-style36
        {
            height: 22px;
        }
        .auto-style37
        {
            font-weight: bold;
            text-align: right;
            width: 229px;
        }
        .auto-style38
        {
            font-size: x-small;
        }
        .auto-style39
        {
            font-weight: normal;
            font-size: small;
        }
        .auto-style40
        {
            text-align: right;
            width: 297px;
        }
        .auto-style41
        {
            text-align: right;
            width: 297px;
            height: 29px;
        }
        .auto-style42
        {
            height: 29px;
        }
        </style>
</head>
<body bgColor="#d7c7a4">
   <form id="form1" runat="server">
 <div>        <P style="FONT-FAMILY: arial" align="center"><STRONG><SPAN style="FONT-WEIGHT: bold" class="auto-style34"><span style="TEXT-ALIGN: center">Non-Geographic Code - Form A</span></SPAN></STRONG></P>
			<P style="FONT-FAMILY: arial" align="center"></P>
	
			<P style="FONT-FAMILY: arial">
				<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="50%" align="center" border="0" class="auto-style27">
					<TR>
						<TD class="auto-style4">
							<P style="FONT-FAMILY: arial; width: 161px;" align="right"><STRONG>NPA:</STRONG>&nbsp;
							</P>
						</TD>
						<TD class="auto-style1">
							<asp:DropDownList id=comboNPA runat="server" Width="66px" DataSource="<%# dtNPAList %>" DataTextField="NPA" DataValueField="NPA" AutoPostBack="True">
							</asp:DropDownList></TD>
						<TD class="auto-style7">
							<P align="right"><STRONG>NXX:</STRONG>&nbsp;&nbsp;
							</P>
						</TD>
						<TD class="auto-style10">
							<asp:TextBox id="txtNXX" runat="server" Width="43px" MaxLength="3"></asp:TextBox></TD>
					</TR>
					<TR>
						<TD class="auto-style6" colspan="4">
							Initial page load and changing the NPA will fill the NXX with the next code available</TD>
					</TR>
				    </TABLE>
			</P>
   		<div align="center"><asp:Label ID="lblNotifications" runat="server" Font-Bold="True" Font-Italic="True" Font-Names="Arial" ForeColor="#CC0000" BorderStyle="None" Visible="False"></asp:Label></div>  <P style="FONT-FAMILY: arial">
				<table align="center" style="width: 65%;">
                    <tr>
                        <td class="auto-style40">Authorized Representative Name:</td>
                        <td>
                            <asp:TextBox ID="txtAuthRepName" runat="server" Width="302px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style40">Title:</td>
                        <td>
                            <asp:TextBox ID="txtAuthRepTitle" runat="server" Width="349px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style40">OCN:</td>
                        <td>
                            <asp:TextBox ID="txtOCN" runat="server" Width="53px" MaxLength="4"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style40">Application Date:</td>
                        <td>
                            <asp:TextBox ID="txtApplicationDate" runat="server" TextMode="Date" MaxLength="10" Width="90px"></asp:TextBox>
                            <span class="auto-style38">&nbsp;(dd/mm/ccyy)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="linkApplicationDateToday" runat="server">today</asp:LinkButton>
                            </span></td>
                    </tr>
                    <tr>
                        <td class="auto-style41">Date of Receipt:</td>
                        <td class="auto-style42">
                            <asp:TextBox ID="txtDateOfReceipt" runat="server" MaxLength="10" Width="90px"></asp:TextBox>
                            <span class="auto-style38">&nbsp;(dd/mm/ccyy)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="linkDateOfRcptToday" runat="server">today</asp:LinkButton>
                            </span>
                        &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="auto-style40">Last Correspondence Date:</td>
                        <td>
                            <asp:TextBox ID="txtLastCorrespondenceDate" runat="server" MaxLength="10" Width="90px"></asp:TextBox>
                            <span class="auto-style38">&nbsp;(dd/mm/ccyy)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="linkLastCorrespondenceDateToday" runat="server">today</asp:LinkButton>
                            </span>
                        &nbsp;</td>
                    </tr>
                    </table>
			</P>
     <P style="FONT-FAMILY: arial">
				<table align="center" class="auto-style11">
                    <tr>
                        <td>
                            <div>
                                <b>Code Applicant Info:</b><asp:Label ID="lblEntityID" runat="server" Visible="False"></asp:Label>
                                <br />
                            </div>
                            <table style="width:100%;">
                                <tr>
                                    <td class="auto-style16">Entity Name:</td>
                                    <td>
                                        <asp:Label ID="lblAppEntityName" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Contact Name:</td>
                                    <td>
                                        <asp:Label ID="lblAppContactName" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Street Address:</td>
                                    <td>
                                        <asp:Label ID="lblAppStreetAddress" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">City:</td>
                                    <td>
                                        <asp:Label ID="lblAppCity" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Province:</td>
                                    <td>
                                        <asp:Label ID="lblAppProvince" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Postal Code:</td>
                                    <td>
                                        <asp:Label ID="lblAppPostalCode" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">E-Mail Address:</td>
                                    <td>
                                        <asp:Label ID="lblAppEmail" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Facsimile:</td>
                                    <td>
                                        <asp:Label ID="lblAppFax" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Telephone:</td>
                                    <td>
                                        <asp:Label ID="lblAppPhone" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style16">Extension:</td>
                                    <td>
                                        <asp:Label ID="lblAppExtension" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td>&nbsp;</td>
                        <td>
                            <div>
                                <b>CNA Info:</b><br />
                            </div>
                            <table style="width:100%;">
                                <tr>
                                    <td class="auto-style17">Entity Name:</td>
                                    <td>
                                        <asp:Label ID="lblCNAEntityName" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Contact Name:</td>
                                    <td>
                                        <asp:Label ID="lblCNAContactName" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Street Address:</td>
                                    <td>
                                        <asp:Label ID="lblCNAStreetAddress" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style35">City:</td>
                                    <td class="auto-style36">
                                        <asp:Label ID="lblCNACity" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Province:</td>
                                    <td>
                                        <asp:Label ID="lblCNAProvince" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Postal Code:</td>
                                    <td>
                                        <asp:Label ID="lblCNAPostalCode" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">E-Mail Address:</td>
                                    <td>
                                        <asp:Label ID="lblCNAEmail" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Facsimile:</td>
                                    <td>
                                        <asp:Label ID="lblCNAFax" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Telephone:</td>
                                    <td>
                                        <asp:Label ID="lblCNAPhone" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style17">Extension:</td>
                                    <td>
                                        <asp:Label ID="lblCNAExtension" runat="server" Font-Bold="True" Font-Names="Arial" Font-Overline="False" Font-Size="Medium"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
			</P>
				<table align="center" style="width: 65%;">
                    <tr>
                        <td class="auto-style37">Name of Service:</td>
                        <td>
                            <asp:TextBox ID="txtNameOfService" runat="server" Width="499px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style37">Effective Date:</td>
                        <td>
                            <asp:TextBox ID="txtEffectiveDate" runat="server" MaxLength="10" Width="90px"></asp:TextBox>
                            <span class="auto-style38">&nbsp;(dd/mm/ccyy)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="linkEffectiveDateAppPlus21Days" runat="server">21 days from application date</asp:LinkButton>
                            </span>
                        &nbsp;</td>
                    </tr>
                </table>
			<div  align="center">
 
			<P style="FONT-FAMILY: arial" align="center">
				<b>Type of Request</b>
         <asp:RadioButtonList ID="rblTypeOfRequest" runat="server" AutoPostBack="True" CssClass="auto-style20" RepeatDirection="Horizontal" Width="294px">
                    <asp:ListItem Value="Initial" Selected="True">Initial Code</asp:ListItem>
                    <asp:ListItem Value="Growth">Growth Code</asp:ListItem>
                </asp:RadioButtonList>
			</P>
     <asp:Panel ID="panelApproval" runat="server" Height="255px" BorderColor="#990033" BorderStyle="Dotted" Width="900px" HorizontalAlign="Left">
        <P class="auto-style28">Regulatory Approval (*section required for Initial Code)</p>
         <blockquote>   <p class="auto-style21" style="FONT-FAMILY: arial">
               If regulatory approval is required to provide the Non-Geographic service, and you have such approval, indicate the type of approval (e.g. CRTC letter, license, approved tariff, etc.) and date, and attach a copy of the approval if not previously submitted.<br />
               <asp:TextBox ID="txtApprovalRequired" runat="server" Width="779px"></asp:TextBox>
               <br />
               <br />
               If regulatory approval is not required, describe the document that confirms regulatory approval is not required, and attach a copy if not previously submitted.<br />
               <asp:TextBox ID="txtApprovalNotRequired" runat="server" Width="779px"></asp:TextBox>
               <br />
           </p>
         </blockquote>
     </asp:Panel>
     <asp:Panel ID="panelExhaustCalc" runat="server" Height="870px" Visible="False" BorderColor="#990033" BorderStyle="Dotted" Width="900px" HorizontalAlign="Left">
         <P style="FONT-FAMILY: arial" class="auto-style29">
				<i><b>Growth History, Forecast and Months-to-Exhaust Table (*section required for Growth Code)</b></i><p style="FONT-FAMILY: arial">
                    &nbsp;<table align="center" class="auto-style11">
                        <tr>
                            <td class="auto-style22" colspan="3">Summary of existing Codes<span class="auto-style39"> (Assigned, Reserved and Pending)</span></td>
                        </tr>
                        <tr>
                            <td>Type of Code</td>
                            <td>Quantity of Codes</td>
                            <td>List Assigned and Reserved Codes <i><b><span class="auto-style23">(no commas!)</span></b></i></td>
                        </tr>
                        <tr>
                            <td>Assigned</td>
                            <td>
                                <asp:TextBox ID="txtSummaryAssigned" runat="server" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td>
                                <asp:TextBox ID="txtSummaryListOfAssigned" runat="server" Width="504px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td>Reserved</td>
                            <td>
                                <asp:TextBox ID="txtSummaryReserved" runat="server" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td>
                                <asp:TextBox ID="txtSummaryListOfReserved" runat="server" Width="506px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td>Pending Assignment</td>
                            <td>
                                <asp:TextBox ID="txtSummaryPendingAssignment" runat="server" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td>&nbsp;</td>
                        </tr>
                        <tr>
                            <td>Pending Reservation</td>
                            <td>
                                <asp:TextBox ID="txtSummaryPendingReservation" runat="server" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td>&nbsp;</td>
                        </tr>
                    </table>
                    <br />
                    <table align="center" class="auto-style11
             ">
                        <tr>
                            <td colspan="7"><i><b>Previous 6-Month growth history</b><span class="auto-style23"> (quantity of numbers assigned to customers in the previous 6 months)</span></i></td>
                        </tr>
                        <tr>
                            <td>Month #</td>
                            <td class="auto-style18">-6</td>
                            <td class="auto-style18">-5</td>
                            <td class="auto-style18">-4</td>
                            <td class="auto-style18">-3</td>
                            <td class="auto-style18">-2</td>
                            <td class="auto-style18">-1</td>
                        </tr>
                        <tr>
                            <td>Quantity each month</td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtHistory6" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtHistory5" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtHistory4" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtHistory3" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtHistory2" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtHistory1" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <table align="center" class="auto-style33">
                        <tr>
                            <td colspan="2" class="auto-style18"><b><i>Quantity of numbers available for assignment</i></b></td>
                        </tr>
                        <tr>
                            <td class="auto-style13">Quantity of numbers available in Assigned Codes:</td>
                            <td>
                                <asp:TextBox ID="txtAvailableAssigned" runat="server" Style="text-align: right">0</asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style13">Quantity of numbers in Codes Pending Assignment:</td>
                            <td>
                                <asp:TextBox ID="txtAvailablePending" runat="server" Style="text-align: right">0</asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style13">Total Quantity of Numbers Available for Assignment:</td>
                            <td>
                                <asp:TextBox ID="txtAvailableTotal" runat="server" BackColor="#F3F3F3" ReadOnly="True" Style="text-align: right"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                   
                    <br />
                   
                    <table align="center" class="auto-style11
             ">
                        <tr>
                            <td colspan="7"><i><b>Projected 12-month growth forecast</b></i><br /> <span class="auto-style23">(quantity of numbers forecast to be assigned to customers each month during the coming 12 months)</span></td>
                        </tr>
                        <tr>
                            <td class="auto-style24">Month #</td>
                            <td class="auto-style25">1</td>
                            <td class="auto-style25">2</td>
                            <td class="auto-style25">3</td>
                            <td class="auto-style25">4</td>
                            <td class="auto-style25">5</td>
                            <td class="auto-style25">6</td>
                        </tr>
                        <tr>
                            <td class="auto-style31">Quantity each month</td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast1" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast2" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast3" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast4" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast5" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast6" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style31">Month #</td>
                            <td class="auto-style32">7</td>
                            <td class="auto-style32">8</td>
                            <td class="auto-style32">9</td>
                            <td class="auto-style32">10</td>
                            <td class="auto-style32">11</td>
                            <td class="auto-style32">12</td>
                        </tr>
                        <tr>
                            <td class="auto-style31">Quantity each month</td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast7" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast8" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast9" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast10" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast11" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                            <td class="auto-style18">
                                <asp:TextBox ID="txtForecast12" runat="server" Width="80px" Style="text-align: right">0</asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <table align="center" class="auto-style11
             ">
                        <tr>
                            <td class="auto-style26">Average Monthly Growth Rate (months 1 to 12)</td>
                            <td>
                                <asp:TextBox ID="txtAverageGrowthRate" runat="server" BackColor="#F3F3F3" ReadOnly="True" Style="text-align: right"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style26">Months-to-Exhaust <span class="auto-style39">(calculated using average growth; different from application sheet!)</span></td>
                            <td>
                                <asp:TextBox ID="txtMonthsToExhaust" runat="server" BackColor="#F3F3F3" ReadOnly="True" Style="text-align: right"></asp:TextBox>
                            </td>
                        </tr>
                        </caption>
                    </table>
                    <br />
                    <table align="center" class="auto-style11
             ">
                        <tr>
                            <td><b><i>Other relevant information to support this request for an Additional Code Assignment or Reserveration</i></b></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:TextBox ID="txtGrowthInfo" runat="server" Width="710px"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                         <table align="center" class="auto-style11
             ">
                             <tr>
                                 <td>&nbsp;</td>
                                 <td class="auto-style13">
                                     <asp:Button ID="btnCalculateTotals" runat="server" Height="26px" Text="Calculate Totals Now" Width="157px" />
                                 </td>
                             </tr>
                    </table>
                </asp:Panel>
      </div>  
     
     
      <P style="FONT-FAMILY: arial" class="auto-style13">
				<asp:Button ID="btnCancel" runat="server" Text="Cancel" />
&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnSubmit" runat="server" Text="Submit" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </P>
     <P style="FONT-FAMILY: arial" class="auto-style13">
				&nbsp;</P>
     <P style="FONT-FAMILY: arial" class="auto-style13">
				&nbsp;</P>
     <P style="FONT-FAMILY: arial" class="auto-style13">
				&nbsp;</P>
        
        
    			<CENTER class="auto-style14"><FONT color="gray" face="Arial" size="1"><B>Leidos Canada CNAS_Supp_vs2017 Version Control 
						Information: $Revision: 1.2 $ $&nbsp;&nbsp;&nbsp; $Date: 2018/01/26 16:42:37 $ 
					(UTC)

</div>    </form>
</body>
</html>
