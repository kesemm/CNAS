<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NonGeo_FormC_FillForm.aspx.vb" Inherits="CNAS_Supp.NonGeo_FormC_FillForm" %>

<!DOCTYPE html>
<%
    '****************************************************************************************
    '* Created by:    Kelly T. Walsh (Leidos Canada)
    '* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
    '* Purpose:       This is the code behind page for the NonGeo_FormC_FillForm.aspx file.
    '*                This page is an In-Service form for Non-Geographic code assignments. It will
    '*                collect all In-Service information and enter it into the CNAS database. If OK, it
    '*                will send the user to the result page that will show the tix and info entered.
    '* CVS File:      NonGeo_FormC_FillForm.aspx,v
    '* Commit Date:   2018/01/25 16:43:49 (UTC)
    '* Committed by:  saic-ottawa\walshkel
    '* CVS Revision:  1.1
    '* Checkout Tag:  $Name$ (Version/Build)
    '**************************************************************************************** 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Non-Geographic Code - Form C (In-Service)</title>
    <style type="text/css">
        .auto-style1
        {
            width: 104px;
        }
        .auto-style4
        {
            width: 134px;
        }
        .auto-style7
        {
            height: 19px;
            width: 65px;
        }
        .auto-style10
        {
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
        .auto-style38
        {
            font-size: x-small;
        }
        .auto-style40
        {
            text-align: right;
            width: 297px;
        }
        .auto-style43
        {
            text-align: left;
        }
        .auto-style44
        {
            text-align: right;
            width: 330px;
        }
        .auto-style45
        {
            height: 24px;
        }
        </style>
</head>
<body bgColor="#d7c7a4">
   <form id="form1" runat="server">
 <div>        <P style="FONT-FAMILY: arial" align="center"><STRONG><SPAN style="FONT-WEIGHT: bold" class="auto-style34"><span style="TEXT-ALIGN: center">Non-Geographic Code - (In-Service) Form C</span></SPAN></STRONG></P>
			<P style="FONT-FAMILY: arial" align="center"></P>
	
			<P style="FONT-FAMILY: arial">
				<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="50%" align="center" border="0" class="auto-style45">
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
                        <td class="auto-style40">Form C Date:</td>
                        <td>
                            <asp:TextBox ID="txtFormCDate" runat="server" TextMode="Date" MaxLength="10" Width="90px"></asp:TextBox>
                            <span class="auto-style38">&nbsp;(dd/mm/ccyy)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="linkFormCDateToday" runat="server">today</asp:LinkButton>
                            </span></td>
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
			<P style="FONT-FAMILY: arial" align="center">
                	<table align="center" style="width: 65%;">
                    <tr>
                        <td class="auto-style44">In Service Date:</td>
                        <td class="auto-style43">
                            <asp:TextBox ID="txtInServiceDate" runat="server" MaxLength="10" Width="90px"></asp:TextBox>
                            <span class="auto-style38">&nbsp;(dd/mm/ccyy)</span>&nbsp;<span class="auto-style38">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="linkInServiceDateFetchEffectiveDateFormA" runat="server">fetch Form A Effective Date</asp:LinkButton>
                            </span></td>
                    </tr>
                </table>
                </P>
     
     
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
