<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CodeValidator.aspx.vb" Inherits="CNAS_Supp.CodeValidator"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<%
'****************************************************************************************
'* Created by:    Kelly T. Walsh (Leidos Canada)
'* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
'* Purpose:       ASP.Net Page - CodeValidator.aspx
'*                It provides a formated report tool, giving CO Code administrators information
'*                about EAS relationships and extra code notifications required when assigning
'*                CO Codes.
'* CVS File:      CodeValidator.aspx,v
'* Commit Date:   2018/01/25 16:43:49 (UTC)
'* Committed by:  saic-ottawa\walshkel
'* CVS Revision:  1.1
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
	<HEAD>
		<title>CodeValidator</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body bgColor="#d7c7a4">
		<form id="Form1" method="post" runat="server">
			<P style="FONT-FAMILY: arial" align="center"><STRONG><BIG><BIG><SPAN style="FONT-WEIGHT: bold"><B><SMALL style="TEXT-ALIGN: center">CNA 
										- Requested Code Validator</SMALL></B></SPAN></BIG></BIG></STRONG></P>
			<P style="FONT-FAMILY: arial" align="center"><STRONG>Enter your parameters then click 
					the Validate button</STRONG></P>
			<P style="FONT-FAMILY: arial">
				<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="50%" align="center" border="0">
					<TR>
						<TD style="WIDTH: 199px; HEIGHT: 44px">
							<P style="FONT-FAMILY: arial" align="right"><STRONG>NPA:</STRONG>&nbsp;
							</P>
						</TD>
						<TD style="HEIGHT: 44px">
							<asp:DropDownList id=comboNPA runat="server" Width="66px" DataSource="<%# dtNPAList %>" DataTextField="NPA" DataValueField="NPA" AutoPostBack="True">
							</asp:DropDownList></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 199px; HEIGHT: 36px">
							<P align="right"><STRONG>Rate Center:</STRONG>&nbsp;
							</P>
						</TD>
						<TD style="HEIGHT: 36px">
							<asp:DropDownList id=comboRateCenter runat="server" Width="245px" DataSource="<%# dtRateCenterList %>" DataTextField="RateCenter" DataValueField="RateCenter" tabIndex=1>
							</asp:DropDownList></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 199px; HEIGHT: 43px">
							<P align="right"><STRONG>NXX:</STRONG>&nbsp;&nbsp;
							</P>
						</TD>
						<TD style="HEIGHT: 43px">
							<asp:TextBox id="txtNXX" runat="server" Width="43px" MaxLength="3" tabIndex="2"></asp:TextBox></TD>
					</TR>
				</TABLE>
			</P>
			<P align="center">
				<asp:Button id="btnValidate" runat="server" Text="Generate Validation Report" Width="204px"
					tabIndex="3" EnableViewState="False" CausesValidation="False"></asp:Button></P>
			<P>&nbsp;</P>
			<P>
				<TABLE id="Table2" cellSpacing="1" cellPadding="1" width="50" border="0" align="center">
					<TR>
						<TD noWrap>
							<P align="center">
								<asp:Label id="StatusPanel" runat="server" Width="494px" Height="102px" Font-Bold="True" ForeColor="DarkRed"
									Font-Names="arial" EnableViewState="False"></asp:Label></P>
						</TD>
					</TR>
				</TABLE>
			</P>
			<CENTER><FONT color="gray" face="Arial" size="1"><B>Leidos Canada CNAS_Supp_vs2017 Version Control 
						Information:</B>&nbsp;&nbsp;&nbsp; 1.1&nbsp;&nbsp;&nbsp; 2018/01/25 16:43:49 
					(UTC)</FONT></CENTER>
		</form>
	</body>
</HTML>
