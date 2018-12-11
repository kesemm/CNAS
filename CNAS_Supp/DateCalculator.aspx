<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DateCalculator.aspx.vb" Inherits="CNAS_Supp.DateCalculator"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<%
'****************************************************************************************
'* Created by:    Kelly T. Walsh (Leidos Canada)
'* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
'* Purpose:       ASP.Net Page - DateCalculator.aspx
'*                It is used to calculate various CNA related dates for CO Code applications
'*                and administration.
'* CVS File:      DateCalculator.aspx,v
'* Commit Date:   2018/01/25 16:43:49 (UTC)
'* Committed by:  saic-ottawa\walshkel
'* CVS Revision:  1.1
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
	<HEAD>
		<title>DateCalculator</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body bgColor="#d7c7a4">
		<form id="Form1" method="post" runat="server">
			<STRONG>
				<P style="FONT-FAMILY: arial" align="center"><STRONG><BIG><BIG><SPAN style="FONT-WEIGHT: bold"><B><SMALL style="TEXT-ALIGN: center">CNA 
											- Date Calculator</SMALL></B></SPAN></BIG></BIG></STRONG></P>
				<P style="FONT-FAMILY: arial" align="center"><STRONG>Select the reference date and the 
						resulting dates are calculated automatically</STRONG></P>
				<P style="FONT-FAMILY: arial" align="center">&nbsp;</P>
				<P style="FONT-FAMILY: arial" align="center">
					<TABLE id="Table1" style="WIDTH: 625px; HEIGHT: 305px" cellSpacing="1" cellPadding="1"
						width="625" border="0">
						<TR>
							<TD style="WIDTH: 383px; HEIGHT: 251px">
								<DIV align="center">
									<TABLE id="Table3" cellSpacing="1" cellPadding="1" width="100%" border="0">
										<TR>
											<TD style="HEIGHT: 15px">
												<P align="center">
													<asp:DropDownList id="comboMonth" runat="server" Width="98px" AutoPostBack="True" EnableViewState="False">
														<asp:ListItem Value="1">January</asp:ListItem>
														<asp:ListItem Value="2">February</asp:ListItem>
														<asp:ListItem Value="3">March</asp:ListItem>
														<asp:ListItem Value="4">April</asp:ListItem>
														<asp:ListItem Value="5">May</asp:ListItem>
														<asp:ListItem Value="6">June</asp:ListItem>
														<asp:ListItem Value="7">July</asp:ListItem>
														<asp:ListItem Value="8">August</asp:ListItem>
														<asp:ListItem Value="9">September</asp:ListItem>
														<asp:ListItem Value="10">October</asp:ListItem>
														<asp:ListItem Value="11">November</asp:ListItem>
														<asp:ListItem Value="12">December</asp:ListItem>
													</asp:DropDownList>&nbsp;&nbsp;&nbsp;
													<asp:Button id="btnYearDown" runat="server" Height="22px" Text="Year Down" Width="75px" EnableViewState="False"></asp:Button>&nbsp;
													<asp:Button id="btnYearUp" runat="server" Width="69px" Height="22px" Text="Year Up" EnableViewState="False"></asp:Button></P>
											</TD>
										</TR>
										<TR>
											<TD>
												<DIV align="center">
													<asp:Calendar id="Calendar1" runat="server" BackColor="#FFFFCC" Width="272px" DayNameFormat="FirstLetter"
														ForeColor="#663399" Height="184px" Font-Size="8pt" Font-Names="Verdana" BorderColor="#FFCC66"
														BorderWidth="1px" ShowGridLines="True">
														<TodayDayStyle ForeColor="White" BackColor="#FFCC66"></TodayDayStyle>
														<SelectorStyle BackColor="#FFCC66"></SelectorStyle>
														<NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC"></NextPrevStyle>
														<DayHeaderStyle Height="1px" BackColor="#FFCC66"></DayHeaderStyle>
														<SelectedDayStyle Font-Bold="True" BackColor="#CCCCFF"></SelectedDayStyle>
														<TitleStyle Font-Size="9pt" Font-Bold="True" ForeColor="#FFFFCC" BackColor="#990000"></TitleStyle>
														<OtherMonthDayStyle ForeColor="#CC9966"></OtherMonthDayStyle>
													</asp:Calendar></DIV>
											</TD>
										</TR>
										<TR>
											<TD>
												<P align="center">
													<asp:Button id="btnToday" runat="server" Width="135px" Height="22px" Text="Goto Today" EnableViewState="False"></asp:Button></P>
											</TD>
										</TR>
									</TABLE>
								</DIV>
							</TD>
							<TD style="HEIGHT: 251px">
								<TABLE id="Table2" style="WIDTH: 357px; HEIGHT: 144px" cellSpacing="1" cellPadding="1"
									width="357" border="0">
									<TR>
										<TD style="WIDTH: 141px; HEIGHT: 32px">
											<P align="right"><STRONG>+ 66 Days =</STRONG></P>
										</TD>
										<TD style="HEIGHT: 32px">
											<P align="center">
												<asp:TextBox id="txtPlus66" runat="server" Width="165px" Height="20px" Font-Names="arial" ReadOnly="True"
													EnableViewState="False"></asp:TextBox></P>
										</TD>
									</TR>
									<TR>
										<TD style="WIDTH: 141px; HEIGHT: 28px">
											<P align="right"><STRONG>- 45 Days =</STRONG></P>
										</TD>
										<TD style="HEIGHT: 28px">
											<P align="center">
												<asp:TextBox id="txtMinus45" runat="server" Width="165px" Height="20px" Font-Names="arial" ReadOnly="True"
													EnableViewState="False"></asp:TextBox></P>
										</TD>
									</TR>
									<TR>
										<TD style="WIDTH: 141px; HEIGHT: 27px">
											<P align="right"><STRONG>+ 52 Days =</STRONG></P>
										</TD>
										<TD style="HEIGHT: 27px">
											<P align="center">
												<asp:TextBox id="txtPlus52" runat="server" Width="165px" Height="20px" Font-Names="arial" ReadOnly="True"
													EnableViewState="False"></asp:TextBox></P>
										</TD>
									</TR>
									<TR>
										<TD style="WIDTH: 141px; HEIGHT: 35px">
											<P align="right"><STRONG>+ 45 Days =</STRONG></P>
										</TD>
										<TD style="HEIGHT: 35px">
											<P align="center">
												<asp:TextBox id="txtPlus45" runat="server" Width="165px" Height="20px" Font-Names="arial" ReadOnly="True"
													EnableViewState="False"></asp:TextBox></P>
										</TD>
									</TR>
									<TR>
										<TD style="WIDTH: 141px">
											<P align="right">
												<asp:TextBox id="txtCustomDays" runat="server" Width="54px" Height="25px" Font-Size="Larger"
													Font-Names="arial" Wrap="False" MaxLength="5" ToolTip="Put in a Valid integer">0</asp:TextBox><STRONG>&nbsp;Days 
													=</STRONG></P>
										</TD>
										<TD>
											<P align="center">
												<asp:TextBox id="txtCustom" runat="server" Width="165px" Height="20px" Font-Names="arial" ReadOnly="True"
													EnableViewState="False"></asp:TextBox></P>
										</TD>
									</TR>
								</TABLE>
								<P align="right"><STRONG><FONT size="1">Click this button if you only change the custom 
											Days Number</FONT></STRONG></P>
								<P align="right">
									<asp:Button id="btnReCalculate" runat="server" Width="102px" Text="ReCalculate" EnableViewState="False"></asp:Button></P>
							</TD>
						</TR>
						<TR>
							<TD style="WIDTH: 383px"></TD>
							<TD>
								<P align="right"><STRONG><FONT size="1"></FONT></STRONG>&nbsp;</P>
							</TD>
						</TR>
					</TABLE>
				</P>
				<CENTER><FONT color="gray" face="Arial" size="1"><B>Leidos Canada CNAS_Supp_vs2017 Version Control 
							Information:</B>&nbsp;&nbsp;&nbsp; 1.1&nbsp;&nbsp;&nbsp; 2018/01/25 16:43:49 
						(UTC)</FONT></CENTER>
			</STRONG>
		</form>
	</body>
</HTML>
