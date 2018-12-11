<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NonGeo_RqstViewCode.aspx.vb" Inherits="CNAS_Supp.NonGeo_RqstViewCode" %>

<!DOCTYPE html>
<%
'****************************************************************************************
'* Created by:    Kelly T. Walsh (Leidos Canada)
'* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
'* Purpose:       ASP.Net Page - NonGeo_FormA_FillForm.aspx
'*                This page is an application form for Non-Geographic codes. It will
'*                let the user select a NonGeo code (npa/nxx) then send them to the
'*                viewing page (NonGeo_ViewCode) to see everything about it.
'* CVS File:      NonGeo_RqstViewCode.aspx,v
'* Commit Date:   2018/01/25 16:43:49 (UTC)
'* Committed by:  saic-ottawa\walshkel
'* CVS Revision:  1.1
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">

        .auto-style34
        {
            font-size: x-large;
            font-weight: bold;
        }
        #Table1
        {
            height: 67px;
        }
        .auto-style27
        {
            height: 50px;
        }
        .auto-style4
        {
            height: 19px;
            width: 134px;
        }
        .auto-style1
        {
            height: 19px;
            width: 104px;
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
        .auto-style35
        {
            text-align: right;
        }
        </style>
</head>
<body bgColor="#d7c7a4">
    <P style="FONT-FAMILY: arial" align="center"><SPAN style="FONT-WEIGHT: bold" class="auto-style34"><STRONG><span style="TEXT-ALIGN: center">Non-Geographic Code - Request </span></STRONG>Code View</SPAN></P>
			<P style="FONT-FAMILY: arial" align="center"><b>Select a Code to View</b></P>
	
    <form id="form1" defaultbutton="btnSubmit" runat="server" defaultfocus="txtNXX">
	
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
					</TABLE>
			</P>
    <div>
    <div align="center"><asp:Label ID="lblNotifications" runat="server" Font-Bold="True" Font-Italic="True" Font-Names="Arial" ForeColor="#CC0000" BorderStyle="None" Visible="False"></asp:Label></div> 
        <br />
    
    </div>
            <p class="auto-style35" style="FONT-FAMILY: arial">
                <asp:Button ID="btnCancel" runat="server" Text="Cancel" />
&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnSubmit" runat="server" Text="Submit" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            </p>
     <P style="FONT-FAMILY: arial" class="auto-style13">
				&nbsp;</P>
     <P style="FONT-FAMILY: arial" class="auto-style13">
				&nbsp;</P>
     <P style="FONT-FAMILY: arial" class="auto-style13">
				&nbsp;</P>
        
        
    			<CENTER class="auto-style14"><FONT color="gray" face="Arial" size="1"><B>Leidos Canada CNAS_Supp_vs2017 Version Control 
						Information: $Revision: 1.3 $ $&nbsp;&nbsp;&nbsp; $Date: 2018/01/26 16:42:37 $ 
					(UTC)


    </form>
</body>
</html>
