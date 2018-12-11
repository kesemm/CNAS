<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NonGeo_ViewCode.aspx.vb" Inherits="CNAS_Supp.NonGeo_ViewCode" %>

<!DOCTYPE html>
<%
    '****************************************************************************************
    '* Created by:    Kelly T. Walsh (Leidos Canada)
    '* Project:       CNAS_Supp [CVS Module CNAS_Supp_vs2017] (.Net Framework 4)
    '* Purpose:       ASP.Net Page - NonGeo_ViewCode.aspx
    '*                This page is an application form for Non-Geographic codes. It will
    '*                let the user view the results of a query for one code showing the NonGeo
    '*                information/status, the FormA information and the FormB information.
    '* CVS File:      NonGeo_ViewCode.aspx,v
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
        .auto-style35
        {
            text-align: right;
        }
        .auto-style36
        {
            width: 100%;
        }
        .auto-style38
        {
            width: 360px;
            font-weight: bold;
            text-align: right;
        }
        .auto-style39
        {
            width: 360px;
            font-weight: bold;
            text-align: right;
            height: 23px;
        }
        .auto-style40
        {
            height: 23px;
        }
       
        </style>
</head>
<body bgColor="#d7c7a4">
    <P style="FONT-FAMILY: arial" align="center"><SPAN style="FONT-WEIGHT: bold" class="auto-style34"><STRONG><span style="TEXT-ALIGN: center">Non-Geographic Code - </span></STRONG>View Code</SPAN></P>

	
    <form id="form1" runat="server">
				<P style="FONT-FAMILY: arial" align="center"><b>NPA:&nbsp;<asp:Label ID="lblNPA" runat="server"></asp:Label>
                    &nbsp;&nbsp; NXX:
                    <asp:Label ID="lblNXX" runat="server"></asp:Label>
                    </b></P>
            <p class="auto-style35" style="FONT-FAMILY: arial">
                <asp:Button ID="btnMenu" runat="server" Text="NonGeo Menu" Width="118px" />
&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnViewAnother" runat="server" Text="View Another" Width="115px" />
&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnNewFormA" runat="server" Text="New Form A" Visible="False" Width="118px" />
                <asp:Button ID="btnNewFormC" runat="server" Text="New Form C" Visible="False" Width="118px" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            </p>
			<P style="FONT-FAMILY: arial">
				<b>Code Details:</b> &nbsp;<asp:Label ID="lblNonGeoDetails" runat="server" Visible="False"></asp:Label>
                    </P>
                
                <asp:Panel ID="panelCodeDetails" runat="server" Height="224px" Visible="False">
                    <table class="auto-style36">
                        <tr>
                            <td class="auto-style38">
                                                                    Tix:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblTix" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    EntityID:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblEntityID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    OCN:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblOCN" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Status:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblStatus" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Public Remarks:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblPublicRemarks" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    CNA Remarks:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblCNARemarks" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Add To BIRRDS Required:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblAddToBIRRDS" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style39">
                                                                    Update BIRRDS Required:&nbsp;&nbsp;
                            </td>
                            <td class="auto-style40">
                                <asp:Label ID="lblUpdateBIRRDS" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style39">
                                                                    DBUpdateStamp:&nbsp;&nbsp;
                            </td>
                            <td class="auto-style40">
                                <asp:Label ID="lblDBUpdateStamp" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <P style="FONT-FAMILY: arial">
				    <b>Form A Details:</b>
                    <asp:Label ID="lblFormADetails" runat="server"></asp:Label>
                    </P>
                
                <asp:Panel ID="panelFormADetails" runat="server" Height="346px" Visible="False">
                     <table class="auto-style36">
                        <tr>
                            <td class="auto-style38">
                                                                    Tix:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormATix" runat="server"></asp:Label>
                            </td>
                        </tr>
                         <tr>
                             <td class="auto-style38">
                                                                      Type of Request:&nbsp;&nbsp;
                             </td>
                             <td>
                                 <asp:Label ID="lblFormATypeOfRequest" runat="server"></asp:Label>
                             </td>
                         </tr>
                         <tr>
                             <td class="auto-style38">
                                                                      Authorized Rep Name:&nbsp;&nbsp;
                             </td>
                             <td>
                                 <asp:Label ID="lblFormAAuthorizedRepName" runat="server"></asp:Label>
                             </td>
                         </tr>
                         <tr>
                             <td class="auto-style38">
                                                                      Authorized Rep Title:&nbsp;&nbsp;
                             </td>
                             <td>
                                 <asp:Label ID="lblFormAAuthorizedRepTitle" runat="server"></asp:Label>
                             </td>
                         </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    OCN:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormAOCN" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    UserID:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormAUserID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    EntityID:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormAEntityID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Form A Status:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormAStatus" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Application Date:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormAApplicationDate" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Date of Receipt:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormADateOfReceipt" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Effective Date:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormAEffectiveDate" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style39">
                                                                    Last Correspondence Date:&nbsp;&nbsp;
                            </td>
                            <td class="auto-style40">
                                <asp:Label ID="lblFormALastCorrespondenceDate" runat="server"></asp:Label>
                            </td>
                        </tr>
                         <tr>
                             <td class="auto-style39">
                                                                      Processed Date:&nbsp;&nbsp;</td>
                             <td class="auto-style40">
                                 <asp:Label ID="lblFormAProcessedDate" runat="server"></asp:Label>
                             </td>
                         </tr>
                         <tr>
                             <td class="auto-style39">DBUpdateStamp:&nbsp;&nbsp; </td>
                             <td class="auto-style40">
                                 <asp:Label ID="lblFormADBUpdateStamp" runat="server"></asp:Label>
                             </td>
                         </tr>
                    </table>
                </asp:Panel>
                <P style="FONT-FAMILY: arial">
				    <b>Form C Details:</b>
                    <asp:Label ID="lblFormCDetails" runat="server"></asp:Label>
                </P>
            
                <asp:Panel ID="panelFormCDetails" runat="server" Height="288px" Visible="False">
                    <table class="auto-style36">
                        <tr>
                            <td class="auto-style38">
                                                                    Tix:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCTix" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    UserID:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCUserID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    EntityID:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCEntityID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    In Service Date:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCInServiceDate" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Authorized Rep Name:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCAuthorizedRepName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Authorized Rep Title:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCAuthorizedRepTitle" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style38">
                                                                    Form C Date:&nbsp;&nbsp;
                            </td>
                            <td>
                                <asp:Label ID="lblFormCDate" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="auto-style39">
                                                                    DBUpdateStamp:&nbsp;&nbsp;
                            </td>
                            <td class="auto-style40">
                                <asp:Label ID="lblFormCDBUpdateStamp" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
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
