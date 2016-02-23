<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Inline_QueryTesting.aspx.vb" MasterPageFile="~/MasterPage.Master" Inherits="TestWebApplication.Inline_QueryTesting" %>
 
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" runat="server">
<br />
<fieldset style="margin-bottom: 0px;">
<legend style="color:Navy;border: solid 1px" runat="server" id="lgInvOrg" class = "labelText">Sample to return the Datatable values </legend>
             <div class="grid">
                <div class="rounded">
                    <div class="top-outer">
                        <div class="top-inner">
                            <div class="top">
                                <h2 class = "GridHeaderText">Inventory Organization</h2>
                            </div>
                        </div>
                    </div>
                    <div class="mid-outer">
                        <div class="mid-inner">
                            <div class="mid">
                                <asp:UpdatePanel ID="updPanel" runat="server" UpdateMode="Conditional">
                                    <ContentTemplate>
                                     <asp:ListView ID="lvInventoryOrgs" runat="server">
                                        <LayoutTemplate>
                                        <table id="products" runat="server" class="datatable" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <th class="first"><asp:Label ID="lblORGANIZATION_CODE" runat="server" Text="ORGANIZATION CODE" /></th>
                                            <th><asp:Label ID="lblORGANIZATION_ID" runat="server" Text="Org ID" /></th>
                                            <th><asp:Label ID="lblOPERATING_UNIT" runat="server" Text="OPERATING UNIT" /></th>
                                            <th><asp:Label ID="lblFUNCTIONINGCURRENCYCODE" runat="server" Text="Currency Code" /></th>
                                            <th><asp:Label ID="lblSTANDARD_GMT_DEVIATION_HOURS" runat="server" Text="STANDARD GMT DEVIATION HOURS" /></th>
                                            <th><asp:Label ID="lblDAYLIGHT_SAVINGS_TIME_FLAG" runat="server" Text="DAYLIGHT SAVINGS TIME FLAG" /></th>
                                        </tr>
                                        <tr id="itemPlaceholder" runat="server" />
                                        </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr id="item" runat="server" class="row">
                                                <td class="first"><%# Eval("ORGANIZATION_CODE")%></td>
                                                <td><%# Eval("ORGANIZATION_ID")%></td>
                                                <td><%# Eval("OPERATING_UNIT")%></td>
                                                <td><%# Eval("FUNCTIONINGCURRENCYCODE")%></td>
                                                <td><%# Eval("STANDARD_GMT_DEVIATION_HOURS")%></td>
                                                <td><%# Eval("DAYLIGHT_SAVINGS_TIME_FLAG")%></td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:ListView>
                                    </ContentTemplate>
                                </asp:UpdatePanel> 
                            </div>
                        </div> 
                    </div>
                    <div class="bottom-outer"><div class="bottom-inner"><div class="bottom"></div></div></div>
                </div> 
            </div>
            </fieldset>
<br /><br />
<fieldset style="margin-bottom: 0px;">
<legend style="color:Navy;border: solid 1px" runat="server" id="lgExecuteReader" class = "labelText">Sample to Test the execute reader method </legend>
<table id="tblOrderPOInfo" height="1" cellspacing="0" cellPadding="0" border="0">
<tr>
<td><asp:TextBox ID = "txtOrderNumber" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
<td><asp:button id="btnGetPO" runat="server" Font-Size="8pt" Font-Names="Tahoma" Text="Get PO Number" onclick="btnGetPO_click"></asp:button> </td>
</tr>
<tr><td colspan="2"></td></tr>
<tr>
<td colspan="2"><asp:Label ID="lblPONumber" runat="server" Font-Size="8pt" ForeColor = "Black" Font-Names="Tahoma"></asp:Label></td>
</tr>
</table>
</fieldset>
<br />
<%--class = "GridHeaderText"--%>
<fieldset style="margin-bottom: 0px;">
<legend style="color:Navy;border: solid 1px" runat="server" id="lgNonQuery" class = "labelText" >Sample to test Execute Nonquery method</legend>
<div class="grid">
                <div class="rounded">
                    <div class="top-outer">
                        <div class="top-inner">
                            <div class="top">
                                <h2 class = "GridHeaderText">Order Information from local DB.</h2>
                            </div>
                        </div>
                    </div>
                    <div class="mid-outer">
                        <div class="mid-inner">
                            <div class="mid">
                                <asp:UpdatePanel ID="updPnlLocalOrderInfo" runat="server" UpdateMode="Conditional">
                                    <ContentTemplate>
                                     <asp:ListView ID="lvLocalOrderInfo" runat="server">
                                        <LayoutTemplate>
                                        <table id="OrderInfo" runat="server" class="datatable" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <th class="first"><asp:Label ID="ORDER_NUMBER" runat="server" Text="ORDER NUMBER" /></th>
                                            <th><asp:Label ID="CUST_PO_NUMBER" runat="server" Text="CUST PO NUMBER" /></th>
                                            <th><asp:Label ID="SHIPPING_METHOD_CODE" runat="server" Text="SHIPPING METHOD CODE" /></th>
                                            <th><asp:Label ID="PACKING_INSTRUCTIONS" runat="server" Text="PACKING INSTRUCTIONS" /></th>
                                        </tr>
                                        <tr id="itemPlaceholder" runat="server" />
                                        </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr id="item" runat="server" class="row">
                                                <td class="first"><%# Eval("ORDER_NUMBER")%></td>
                                                <td><%# Eval("CUST_PO_NUMBER")%></td>
                                                <td><%# Eval("SHIPPING_METHOD_CODE")%></td>
                                                <td><%# Eval("PACKING_INSTRUCTIONS")%></td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:ListView>
                                    </ContentTemplate>
                                </asp:UpdatePanel> 
                            </div>
                        </div> 
                    </div>
                    <div class="bottom-outer"><div class="bottom-inner"><div class="bottom"></div></div></div>
                </div> 
            </div>
<br />
<table id="tblOrderDelete" height="1" cellspacing="0" cellPadding="0" border="0">
<tr>
<td><asp:TextBox ID = "txtOrderNoToDelete" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
<td><asp:button id="btnDeleteOrder" runat="server" Font-Size="8pt" Font-Names="Tahoma" Text="Delete Order Number" onclick="btnDeleteOrder_click"></asp:button> </td>
</tr>
<tr><td colspan="2"></td></tr>
<tr>
<td colspan="2"><asp:Label ID="lblDeleteStatus" runat="server" Font-Size="8pt" ForeColor="Black" Font-Names="Tahoma"></asp:Label></td>
</tr>
</table>
</fieldset>
</asp:Content>
