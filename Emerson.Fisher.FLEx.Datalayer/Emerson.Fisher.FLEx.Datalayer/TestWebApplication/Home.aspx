<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Home.aspx.vb" MasterPageFile="~/MasterPage.Master"
    Inherits="TestWebApplication._Default" %>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" runat="server">
    <center>
        <asp:Panel ID="Panel1" runat="server">
            <br />
            <br />
            <asp:HyperLink ID="lnkInlineQueryTesting" Visible="false" runat="server" NavigateUrl="~/Inline_QueryTesting.aspx">Inline query testing</asp:HyperLink>
            <br />
            <br />
            <asp:HyperLink ID="lnkSPTesting" runat="server" Visible="false" NavigateUrl="~/SP_QueryTesting.aspx">SP testing</asp:HyperLink>
            <br />
            <br />
            <asp:Button ID ="btnCauseExc" runat="server" Text="Throw Exception" /><br />
            <br />
            <asp:HyperLink Visible="false" ID="lnkExceptionHandlingTesting" runat="server" NavigateUrl="~/EPMLogging.aspx">SP testing</asp:HyperLink>
        </asp:Panel>
    </center>
</asp:Content>
