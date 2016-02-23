<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MasterPage.Master"
    CodeBehind="ErrorPage.aspx.vb" ValidateRequest="false" Inherits="TestWebApplication.ErrorPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="headContent" runat="server">
    <link type="text/css" href="../CSS/errorMsg.css" rel="Stylesheet" />
    <script type="text/javascript" language="javascript" src="Scripts/jquery-1.4.1.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            var initMsg = '';
            initMsg = '<Pleae enter the additional informaion to support team>';
            $('#<%=txtAInfo.ClientID %>').bind('focus', function () {
                $(this).removeClass('WaterMarkedTextArea')
                $(this).addClass('NormalTextBox');
                if ($(this).val() == initMsg) $(this).val('');
            }).bind("blur", function () {
                if ($(this).val() == "") {
                    $(this).removeClass("NormalTextBox");
                    $(this).addClass('WaterMarkedTextArea')
                    $(this).val(initMsg);
                }
            });
        });
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" runat="server">
    <p class="cusTitle">
        An error has occurred</p>
    <fieldset>
        <legend>Error Message</legend>
        <label>
            Application:</label><asp:TextBox ID="txtApp" ReadOnly="true" runat="server" Text="" /><br />
        <label>
            Raised By:</label><asp:TextBox ID="txtUserName" ReadOnly="true" runat="server" Text="" /><br />
        <label>
            Date Time:</label><asp:TextBox ID="txtDate" ReadOnly="true" runat="server" Text="" /><br />
        <label>
            Error Message:</label><asp:TextBox ID="txtErrMsg" ReadOnly="true" runat="server"
                Text="" /><br />
        <label for="aInfo">
            Additional Information:</label><asp:TextBox TextMode="MultiLine" ID="txtAInfo" runat="server"
                CssClass="WaterMarkedTextArea" Text="<Pleae enter the additional informaion to support team>"></asp:TextBox><br />
    </fieldset>
    &nbsp;
    <center>
        <div id="pagefooter">
            <p>
                Click Submit Ticket button to create a service request.
            </p>
            <p>
                <asp:Button ID="btnSendEmail" runat="server" Text="Submit Ticket" OnClick="btnSendEmail_OnClick" /></p>
            <label id="lblmsg" runat="server" visible="false">
                Mail sent sccessfully.</label>
        </div>
    </center>
    <asp:HiddenField ID="hdnStkTrc" Visible="false" runat="server"/>
    <asp:HiddenField ID="hdnUserName" Visible="false" runat="server" EnableViewState="true"/>
    <asp:HiddenField ID="hdnUserMail" Visible="false" runat="server" EnableViewState="true"/>
</asp:Content>
