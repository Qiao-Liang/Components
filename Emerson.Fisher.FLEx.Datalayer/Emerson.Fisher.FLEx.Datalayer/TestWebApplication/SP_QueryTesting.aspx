<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SP_QueryTesting.aspx.vb" MasterPageFile="~/MasterPage.Master" Inherits="TestWebApplication.SP_QueryTesting" %>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" runat="server">
    <fieldset style="margin-bottom: 0px;">
<legend style="color:Navy;border: solid 1px"  runat="server" id="lgSPTesting" class = "labelText">Sample to Test the ExecuteDatatable method in StoredProcedure way</legend>
<table id="tblLineNotes" height="1" cellspacing="0" cellPadding="0" border="0">
<tr>
<td><asp:Label  CssClass = "labelText" ID="lblLineId" runat="server">Input Line ID </asp:Label></td><td><asp:TextBox Text = "14175768" ID = "txtLineId" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<tr>
<td><asp:Label  CssClass = "labelText" ID="lblIO5Ind" runat="server">Input IO5 Indicator </asp:Label></td><td><asp:TextBox ID = "txtIO5Indicator" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<tr>
<td><asp:Label CssClass = "labelText" ID="lblInternalLBP" runat="server">LBP/Internal Info (1-Internal, others for LBP)</asp:Label></td><td><asp:TextBox ID = "txtInternalInd" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<tr>
<td><asp:Button CssClass = "labelText" id="btnLineNotes" runat="server" Font-Size="8pt" Font-Names="Tahoma" Text="Get Line Notes" onclick="btnLineNotes_click"></asp:button> </td>
</tr>
<tr><td colspan="2"></td></tr>
<tr>
<td colspan="2"><asp:Label ID="lblLineNotes" runat="server" CssClass = "labelText"></asp:Label></td>
</tr>
<tr>
<td colspan= "2">
<asp:GridView ID="grdViewLineNotes" runat="server" AutoGenerateColumns="True"
                            CellPadding="4" Font-Names="Tahoma" Font-Size="8pt" ForeColor="#333333" GridLines="Both">
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <RowStyle BackColor="#EFF3FB" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <AlternatingRowStyle BackColor="White" />
</asp:GridView>
</td>
</tr>
</table>

</fieldset>


<fieldset style="margin-bottom: 0px;">
<legend style="color:Navy;border: solid 1px"  runat="server" id="lgSPInsert" class = "labelText">Sample to test the SP Insert</legend>
<table id="Table1" height="1" cellspacing="0" cellPadding="0" border="0">
<tr>
<td><asp:Label  CssClass = "labelText" ID="Label7" runat="server">Product Type</asp:Label></td><td><asp:TextBox Text = "" ID = "txtProdType" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<tr>
<td><asp:Label  CssClass = "labelText" ID="Label8" runat="server">Quantity </asp:Label></td><td><asp:TextBox ID = "txtQuantity" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<tr>
<td><asp:Label CssClass = "labelText" ID="Label9" runat="server">Problem Description</asp:Label></td><td><asp:TextBox ID = "txtProbDesc" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<td><asp:Label CssClass = "labelText" ID="Label11" runat="server">Problem Type</asp:Label></td><td><asp:TextBox ID = "txtProbType" runat="server" Font-Size="8pt" Font-Names="Tahoma"></asp:TextBox></td>
</tr>
<tr>
<td><asp:Button CssClass = "labelText" id="btnProbInsert" runat="server" Font-Size="8pt" Font-Names="Tahoma" Text="Insert Prob Information" onclick="btnProbInsert_click"></asp:button> </td>
</tr>
<tr>
<td colspan= "2">
<asp:GridView ID="grdProbDat" runat="server" AutoGenerateColumns="True"
                            CellPadding="4" Font-Names="Tahoma" Font-Size="8pt" ForeColor="#333333" GridLines="Both">
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <RowStyle BackColor="#EFF3FB" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <AlternatingRowStyle BackColor="White" />
</asp:GridView>
</td>
</tr>
</table>

</fieldset>

<br /><br />
<center>
    <div>
    <asp:Label ID="Label6" runat="server" Text="USING SP TESTING" Font-Bold=true Font-Size=Large  ForeColor="Blue"></asp:Label>
        
        <br />
        <asp:Panel ID="Panel1" runat="server"  HorizontalAlign="Left" Direction="LeftToRight">
         <asp:Button CssClass = "labelText" ID="Button1" runat="server" Text="Select Sing table" Width="170px" />
        <asp:GridView ID="GridView1" runat="server" HeaderStyle-ForeColor="Black" RowStyle-ForeColor="Black" Font-Names="Tahoma" Font-Size="12px" Font-Overline="false">
        </asp:GridView>
        <br />
        <br />
        <asp:Button CssClass = "labelText" ID="Button2" runat="server" Text="Select multiple tables" 
                Width="170px" />
        <asp:GridView ID="GridView2" runat="server" HeaderStyle-ForeColor="Black" RowStyle-ForeColor="Black" Font-Names="Tahoma" Font-Size="12px" Font-Overline="false">
        </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="Panel2" runat="server" HorizontalAlign="Left" Direction="LeftToRight">
        <br />
        
        <asp:Button CssClass = "labelText" ID="Button3" runat="server" Text="Insert" Width="170px" />
        <br />
        <asp:Label ID="Label1" runat="server" Text="Insert stauts:" Font-Size="14px" ForeColor="blue"></asp:Label>
        <br />
        <br />
        
    <asp:Button CssClass = "labelText" ID="Button4" runat="server" Text="Insert and return primary" 
                Width="170px" />
        <br />
        <asp:Label ID="Label2" runat="server" Text="Insert results:" Font-Size="14px" ForeColor="blue"></asp:Label>
        <br />
        <br />
       <span style="color:Blue; font-size:12px;"> LBPID:</span><asp:DropDownList ID="DropDownList1" runat="server">
        </asp:DropDownList>
        <br />
        <br />
        <asp:Button CssClass = "labelText" ID="Button5" runat="server" Text="Update" Width="170px" />
        <br />
        <asp:Label ID="Label3" runat="server" Text="Update status:" Font-Size="14px" ForeColor="blue"></asp:Label>
        <br />
        <br />
        <asp:Button CssClass = "labelText" ID="Button6" runat="server" Text="Delete" Width="170px" />
        <br />
        <asp:Label ID="Label4" runat="server" Text="Delete status:" Font-Size="14px" ForeColor="blue"></asp:Label>
        <br />
        <br />
        <br />

        <asp:Button CssClass = "labelText" ID="Button7" runat="server" Text="ExecuteScalar" Width="170px" />
        <br />
    <asp:Label ID="Label5" runat="server" Text="test ExecuteScalar:" Font-Size="14px" ForeColor="blue"></asp:Label>
        <br />
        <br />
        </asp:Panel>
    </div>
<%--<asp:hyperlink runat="server" NavigateUrl="~/Default.aspx">back</asp:hyperlink>--%>

</center>
</asp:Content>
