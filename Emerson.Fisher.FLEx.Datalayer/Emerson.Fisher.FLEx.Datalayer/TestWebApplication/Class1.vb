Public Class Class1
    '    <asp:button id="showOrders" runat="server" Font-Size="8pt" Font-Names="Tahoma" Text="Show Order Information" onclick="btnShowOrders_click"></asp:button>
    '<fieldset style="margin-bottom: 0px;">
    '<legend style="color:Navy;border: solid 1px" class = "GridHeaderText">Sample to return the dataset values </legend>
    '<TABLE id="tblDatasetValues" height="1" cellspacing="0" cellPadding="0" width="100%" border="0">
    '				<TR>
    '					<TD noWrap align="left"><asp:label id="LabelTitle" runat="server" Font-Size="8pt" Font-Names="Tahoma" Font-Bold="True"
    '							ForeColor="Blue">Oracle Order information</asp:label><br />
    '                        <asp:Label ID="Label7" runat="server" Font-Bold="True" Font-Names="Tahoma" Font-Size="8pt"
    '                            ForeColor="Blue">Click on the [+] to view details of the order.</asp:Label></TD>
    '				</TR>
    '				<TR>
    '					<TD noWrap align="left">
    '                        &nbsp;<asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
    '                            CellPadding="4" Font-Names="Tahoma" Font-Size="8pt" ForeColor="#333333" GridLines="None"
    '                             OnRowCreated="GridView1_RowCreated"
    '                            OnRowDataBound="GridView1_RowDataBound">
    '                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
    '                            <Columns>
    '                                <asp:HyperLinkField Text="[+]" />
    '                                <asp:BoundField DataField="ORDER_NUMBER" HeaderText="ORDER NUMBER" />
    '                                <asp:BoundField DataField="ORDERED_DATE" HeaderText="ORDERED DATE" />
    '                                <asp:BoundField DataField="CUST_PO_NUMBER" HeaderText="CUST PO NUMBER" />
    '                                <asp:BoundField DataField="SHIPPING_METHOD_CODE" HeaderText="SHIPPING METHOD CODE" />
    '                            </Columns>
    '                            <RowStyle BackColor="#EFF3FB" />
    '                            <EditRowStyle BackColor="#2461BF" />
    '                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
    '                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
    '                            <AlternatingRowStyle BackColor="White" />
    '                        </asp:GridView>
    '                    </TD>
    '				</TR>
    '				<TR>
    '					<TD noWrap align="left">
    '						<P><asp:button id="ButtonSample" runat="server" Font-Size="8pt" Font-Names="Tahoma" Text="What happens during a postback?" onclick="ButtonSample_Click"></asp:button>
    '                        <asp:textbox id="txtExpandedDivs" runat="server" Font-Size="8pt" Font-Names="Tahoma" Width="0px"></asp:textbox></P>
    '					</TD>
    '				</TR>
    '				<TR>
    '					<TD style="HEIGHT: 2px" noWrap align="left"><asp:label id="LabelWhatHappens" runat="server" Font-Size="8pt" Font-Names="Tahoma" Font-Bold="True"
    '							Width="100%">What are we storing in the hidden textbox field (txtExpandedDivs TextBox Control)?</asp:label></TD>
    '				</TR>
    '				<TR>
    '					<TD noWrap align="left"><asp:label id="LabelPostBack" runat="server" Font-Size="8pt" Font-Names="Tahoma" Width="100%"></asp:label></TD>
    '				</TR>
    '			</TABLE>
    '            </fieldset> 
    '<br />
End Class
