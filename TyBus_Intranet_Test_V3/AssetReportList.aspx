<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="AssetReportList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.AssetReportList" %>

<asp:Content ID="AssetReportListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">固定資產報表</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="8">
                    <asp:RadioButtonList ID="eReportList" runat="server" CssClass="text-Left-Black" Width="100%" RepeatColumns="3">
                        <asp:ListItem Value="AssetP001" Selected="True">車輛折舊明細表</asp:ListItem>
                        <asp:ListItem Value="AssetP002">生財器具折舊明細表</asp:ListItem>
                        <asp:ListItem Value="AssetP003">機具折舊明細表</asp:ListItem>
                        <asp:ListItem Value="AssetP004">未攤提費用攤銷明細表</asp:ListItem>
                        <asp:ListItem Value="AssetP005">建築物折舊明細表</asp:ListItem>
                        <asp:ListItem Value="AssetP006">土地明細表</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbStopYM_Search" runat="server" CssClass="text-Right-Blue" Text="截止年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col" colspan="7">
                    <asp:TextBox ID="eStopDate_Year" runat="server" CssClass="text-Left-Black" Width="100px" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年" Width="20px" />
                    <asp:TextBox ID="eStopDate_Month" runat="server" CssClass="text-Left-Black" Width="70px" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" Width="20px" />
                    <asp:Label ID="lbNote" runat="server" CssClass="textRed" Text="(請輸入西元年月，範例：2018 年 12 月)" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="預覽資料" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="結束" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
                <td class="ColHeight_Low ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" GroupingText="查詢結果" CssClass="ShowPanel">
        <asp:Button ID="bbLayout" runat="server" CssClass="button-Black" Text="匯出 EXCEL" OnClick="bbLayout_Click" Width="100px" />
        <asp:Button ID="bbClearLayout" runat="server" CssClass="button-Red" Text="清除預覽" OnClick="bbClearLayout_Click" Width="100px" />
        <br />
        <asp:Label ID="lbReportTitle" runat="server" CssClass="errorMessageText" Width="300px" />
        <br />
        <asp:Panel ID="plAssetP001" runat="server" Visible="false" CssClass="ShowPanel-Detail">
            <asp:GridView ID="gridAssetP001" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3">
                <Columns>
                    <asp:BoundField DataField="SerialNo" HeaderText="序號" SortExpression="SerialNo" />
                    <asp:BoundField DataField="Name" HeaderText="廠牌" SortExpression="Name" />
                    <asp:BoundField DataField="Unit" HeaderText="單位" SortExpression="Unit" />
                    <asp:BoundField DataField="AssetNo" HeaderText="車號" SortExpression="AssetNo" />
                    <asp:BoundField DataField="GetDate" HeaderText="取得日期" DataFormatString="{0:yyyy/MM/dd}" />
                    <asp:BoundField DataField="Price" HeaderText="取得金額" />
                    <asp:BoundField DataField="AMT1" HeaderText="改良或修理金額" />
                    <asp:BoundField DataField="AMT2" HeaderText="本年改良或修理金額" SortExpression="AMT2" />
                    <asp:BoundField DataField="AMT3" HeaderText="合計" />
                    <asp:BoundField DataField="Durable_Years" HeaderText="折舊年限" SortExpression="Durable_Years" />
                    <asp:BoundField DataField="AMT4" HeaderText="本月提列金額" />
                    <asp:BoundField DataField="AMT5" HeaderText="本年度累計提列金額" />
                    <asp:BoundField DataField="AMT6" HeaderText="累計提列金額" />
                    <asp:BoundField DataField="AMT7" HeaderText="未折減餘額" />
                </Columns>
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                <RowStyle ForeColor="#000066" />
                <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#007DBB" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#00547E" />
            </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="plAssetP002" runat="server" Visible="false" CssClass="ShowPanel-Detail">
            <asp:GridView ID="gridAssetP002" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
                <Columns>
                    <asp:BoundField DataField="SerialNo" HeaderText="序號" SortExpression="SerialNo" />
                    <asp:BoundField DataField="AssetNo" HeaderText="編號" SortExpression="AssetNo" />
                    <asp:BoundField DataField="Name" HeaderText="名稱" SortExpression="Name" />
                    <asp:BoundField DataField="Unit" HeaderText="單位代號" SortExpression="Unit" />
                    <asp:BoundField DataField="GetDate" HeaderText="取得日期" DataFormatString="{0:yyyy/MM/dd}" />
                    <asp:BoundField DataField="Price" HeaderText="取得金額" />
                    <asp:BoundField DataField="RevalValue" HeaderText="改良或修理金額" />
                    <asp:BoundField DataField="AMT1" HeaderText="合計" />
                    <asp:BoundField DataField="Durable_Years" HeaderText="折舊年限" />
                    <asp:BoundField DataField="SalvageValue" HeaderText="保留殘值" />
                    <asp:BoundField DataField="AMT2" HeaderText="本月提列金額" />
                    <asp:BoundField DataField="AMT3" HeaderText="本年度累計提列金額" />
                    <asp:BoundField DataField="AMT4" HeaderText="累計提列金額" />
                    <asp:BoundField DataField="AMT5" HeaderText="未折減餘額" />
                </Columns>
                <FooterStyle BackColor="#C6C3C6" ForeColor="Black" />
                <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#E7E7FF" />
                <PagerStyle BackColor="#C6C3C6" ForeColor="Black" HorizontalAlign="Right" />
                <RowStyle BackColor="#DEDFDE" ForeColor="Black" />
                <SelectedRowStyle BackColor="#9471DE" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#594B9C" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#33276A" />
            </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="plAssetP003" runat="server" Visible="false" CssClass="ShowPanel-Detail">
            <asp:GridView ID="gridAssetP003" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3">
                <Columns>
                    <asp:BoundField DataField="SerialNo" HeaderText="序號" SortExpression="SerialNo" />
                    <asp:BoundField DataField="AssetNo" HeaderText="編號" SortExpression="AssetNo" />
                    <asp:BoundField DataField="Name" HeaderText="名稱" SortExpression="Name" />
                    <asp:BoundField DataField="Unit" HeaderText="單位代號" SortExpression="Unit" />
                    <asp:BoundField DataField="GetDate" HeaderText="取得日期" DataFormatString="{0:yyyy/MM/dd}" />
                    <asp:BoundField DataField="Price" HeaderText="取得金額" />
                    <asp:BoundField DataField="RevalValue" HeaderText="改良或修理金額" />
                    <asp:BoundField DataField="AMT1" HeaderText="合計" />
                    <asp:BoundField DataField="Durable_Years" HeaderText="折舊年限" />
                    <asp:BoundField DataField="SalvageValue" HeaderText="保留殘值" />
                    <asp:BoundField DataField="AMT2" HeaderText="本月提列金額" />
                    <asp:BoundField DataField="AMT3" HeaderText="本年度累計提列金額" />
                    <asp:BoundField DataField="AMT4" HeaderText="累計提列金額" />
                    <asp:BoundField DataField="AMT5" HeaderText="未折減餘額" />
                </Columns>
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                <RowStyle ForeColor="#000066" />
                <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#007DBB" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#00547E" />
            </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="plAssetP004" runat="server" Visible="false" CssClass="ShowPanel-Detail">
            <asp:GridView ID="gridAssetP004" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
                <Columns>
                    <asp:BoundField DataField="SerialNo" HeaderText="序號" SortExpression="SerialNo" />
                    <asp:BoundField DataField="AssetNo" HeaderText="編號" SortExpression="AssetNo" />
                    <asp:BoundField DataField="Name" HeaderText="名稱" SortExpression="Name" />
                    <asp:BoundField DataField="Unit" HeaderText="單位代號" SortExpression="Unit" />
                    <asp:BoundField DataField="GetDate" HeaderText="取得日期" DataFormatString="{0:yyyy/MM/dd}" />
                    <asp:BoundField DataField="Price" HeaderText="取得金額" />
                    <asp:BoundField DataField="Years" HeaderText="攤提年限" />
                    <asp:BoundField DataField="Month_Dep" HeaderText="本月攤提金額" />
                    <asp:BoundField DataField="AMT1" HeaderText="本期攤提金額" />
                    <asp:BoundField DataField="Addup_Dep" HeaderText="累計攤提金額" />
                    <asp:BoundField DataField="AMT2" HeaderText="未攤提金額" />
                    <asp:BoundField DataField="Subject2" HeaderText="科目代號" />
                    <asp:BoundField DataField="SubjectName" HeaderText="科目名稱" />
                </Columns>
                <FooterStyle BackColor="#C6C3C6" ForeColor="Black" />
                <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#E7E7FF" />
                <PagerStyle BackColor="#C6C3C6" ForeColor="Black" HorizontalAlign="Right" />
                <RowStyle BackColor="#DEDFDE" ForeColor="Black" />
                <SelectedRowStyle BackColor="#9471DE" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#594B9C" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#33276A" />
            </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="plAssetP005" runat="server" Visible="false" CssClass="ShowPanel-Detail">
            <asp:GridView ID="gridAssetP005" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3">
                <Columns>
                    <asp:BoundField DataField="SerialNo" HeaderText="序號" SortExpression="SerialNo" />
                    <asp:BoundField DataField="AssetNo" HeaderText="編號" SortExpression="AssetNo" />
                    <asp:BoundField DataField="Unit" HeaderText="單位代號" SortExpression="Unit" />
                    <asp:BoundField DataField="Name" HeaderText="構造面積 (M^2)" SortExpression="Name" />
                    <asp:BoundField DataField="Location" HeaderText="所在地" SortExpression="Location" />
                    <asp:BoundField DataField="GetDate" HeaderText="取得日期" DataFormatString="{0:yyyy/MM/dd}" />
                    <asp:BoundField DataField="Price" HeaderText="取得金額" />
                    <asp:BoundField DataField="AMT1" HeaderText="改良或修理金額" />
                    <asp:BoundField DataField="AMT2" HeaderText="本年度改良或修理金額" />
                    <asp:BoundField DataField="AMT3" HeaderText="合計" />
                    <asp:BoundField DataField="Durable_Years" HeaderText="折舊年限" />
                    <asp:BoundField DataField="AMT4" HeaderText="本月提列金額" />
                    <asp:BoundField DataField="AMT5" HeaderText="本年度累計提列金額" />
                    <asp:BoundField DataField="AMT6" HeaderText="累計提列金額" />
                    <asp:BoundField DataField="AMT7" HeaderText="未折減餘額" />
                </Columns>
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                <RowStyle ForeColor="#000066" />
                <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#007DBB" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#00547E" />
            </asp:GridView>
        </asp:Panel>
        <asp:Panel ID="plAssetP006" runat="server" Visible="false" CssClass="ShowPanel-Detail">
            <asp:GridView ID="gridAssetP006" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px" CellPadding="3" CellSpacing="1" GridLines="None">
                <Columns>
                    <asp:BoundField DataField="SerialNo" HeaderText="序號" SortExpression="SerialNo" />
                    <asp:BoundField DataField="AssetNo" HeaderText="編號" SortExpression="AssetNo" />
                    <asp:BoundField DataField="Location" HeaderText="地段地號" SortExpression="Location" />
                    <asp:BoundField DataField="Area" HeaderText="面積 (平方公尺)" SortExpression="Area" />
                    <asp:BoundField DataField="GetDate" HeaderText="取得日期" DataFormatString="{0:yyyy/MM/dd}" />
                    <asp:BoundField DataField="Price" HeaderText="取得金額" SortExpression="Price" />
                    <asp:BoundField DataField="L_N_Amount" HeaderText="重估增值" SortExpression="L_N_Amount" />
                    <asp:BoundField DataField="L_IncTax_Prepare" HeaderText="土地增值稅準備" SortExpression="L_IncTax_Prepare" />
                    <asp:BoundField DataField="AMT1" HeaderText="重估差價" />
                </Columns>
                <FooterStyle BackColor="#C6C3C6" ForeColor="Black" />
                <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#E7E7FF" />
                <PagerStyle BackColor="#C6C3C6" ForeColor="Black" HorizontalAlign="Right" />
                <RowStyle BackColor="#DEDFDE" ForeColor="Black" />
                <SelectedRowStyle BackColor="#9471DE" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#594B9C" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#33276A" />
            </asp:GridView>
        </asp:Panel>
    </asp:Panel>
</asp:Content>
