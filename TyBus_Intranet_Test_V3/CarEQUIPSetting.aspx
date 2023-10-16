<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarEQUIPSetting.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarEQUIPSetting" %>

<asp:Content ID="CarEQUIPSettingForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">ERP車輛固定配備名稱維護</a>
    </div>
    <br />
    <asp:Panel ID="plData" runat="server" CssClass="ShowPanel" Width="100%">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                    <asp:Label ID="lbTitleFix" runat="server" CssClass="titleText-S-Blue" Text="固定裝備" Width="100%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK01" runat="server" CssClass="text-Right-Blue" Text="固定配備 01" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK01" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK1" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK02" runat="server" CssClass="text-Right-Blue" Text="固定配備 02" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK02" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK2" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK03" runat="server" CssClass="text-Right-Blue" Text="固定配備 03" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK03" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK3" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK04" runat="server" CssClass="text-Right-Blue" Text="固定配備 04" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK04" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK4" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK05" runat="server" CssClass="text-Right-Blue" Text="固定配備 05" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK05" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK5" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK06" runat="server" CssClass="text-Right-Blue" Text="固定配備 06" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK06" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK6" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK07" runat="server" CssClass="text-Right-Blue" Text="固定配備 07" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK07" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK7" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK08" runat="server" CssClass="text-Right-Blue" Text="固定配備 08" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK08" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK8" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK09" runat="server" CssClass="text-Right-Blue" Text="固定配備 09" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK09" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK9" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK10" runat="server" CssClass="text-Right-Blue" Text="固定配備 10" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK10" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK10" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK11" runat="server" CssClass="text-Right-Blue" Text="固定配備 11" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK11" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK11" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK12" runat="server" CssClass="text-Right-Blue" Text="固定配備 12" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK12" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK12" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK13" runat="server" CssClass="text-Right-Blue" Text="固定配備 13" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK13" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK13" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK14" runat="server" CssClass="text-Right-Blue" Text="固定配備 14" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK14" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK14" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK15" runat="server" CssClass="text-Right-Blue" Text="固定配備 15" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK15" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK15" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK16" runat="server" CssClass="text-Right-Blue" Text="固定配備 16" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK16" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK16" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK17" runat="server" CssClass="text-Right-Blue" Text="固定配備 17" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK17" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK17" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK18" runat="server" CssClass="text-Right-Blue" Text="固定配備 18" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK18" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK18" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK19" runat="server" CssClass="text-Right-Blue" Text="固定配備 19" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK19" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK19" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK20" runat="server" CssClass="text-Right-Blue" Text="固定配備 20" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK20" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK20" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK21" runat="server" CssClass="text-Right-Blue" Text="固定配備 21" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK21" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK21" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK22" runat="server" CssClass="text-Right-Blue" Text="固定配備 22" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK22" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK22" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK23" runat="server" CssClass="text-Right-Blue" Text="固定配備 23" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK23" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK23" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK24" runat="server" CssClass="text-Right-Blue" Text="固定配備 24" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK24" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK24" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK25" runat="server" CssClass="text-Right-Blue" Text="固定配備 25" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK25" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK25" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK26" runat="server" CssClass="text-Right-Blue" Text="固定配備 26" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK26" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK26" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK27" runat="server" CssClass="text-Right-Blue" Text="固定配備 27" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK27" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK27" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK28" runat="server" CssClass="text-Right-Blue" Text="固定配備 28" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK28" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK28" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK29" runat="server" CssClass="text-Right-Blue" Text="固定配備 29" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK29" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK29" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK30" runat="server" CssClass="text-Right-Blue" Text="固定配備 30" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK30" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK30" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK31" runat="server" CssClass="text-Right-Blue" Text="固定配備 31" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK31" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK31" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK32" runat="server" CssClass="text-Right-Blue" Text="固定配備 32" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK32" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK32" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK33" runat="server" CssClass="text-Right-Blue" Text="固定配備 33" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK33" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK33" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK34" runat="server" CssClass="text-Right-Blue" Text="固定配備 34" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK34" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK34" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK35" runat="server" CssClass="text-Right-Blue" Text="固定配備 35" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK35" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK35" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK36" runat="server" CssClass="text-Right-Blue" Text="固定配備 36" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID ="eCK36" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK36" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK37" runat="server" CssClass="text-Right-Blue" Text="固定配備 37" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK37" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK37" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK38" runat="server" CssClass="text-Right-Blue" Text="固定配備 38" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK38" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK38" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK39" runat="server" CssClass="text-Right-Blue" Text="固定配備 39" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK39" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK39" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbCK40" runat="server" CssClass="text-Right-Blue" Text="固定配備 40" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eCK40" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_CK40" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                    <asp:Label ID="lbTitleMoveable" runat="server" CssClass="titleText-S-Black" Text="選配裝備一" Width="100%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB41" runat="server" CssClass="text-Right-Blue" Text="選配裝備 01" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB41" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb41" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB42" runat="server" CssClass="text-Right-Blue" Text="選配裝備 02" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB42" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb42" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB43" runat="server" CssClass="text-Right-Blue" Text="選配裝備 03" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB43" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb43" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB44" runat="server" CssClass="text-Right-Blue" Text="選配裝備 04" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB44" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb44" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB45" runat="server" CssClass="text-Right-Blue" Text="選配裝備 05" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB45" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb45" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB46" runat="server" CssClass="text-Right-Blue" Text="選配裝備 06" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB46" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb46" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB47" runat="server" CssClass="text-Right-Blue" Text="選配裝備 07" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB47" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb47" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB48" runat="server" CssClass="text-Right-Blue" Text="選配裝備 08" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB48" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb48" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB49" runat="server" CssClass="text-Right-Blue" Text="選配裝備 09" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB49" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb49" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB50" runat="server" CssClass="text-Right-Blue" Text="選配裝備 10" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB50" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb50" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB51" runat="server" CssClass="text-Right-Blue" Text="選配裝備 11" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB51" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb51" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB52" runat="server" CssClass="text-Right-Blue" Text="選配裝備 12" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB52" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb52" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB53" runat="server" CssClass="text-Right-Blue" Text="選配裝備 13" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB53" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb53" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB54" runat="server" CssClass="text-Right-Blue" Text="選配裝備 14" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB54" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb54" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB55" runat="server" CssClass="text-Right-Blue" Text="選配裝備 15" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB55" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb55" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col" colspan="10">
                    <asp:Label ID="lbTitleMoveable2" runat="server" CssClass="titleText-S-Red" Text="選配裝備二" Width="100%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB56" runat="server" CssClass="text-Right-Blue" Text="選配裝備 16" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB56" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb56" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB57" runat="server" CssClass="text-Right-Blue" Text="選配裝備 17" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB57" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb57" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB58" runat="server" CssClass="text-Right-Blue" Text="選配裝備 18" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB58" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb58" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB59" runat="server" CssClass="text-Right-Blue" Text="選配裝備 19" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB59" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb59" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB60" runat="server" CssClass="text-Right-Blue" Text="選配裝備 20" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB60" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb60" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB61" runat="server" CssClass="text-Right-Blue" Text="選配裝備 21" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB61" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb61" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB62" runat="server" CssClass="text-Right-Blue" Text="選配裝備 22" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB62" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb62" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB63" runat="server" CssClass="text-Right-Blue" Text="選配裝備 23" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB63" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb63" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB64" runat="server" CssClass="text-Right-Blue" Text="選配裝備 24" Width="100%" />
                </td>
                <td class="coolh ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB64" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb64" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB65" runat="server" CssClass="text-Right-Blue" Text="選配裝備 25" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB65" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb65" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB66" runat="server" CssClass="text-Right-Blue" Text="選配裝備 26" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB66" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb66" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB67" runat="server" CssClass="text-Right-Blue" Text="選配裝備 27" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB67" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb67" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB68" runat="server" CssClass="text-Right-Blue" Text="選配裝備 28" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB68" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb68" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB69" runat="server" CssClass="text-Right-Blue" Text="選配裝備 29" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB69" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb69" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Label ID="lbLB70" runat="server" CssClass="text-Right-Blue" Text="選配裝備 30" Width="100%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:TextBox ID="eLB70" runat="server" CssClass="text-Left-Black" Width="95%" ToolTip="a_lb70" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Button ID="bbUpdate" runat="server" CssClass="button-Black" Text="更新" OnClick="bbUpdate_Click" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Button ID="bbCancel" runat="server" CssClass="button-Blue" Text="取消" OnClick="bbCancel_Click" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-10Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" Text="離開" OnClick="bbClose_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
                <td class="ColHeight ColWidth-10Col" />
            </tr>
        </table>
    </asp:Panel>
</asp:Content>
