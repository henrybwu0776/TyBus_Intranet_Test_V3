<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="TyBus_Intranet_Test_V3.Index" %>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <br />
        <br />
        <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
            <a class="titleText-Red">最新公告</a>
            <br />
            <asp:ListView ID="ListAnnList" runat="server" DataSourceID="sdsAnnList_List">
                <AlternatingItemTemplate>
                    <li style="background-color: #FFE4E1;">公告主旨:
                        <asp:Label ID="公告主旨Label" runat="server" Text='<%# Eval("公告主旨") %>' />
                        <br />
                        公告本文:
                        <asp:Label ID="公告本文Label" runat="server" Text='<%# Eval("公告本文") %>' />
                        <br />
                        附檔檔名:
                        <asp:Label ID="附檔檔名Label" runat="server" Text='<%# Eval("附檔檔名") %>' />
                        <br />
                    </li>
                </AlternatingItemTemplate>
                <EditItemTemplate>
                    <li style="background-color: #008A8C; color: #FFFFFF;">公告主旨:
                        <asp:TextBox ID="公告主旨TextBox" runat="server" Text='<%# Bind("公告主旨") %>' />
                        <br />
                        公告本文:
                        <asp:TextBox ID="公告本文TextBox" runat="server" Text='<%# Bind("公告本文") %>' />
                        <br />
                        附檔檔名:
                        <asp:TextBox ID="附檔檔名TextBox" runat="server" Text='<%# Bind("附檔檔名") %>'/>
                        <br />
                        <asp:Button ID="UpdateButton" runat="server" CommandName="Update" Text="更新" />
                        <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" Text="取消" />
                    </li>
                </EditItemTemplate>
                <EmptyDataTemplate>
                    <img alt="" src="img/TYBUS_LOGO.png" />
                </EmptyDataTemplate>
                <InsertItemTemplate>
                    <li style="">公告主旨:
                        <asp:TextBox ID="公告主旨TextBox" runat="server" Text='<%# Bind("公告主旨") %>' />
                        <br />
                        公告本文:
                        <asp:TextBox ID="公告本文TextBox" runat="server" Text='<%# Bind("公告本文") %>' />
                        <br />
                        附檔檔名:
                        <asp:TextBox ID="附檔檔名TextBox" runat="server" Text='<%# Bind("附檔檔名") %>' />
                        <br />
                        <asp:Button ID="InsertButton" runat="server" CommandName="Insert" Text="插入" />
                        <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" Text="清除" />
                    </li>
                </InsertItemTemplate>
                <ItemSeparatorTemplate>
                    <br />
                </ItemSeparatorTemplate>
                <ItemTemplate>
                    <li style="background-color: #DCDCDC; color: #000000;">公告主旨:
                        <asp:Label ID="公告主旨Label" runat="server" Text='<%# Eval("公告主旨") %>' />
                        <br />
                        公告本文:
                        <asp:Label ID="公告本文Label" runat="server" Text='<%# Eval("公告本文") %>' />
                        <br />
                        附檔檔名:
                        <asp:Label ID="附檔檔名Label" runat="server" Text='<%# Eval("附檔檔名") %>' />
                        <br />
                    </li>
                </ItemTemplate>
                <LayoutTemplate>
                    <ul id="itemPlaceholderContainer" runat="server" style="font-family: Verdana, Arial, Helvetica, sans-serif;">
                        <li runat="server" id="itemPlaceholder" />
                    </ul>
                    <div style="text-align: center; background-color: #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; color: #000000;">
                    </div>
                </LayoutTemplate>
                <SelectedItemTemplate>
                    <li style="background-color: #008A8C; font-weight: bold; color: #FFFFFF;">公告主旨:
                        <asp:Label ID="公告主旨Label" runat="server" Text='<%# Eval("公告主旨") %>' />
                        <br />
                        公告本文:
                        <asp:Label ID="公告本文Label" runat="server" Text='<%# Eval("公告本文") %>' />
                        <br />
                        附檔檔名:
                        <asp:Label ID="附檔檔名Label" runat="server" Text='<%# Eval("附檔檔名") %>' />
                        <br />
                    </li>
                </SelectedItemTemplate>
            </asp:ListView>
            <asp:SqlDataSource ID="sdsAnnList_List" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>" SelectCommand="SELECT PostTitle AS 公告主旨, Remark AS 公告本文, PostFiles AS 附檔檔名, StartDate, EndDate FROM AnnList WHERE (ISNULL(PostOpen, N'X') = 'V') AND (CONVERT (nvarchar(10), StartDate, 111) &lt;= CONVERT (nvarchar(10), GETDATE(), 111)) AND (CONVERT (nvarchar(10), EndDate, 111) &gt;= CONVERT (nvarchar(10), GETDATE(), 111)) OR (ISNULL(PostOpen, N'X') = 'V') AND (ISNULL(EndDate, '') = '') ORDER BY AnnNo DESC"></asp:SqlDataSource>
        </asp:Panel>
    </div>
</asp:Content>
