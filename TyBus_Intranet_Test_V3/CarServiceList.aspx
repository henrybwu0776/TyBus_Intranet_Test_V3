<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="CarServiceList.aspx.cs" Inherits="TyBus_Intranet_Test_V3.CarServiceList" %>

<asp:Content ID="CarServiceListForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">車輛保養預排</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbDepNo_Search" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:DropDownList ID="ddlDepNo_Search" runat="server" CssClass="text-Left-Black" Width="90%"
                        DataSourceID="sdsDepList" DataTextField="DepName" DataValueField="DepNo" AutoPostBack="true"
                        OnSelectedIndexChanged="ddlDepNo_Search_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:Label ID="eDepNo_Search" runat="server" Visible="false" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbCalYM_Search" runat="server" CssClass="text-Right-Blue" Text="計算年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eCalYear_Search" runat="server" CssClass="text-Left-Black" Width="30%" AutoPostBack="True" />
                    <asp:Label ID="lbSplit_1" runat="server" CssClass="text-Left-Black" Text=" 年 " />
                    <asp:TextBox ID="eCalMonth_Search" runat="server" CssClass="text-Left-Black" Width="25%" AutoPostBack="True" />
                    <asp:Label ID="lbSplit_2" runat="server" CssClass="text-Left-Black" Text=" 月" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbSecondServiceKM_Search" runat="server" CssClass="text-Right-Blue" Text="二級保養間隔：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eSecondServiceKM_Search" runat="server" CssClass="text-Left-Black" Width="70%" />
                    <asp:Label ID="lbSplit_3" runat="server" CssClass="text-Left-Black" Text=" KM" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:Label ID="lbThirdServiceKM_Search" runat="server" CssClass="text-Right-Blue" Text="三級保養間隔：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-8Col">
                    <asp:TextBox ID="eThirdServiceKM_Search" runat="server" CssClass="text-Left-Black" Width="70%" />
                    <asp:Label ID="lbSplit_4" runat="server" CssClass="text-Left-Black" Text=" KM" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbExportData" runat="server" CssClass="button-Black" OnClick="bbExportData_Click" Text="匯出預排資料" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbGetList" runat="server" CssClass="button-Black" OnClick="bbGetList_Click" Text="計算預排" Width="90%" UseSubmitBehavior="False" OnClientClick="this.disabled=true;" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbStopCalYM" runat="server" CssClass="button-Red" OnClick="bbStopCalYM_Click" Text="預排關帳" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbReOpenYM" runat="server" CssClass="button-Red" OnClick="bbReOpenYM_Click" Text="解鎖開放修改" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-8Col">
                    <asp:Button ID="bbClose" runat="server" CssClass="button-Red" OnClick="bbClose_Click" Text="結束" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
                <td class="ColWidth-8Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" CssClass="ShowPanel">
        <asp:FormView ID="fvServicePreDetail" runat="server" Width="100%" DataKeyNames="ServicePreNo" DataSourceID="sdsShowDetail" OnDataBound="fvServicePreDetail_DataBound">
            <EmptyDataTemplate>
                <asp:Button ID="bbInsert_Empty" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
            </EmptyDataTemplate>
            <EditItemTemplate>
                <asp:Button ID="bbOK_Edit" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_Edit_Click" Text="確定" />
                &nbsp;<asp:Button ID="bbCancel_Edit" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" />
                <asp:UpdatePanel ID="upDataEdit" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServicePreNo_Edit" runat="server" CssClass="text-Right-Blue" Text="預排單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eServicePreNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServicePreNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eOriDate_Edit" runat="server" Text='<%# Eval("OriDate","{0:yyyy/MM/dd}") %>' Visible="false" />
                                    <asp:Label ID="eOldDate_Edit" runat="server" Text='<%# Eval("OldDate","{0:yyyy/MM/dd}") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_ID_Edit" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eCar_ID_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_Edit" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eDepNo_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="25%" />
                                    <asp:Label ID="eDepName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_Edit" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Edit_TextChanged" Text='<%# Bind("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceType_Edit" runat="server" CssClass="text-Right-Blue" Text="保養級別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlServiceType_Edit" runat="server" CssClass="text-Left-Black" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlServiceType_Edit_SelectedIndexChanged" Width="90%"
                                        DataSourceID="sdsServiceType_Edit" DataTextField="CLASSTXT" DataValueField="CLASSNO">
                                    </asp:DropDownList>
                                    <asp:Label ID="eServiceType_Edit" runat="server" Text='<%# Bind("ServiceType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLastDate_Edit" runat="server" CssClass="text-Right-Blue" Text="上次保養日：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLastDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceDate_Edit" runat="server" CssClass="text-Right-Blue" Text="預排保養日：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eServiceDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_Edit" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                                    <asp:TextBox ID="eRemark_Edit" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="98%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_Edit" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_Edit" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_Edit" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_Edit" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </EditItemTemplate>
            <InsertItemTemplate>
                <asp:Button ID="bbOK_INS" runat="server" CausesValidation="True" CssClass="button-Black" OnClick="bbOK_INS_Click" Text="確定" Width="120px" />
                &nbsp;<asp:Button ID="bbCancel_INS" runat="server" CausesValidation="False" CssClass="button-Red" CommandName="Cancel" Text="取消" Width="120px" />
                <asp:UpdatePanel ID="upDataInsert" runat="server">
                    <ContentTemplate>
                        <table class="TableSetting">
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServicePreNo_INS" runat="server" CssClass="text-Right-Blue" Text="預排單號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eServicePreNo_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServicePreNo") %>' Width="90%" />
                                </td>
                                <td class="ColHeight ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eOriDate_INS" runat="server" Text='<%# Eval("OriDate","{0:yyyy/MM/dd}") %>' Visible="false" />
                                    <asp:Label ID="eOldDate_INS" runat="server" Text='<%# Eval("OldDate","{0:yyyy/MM/dd}") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbCar_ID_INS" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eCar_ID_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eCar_ID_INS_TextChanged" Text='<%# Eval("Car_ID") %>' Width="90%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDepNo_INS" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDepNo_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDepNo_INS_TextChanged" Text='<%# Eval("DepNo") %>' Width="25%" />
                                    <asp:Label ID="eDepName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbDriver_INS" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:TextBox ID="eDriver_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_INS_TextChanged" Text='<%# Eval("Driver") %>' Width="35%" />
                                    <asp:Label ID="eDriverName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceType_INS" runat="server" CssClass="text-Right-Blue" Text="保養級別：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:DropDownList ID="ddlServiceType_INS" runat="server" CssClass="text-Left-Black" AutoPostBack="True" OnSelectedIndexChanged="ddlServiceType_INS_SelectedIndexChanged" Width="90%" DataSourceID="sdsServiceType_INS" DataTextField="CLASSTXT" DataValueField="CLASSNO" />
                                    <asp:Label ID="eServiceType_INS" runat="server" Text='<%# Eval("ServiceType") %>' Visible="false" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbLastDate_INS" runat="server" CssClass="text-Right-Blue" Text="上次保養日：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eLastDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbServiceDate_INS" runat="server" CssClass="text-Right-Blue" Text="預排保養日：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:TextBox ID="eServiceDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbRemark_INS" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                                </td>
                                <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                                    <asp:TextBox ID="eRemark_INS" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Text='<%# Eval("Remark") %>' Width="98%" Height="95%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuDate_INS" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eBuDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbBuMan_INS" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eBuMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                                    <asp:Label ID="eBuManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColHeight ColWidth-8Col" />
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyDate_INS" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="eModifyDate_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col">
                                    <asp:Label ID="lbModifyMan_INS" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                                </td>
                                <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                                    <asp:Label ID="eModifyMan_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                                    <asp:Label ID="eModifyManName_INS" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                                </td>
                            </tr>
                            <tr>
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                                <td class="ColWidth-8Col" />
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </InsertItemTemplate>
            <ItemTemplate>
                <asp:Button ID="bbInsert" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="New" Text="新增" Width="120px" />
                &nbsp;<asp:Button ID="bbEdit" runat="server" CausesValidation="False" CssClass="button-Black" CommandName="Edit" Text="編輯" Width="120px" />
                &nbsp;<asp:Button ID="bbDelete" runat="server" CausesValidation="False" CssClass="button-Red" OnClick="bbDelete_Click" Text="刪除" Width="120px" />
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServicePreNo_List" runat="server" CssClass="text-Right-Blue" Text="預排單號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServicePreNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServicePreNo") %>' Width="90%" />
                        </td>
                        <td class="ColHeight ColWidth-8Col" colspan="2">
                            <asp:Label ID="eOriDate_List" runat="server" Text='<%# Eval("OriDate","{0:yyyy/MM/dd}") %>' Visible="false" />
                            <asp:Label ID="eOldDate_List" runat="server" Text='<%# Eval("OldDate","{0:yyyy/MM/dd}") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbCar_ID_List" runat="server" CssClass="text-Right-Blue" Text="車號：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eCar_ID_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Car_ID") %>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="站別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepNo") %>' Width="25%" />
                            <asp:Label ID="eDepName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DepName") %>' Width="65%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbDriver_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eDriver_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("Driver") %>' Width="35%" />
                            <asp:Label ID="eDriverName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("DriverName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceType_List" runat="server" CssClass="text-Right-Blue" Text="保養級別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceType_C_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceType_C") %>' Width="90%" />
                            <asp:Label ID="eServiceType_List" runat="server" Text='<%# Eval("ServiceType") %>' Visible="false" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbLastDate_List" runat="server" CssClass="text-Right-Blue" Text="上次保養日：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eLastDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("LastDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbServiceDate_List" runat="server" CssClass="text-Right-Blue" Text="預排保養日：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eServiceDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ServiceDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col">
                            <asp:Label ID="lbRemark_List" runat="server" CssClass="text-Right-Blue" Text="備註：" Width="90%" />
                        </td>
                        <td class="MultiLine-Low ColBorder ColWidth-8Col" colspan="5">
                            <asp:TextBox ID="eRemark_List" runat="server" CssClass="text-Left-Black" TextMode="MultiLine" Enabled="false" Text='<%# Eval("Remark") %>' Width="98%" Height="95%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="建檔日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbBuMan_List" runat="server" CssClass="text-Right-Blue" Text="建檔人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eBuMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuMan") %>' Width="35%" />
                            <asp:Label ID="eBuManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("BuManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColWidth-8Col" />
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyDate_List" runat="server" CssClass="text-Right-Blue" Text="異動日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="eModifyDate_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyDate","{0:yyyy/MM/dd}") %>' wid="95%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col">
                            <asp:Label ID="lbModifyMan_List" runat="server" CssClass="text-Right-Blue" Text="異動人員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-8Col" colspan="2">
                            <asp:Label ID="eModifyMan_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyMan") %>' Width="35%" />
                            <asp:Label ID="eModifyManName_List" runat="server" CssClass="text-Left-Black" Text='<%# Eval("ModifyManName") %>' Width="55%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                        <td class="ColWidth-8Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
        <asp:GridView ID="gridDailyList" runat="server" Width="100%" AutoGenerateColumns="False" BackColor="White" BorderColor="White" BorderStyle="Ridge" BorderWidth="2px"
            CellPadding="3" CellSpacing="1" DataKeyNames="ServicePreNo" DataSourceID="sdsShowList" GridLines="None">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
                <asp:BoundField DataField="ServicePreNo" HeaderText="ServicePreNo" ReadOnly="True" SortExpression="ServicePreNo" Visible="False" />
                <asp:BoundField DataField="DepNo" HeaderText="DepNo" SortExpression="DepNo" Visible="False" />
                <asp:BoundField DataField="DepName" HeaderText="站別" ReadOnly="True" SortExpression="DepName" />
                <asp:BoundField DataField="Car_ID" HeaderText="車號" SortExpression="Car_ID" />
                <asp:BoundField DataField="Driver" HeaderText="Driver" SortExpression="Driver" Visible="False" />
                <asp:BoundField DataField="DriverName" HeaderText="駕駛員" ReadOnly="True" SortExpression="DriverName" />
                <asp:BoundField DataField="ServiceType" HeaderText="ServiceType" SortExpression="ServiceType" Visible="False" />
                <asp:BoundField DataField="ServiceType_C" HeaderText="保養級別" ReadOnly="True" SortExpression="ServiceType_C" />
                <asp:BoundField DataField="LastDate" DataFormatString="{0:D}" HeaderText="上次保養日" SortExpression="LastDate" />
                <asp:BoundField DataField="ServiceDate" DataFormatString="{0:D}" HeaderText="預排保養日" SortExpression="ServiceDate" />
                <asp:BoundField DataField="Remark" HeaderText="備註" SortExpression="Remark" />
                <asp:BoundField DataField="BuMan" HeaderText="BuMan" SortExpression="BuMan" Visible="False" />
                <asp:BoundField DataField="BuManName" HeaderText="BuManName" ReadOnly="True" SortExpression="BuManName" Visible="False" />
                <asp:BoundField DataField="BuDate" HeaderText="BuDate" SortExpression="BuDate" Visible="False" />
                <asp:BoundField DataField="ModifyMan" HeaderText="ModifyMan" SortExpression="ModifyMan" Visible="False" />
                <asp:BoundField DataField="ModifyManName" HeaderText="ModifyManName" ReadOnly="True" SortExpression="ModifyManName" Visible="False" />
                <asp:BoundField DataField="ModifyDate" HeaderText="ModifyDate" SortExpression="ModifyDate" Visible="False" />
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
    <asp:Panel ID="plCalender" runat="server" Width="100%">
        <asp:Calendar ID="calCarServiceList" runat="server" Width="100%"
            OnDayRender="calCarServiceList_DayRender"
            OnSelectionChanged="calCarServiceList_SelectionChanged" BackColor="#CCFFFF">
            <DayStyle Font-Bold="True" Font-Size="14pt" />
            <TitleStyle Font-Bold="True" Font-Size="20pt" ForeColor="#FFFF66" BackColor="#6666FF" />
        </asp:Calendar>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsShowList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ServicePreNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = c.DepNo)) AS DepName, Car_ID, Driver, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = c.Driver)) AS DriverName, ServiceType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.ServiceType) AND (FKEY = '工作單A         FixworkA        SERVICE')) AS ServiceType_C, LastDate, ServiceDate, Remark, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = c.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = c.ModifyMan)) AS ModifyManName, ModifyDate FROM CarServicePre AS c WHERE (1 &lt;&gt; 1)"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsShowDetail" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT ServicePreNo, DepNo, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = c.DepNo)) AS DepName, Car_ID, Driver, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = c.Driver)) AS DriverName, ServiceType, (SELECT CLASSTXT FROM DBDICB WHERE (CLASSNO = c.ServiceType) AND (FKEY = '工作單A         FixworkA        SERVICE')) AS ServiceType_C, LastDate, ServiceDate, ServiceDate AS OldDate, Remark, BuMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_2 WHERE (EMPNO = c.BuMan)) AS BuManName, BuDate, ModifyMan, (SELECT NAME FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = c.ModifyMan)) AS ModifyManName, ModifyDate, OriDate FROM CarServicePre AS c WHERE (ServicePreNo = @ServicePreNo)"
        DeleteCommand="DELETE FROM CarServicePre WHERE (ServicePreNo = @ServicePreNo)"
        InsertCommand="INSERT INTO CarServicePre(ServicePreNo, DepNo, Car_ID, Driver, ServiceType, LastDate, ServiceDate, Remark, BuMan, BuDate, OriDate) VALUES (@ServicePreNo, @DepNo, @Car_ID, @Driver, @ServiceType, @LastDate, @ServiceDate, @Remark, @BuMan, @BuDate, @OriDate)"
        UpdateCommand="UPDATE CarServicePre SET ServiceType = @ServiceType, ServiceDate = @ServiceDate, Driver = @Driver, Remark = @Remark, ModifyMan = @ModifyMan, ModifyDate = @ModifyDate WHERE (ServicePreNo = @ServicePreNo)">
        <DeleteParameters>
            <asp:Parameter Name="ServicePreNo" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="ServicePreNo" />
            <asp:Parameter Name="DepNo" />
            <asp:Parameter Name="Car_ID" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="ServiceType" />
            <asp:Parameter Name="LastDate" />
            <asp:Parameter Name="ServiceDate" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="BuMan" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="OriDate" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="gridDailyList" Name="ServicePreNo" PropertyName="SelectedValue" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="ServiceType" />
            <asp:Parameter Name="ServiceDate" />
            <asp:Parameter Name="Driver" />
            <asp:Parameter Name="ServicePreNo" />
            <asp:Parameter Name="Remark" />
            <asp:Parameter Name="ModifyMan" />
            <asp:Parameter Name="ModifyDate" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsServiceType_INS" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT RTRIM(LTRIM(CLASSNO)) AS ClassNo, CLASSTXT FROM DBDICB WHERE (CLASSNO IN ('1', '2')) AND (FKEY = '工作單A         FixworkA        SERVICE')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsServiceType_Edit" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT RTRIM(LTRIM(CLASSNO)) AS ClassNo, CLASSTXT FROM DBDICB WHERE (CLASSNO IN ('1', '2')) AND (FKEY = '工作單A         FixworkA        SERVICE')"></asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDepList" runat="server" ConnectionString="<%$ ConnectionStrings:connERPSQL %>" ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT DEPNO, DEPNO + '-' + NAME AS DepName FROM DEPARTMENT WHERE 1&lt;&gt;1"></asp:SqlDataSource>
</asp:Content>
