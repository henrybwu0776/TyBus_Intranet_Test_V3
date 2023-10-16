<%@ Page Title="" Language="C#" MasterPageFile="~/MainPage.Master" AutoEventWireup="true" CodeBehind="RequestDriverWorkData.aspx.cs" Inherits="TyBus_Intranet_Test_V3.RequestDriverWorkData" %>

<asp:Content ID="RequestDriverWorkDataForm" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div>
        <a class="titleText-Red">駕駛員公里數及津貼查詢</a>
    </div>
    <br />
    <asp:Panel ID="plSearch" runat="server" GroupingText="查詢駕駛員資料" CssClass="SearchPanel">
        <table class="TableSetting">
            <tr>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDriver_Search" runat="server" CssClass="text-Right-Blue" Text="駕駛員工號：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="2">
                    <asp:TextBox ID="eDriver_Search" runat="server" CssClass="text-Left-Black" AutoPostBack="true" OnTextChanged="eDriver_Search_TextChanged" Width="40%" />
                    <asp:Label ID="eDriverName_Search" runat="server" CssClass="text-Left-Black" Width="50%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col">
                    <asp:Label ID="lbDriveYM_Search" runat="server" CssClass="text-Right-Blue" Text="行車年月：" Width="90%" />
                </td>
                <td class="ColHeight ColBorder ColWidth-7Col" colspan="3">
                    <asp:TextBox ID="eDriveYM_Search" runat="server" CssClass="text-Left-Black" Width="30%" />
                    <asp:Label ID="eDriveYM_Note" runat="server" CssClass="text-Left-Red" Text=" (西元年 4 碼 + 月份 2 碼 ; 範例：201801)" Width="55%" />
                </td>
            </tr>
            <tr>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbSearch" runat="server" CssClass="button-Black" Text="查詢" OnClick="bbSearch_Click" Width="90%" />
                </td>
                <td class="ColHeight ColWidth-7Col">
                    <asp:Button ID="bbCancel" runat="server" CssClass="button-Red" Text="結束" OnClick="bbCancel_Click" Width="90%" />
                </td>
            </tr>
            <tr>
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
                <td class="ColWidth-7Col" />
            </tr>
        </table>
    </asp:Panel>
    <asp:ScriptManager ID="smMain" runat="server"></asp:ScriptManager>
    <asp:Panel ID="plShowData" runat="server" GroupingText="查詢結果" CssClass="ShowPanel">
        <asp:FormView ID="fvShowDriverData" runat="server" DataSourceID="sdsDriverData" Width="100%">
            <ItemTemplate>
                <table class="TableSetting">
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbBuDate_List" runat="server" CssClass="text-Right-Blue" Text="統計年月：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eBuDate_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("BuDate","{0: yyyy/MM/dd}")%>' />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbDepNo_List" runat="server" CssClass="text-Right-Blue" Text="所屬單位：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eDepNo_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("DepNo")%>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eDepNoName_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("DepNoName")%>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbCarType_List" runat="server" CssClass="text-Right-Blue" Text="單位類別：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eCarType_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("CarType")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="lbEmpNo_List" runat="server" CssClass="text-Right-Blue" Text="駕駛員：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eEmpNo_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("EmpNo")%>' Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col" colspan="2">
                            <asp:Label ID="eDriverName_List" runat="server" Text='<%#Eval("DriverName")%>' CssClass="text-Left-Blue" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col TableColAlign_Right">
                            <asp:Label ID="lbLeaveDay_List" runat="server" CssClass="text-Right-Blue" Text="離職日期：" Width="90%" />
                        </td>
                        <td class="ColHeight ColBorder ColWidth-6Col">
                            <asp:Label ID="eLeaveDay_List" runat="server" CssClass="text-Left-Blue" Text='<%#Eval("LeaveDay","{0: yyyy/MM/dd}")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight_Low" colspan="6"></td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbRoutineCashAMT_List" runat="server" CssClass="text-Right-Black" Text="班車現金收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eRoutineCashAMT_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("RoutineCashAMT")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbBusCashAMT_List" runat="server" CssClass="text-Right-Black" Text="公車現金收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eBusCashAMT_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("BusCashAMT")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbCashAMT_List" runat="server" CssClass="text-Right-Black" Text="現金收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eCashAMT_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("CashAMT")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbRoutineCardAMT_List" runat="server" CssClass="text-Right-Black" Text="班車普卡+聯營卡收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eRoutineCardAMT_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("RoutineCardAMT")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder a3 TableColAlign_Right">
                            <asp:Label ID="lbBusCardAMT_List" runat="server" CssClass="text-Right-Black" Text="公車普卡+聯營卡收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eBusCardAMT_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("BusCardAMT")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbCardAMT_List" runat="server" CssClass="text-Right-Black" Text="普卡+聯營卡收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eCardAMT_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("CardAMT")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbSCardQTY_List" runat="server" CssClass="text-Right-Black" Text="學生卡人數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eSCardQTY_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("SCardQTY")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbOCardQTY_List" runat="server" CssClass="text-Right-Black" Text="敬老人數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eOCardQTY_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("OCardQTY")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbPoolLine_List" runat="server" CssClass="text-Right-Black" Text="聯營路線收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="ePoolLine_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("PoolLine")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbTraffic_List" runat="server" CssClass="text-Right-Black" Text="交通車收入：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eTraffic_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("Traffic")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight_Low" colspan="6"></td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col" colspan="4"></td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbTotalBounds2_List" runat="server" CssClass="text-Right-Red" Text="載客獎金：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eTotalBounds2_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("TotalBounds2")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight_Low" colspan="6"></td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbRoutineKM101_List" runat="server" CssClass="text-Right-Black" Text="班車 (聯營) 公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eRoutineKM101_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("RoutineKM101")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbBusKM_List" runat="server" CssClass="text-Right-Black" Text="公車公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eBusKM_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("BusKM")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbTrafficKM_List" runat="server" CssClass="text-Right-Black" Text="交通車公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eTrafficKM_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("TrafficKM")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbAreaKM_List" runat="server" CssClass="text-Right-Black" Text="區間租車公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eAreaKM_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("AreaKM")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbSightseeingKM_List" runat="server" CssClass="text-Right-Black" Text="遊覽車公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eSightseeingKM_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("SightseeingKM")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbSpecialKM_List" runat="server" CssClass="text-Right-Black" Text="專車公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eSpecialKM_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("SpecialKM")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbOperationKM_List" runat="server" CssClass="text-Right-Black" Text="非營運公里數：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eOperationKM_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("OperationKM")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbMid_SizeBus_List" runat="server" CssClass="text-Right-Black" Text="借中巴加給：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eMid_SizeBus_List" runat="server" CssClass="text-Left-Black" Text='<%#Eval("Mid_SizeBus")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight_Low" colspan="6"></td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbTotalKMBounds2_List" runat="server" CssClass="text-Right-Red" Text="公里獎金合計：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eTotalKMBounds2_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("TotalKMBounds2")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbLinesAllowance2_List" runat="server" CssClass="text-Right-Red" Text="路線補貼 / 租車趟次加給：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eLinesAllowance2_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("LinesAllowance2")%>' Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="lbFarawayAllowance2_List" runat="server" CssClass="text-Right-Red" Text="偏遠路線加給：" Width="90%" />
                        </td>
                        <td class="ColWidth-6Col ColHeight ColBorder">
                            <asp:Label ID="eFarawayAllowance2_List" runat="server" CssClass="text-Left-Red" Text='<%#Eval("FarawayAllowance2")%>' Width="90%" />
                        </td>
                    </tr>
                    <tr>
                        <td class="ColWidth-6Col" />
                        <td class="ColWidth-6Col" />
                        <td class="ColWidth-6Col" />
                        <td class="ColWidth-6Col" />
                        <td class="ColWidth-6Col" />
                        <td class="ColWidth-6Col" />
                    </tr>
                </table>
            </ItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <asp:SqlDataSource ID="sdsDriverData" runat="server"
        ConnectionString="<%$ ConnectionStrings:connERPSQL %>"
        ProviderName="<%$ ConnectionStrings:connERPSQL.ProviderName %>"
        SelectCommand="SELECT budate, depno, (SELECT NAME FROM DEPARTMENT WHERE (DEPNO = RsEmpMonth.depno)) AS DepNoName, empno, (SELECT NAME FROM EMPLOYEE WHERE (EMPNO = RsEmpMonth.empno)) AS DriverName, SUM(ISNULL(ACashAmt, 0)) AS RoutineCashAMT, SUM(ISNULL(ACashAmt1, 0)) AS BusCashAMT, SUM(ISNULL(ACashAmt, 0) + ISNULL(ACashAmt1, 0)) AS CashAMT, SUM(ISNULL(Aoamt, 0)) AS RoutineCardAMT, SUM(ISNULL(Aoamt1, 0)) AS BusCardAMT, SUM(ISNULL(Aoamt, 0) + ISNULL(Aoamt1, 0)) AS CardAMT, SUM(ISNULL(BScount, 0)) AS RoutineSCardQTY, SUM(ISNULL(BScount1, 0)) AS BusSCardQTY, SUM(ISNULL(BScount, 0) + ISNULL(BScount1, 0)) AS SCardQTY, SUM(ISNULL(BOldCount, 0)) AS RoutineOCardQTY, SUM(ISNULL(BOldCount1, 0)) AS BusOCardQTY, SUM(ISNULL(BOldCount, 0) + ISNULL(BOldCount1, 0)) AS OCardQTY, SUM(ISNULL(ALines, 0) + ISNULL(ALines1, 0)) AS PoolLine, SUM(ISNULL(BOldQty, 0) + ISNULL(BOldQty1, 0)) AS BOldQTY, SUM(ISNULL(AunTraincome, 0)) AS Traffic, CASE WHEN CharIndex('公車站' , (SELECT [Name] FROM Department WHERE Department.DepNo = RSEmpMonth.DepNo) , 0) &gt; 0 THEN '公車站' WHEN CharIndex('站' , (SELECT [Name] FROM Department WHERE Department.DepNo = RSEmpMonth.DepNo) , 0) &gt; 0 THEN '班車站' END AS CarType, SUM(ISNULL(DCashAmt, 0)) AS DCashAMT, SUM(ISNULL(DOamt, 0)) AS DOAMT, SUM(ISNULL(DRentTraAmt, 0)) AS DRentTraAMT, SUM(ISNULL(DLinesAmt, 0)) AS DLinesAMT, SUM(ISNULL(DspieceAmt, 0)) AS DspieceAMT, SUM(ISNULL(DOldQtyAmt, 0)) AS DOldQtyAMT, SUM(ISNULL(DOldAmt, 0)) AS DOldAMT, SUM(ISNULL(Aoamt_A, 0)) AS AOAMT_A, SUM(ISNULL(Aoamt1_A, 0)) AS AOAMT1_A, SUM(ISNULL(Aoamt_B, 0)) AS AOAMT_B, SUM(ISNULL(Aoamt1_B, 0)) AS AOAMT1_B, SUM(ISNULL(ALines, 0)) AS ALines, SUM(ISNULL(ALines1, 0)) AS ALines1, CAST((((((((((SUM(ISNULL(ACashAmt, 0)) / @CashBoundsR + SUM(ISNULL(ACashAmt1, 0)) / @CashBoundsB) + SUM(ISNULL(Aoamt_A, 0)) / @NormalBoundsR) + SUM(ISNULL(Aoamt_B, 0)) / @NormalBoundsR) + SUM(ISNULL(Aoamt1_A, 0)) / @NormalBoundsB) + SUM(ISNULL(Aoamt1_B, 0)) / @NormalBoundsB) + SUM(ISNULL(AunTraincome, 0)) / @ContractBoundsR) + SUM(ISNULL(ALines, 0)) / @HighwayCashR) + SUM(ISNULL(ALines1, 0)) / @HighwayCashR) + SUM(ISNULL(BScount, 0)) * @StudentBoundsR) + SUM(ISNULL(BScount1, 0)) * @StudentBoundsB AS Decimal(10 , 0)) + CAST(SUM(ISNULL(BOldCount, 0)) * @OlderBoundsR + SUM(ISNULL(BOldCount1, 0)) * @OlderBoundsB AS Decimal(10 , 0)) AS TotalBounds, (SELECT ISNULL(cashnum05, 0) + ISNULL(cashnum09, 0) AS Expr1 FROM PAYREC WHERE (empno = RsEmpMonth.empno) AND (paydate = RsEmpMonth.budate) AND (paydur = '1')) AS TotalBounds2, SUM(ISNULL(CBus1Km, 0)) AS BusKM, SUM(ISNULL(CBus2Km, 0)) AS RoutineKM, SUM(ISNULL(CBus2Km, 0) + ISNULL(CBus4Km, 0)) AS RoutineKM101, SUM(ISNULL(CRentTraKm, 0)) AS TrafficKM, SUM(ISNULL(CRentAKm, 0)) AS AreaKM, SUM(ISNULL(CRentBKm, 0)) AS SightSeeingKM, SUM(ISNULL(CBus3Km, 0)) AS SpecialKM, SUM(ISNULL(CBus4Km, 0)) AS UnionKM, SUM(ISNULL(CBus5Km, 0)) AS OperationKM, SUM(ISNULL(CBus6Km, 0)) AS CBus6KM, SUM(ISNULL(CBus7Km, 0)) AS CBus7KM, SUM(ISNULL(CBus8Km, 0)) AS CBus8KM, SUM(ISNULL(CBus9Km, 0)) AS CBus9KM, SUM(ISNULL(Emid_BusAmt, 0)) AS Mid_SizeBus, SUM(ISNULL(EIsExtraAmt, 0) + ISNULL(EIsExtraAmt1, 0)) AS IsPlus, SUM(ISNULL(EBus1KmAmt, 0)) AS RoutineKMAMT, SUM(ISNULL(EBus2KmAmt, 0)) AS BusKMAMT, SUM(ISNULL(ERentTraKmAmt, 0)) AS ERentTraKMAMT, SUM(ISNULL(ERentAKmAmt, 0)) AS ERentAKMAMT, SUM(ISNULL(ERentBKmAmt, 0)) AS ERentBKMAMT, SUM(ISNULL(EBus3KmAmt, 0)) AS SpecialKMAMT, SUM(ISNULL(EBus5KmAmt, 0)) AS OperationAMT, SUM(ISNULL(EIsPlusAmt, 0)) AS EIsPlusAMT, SUM(ISNULL(RentNumberAmt, 0)) AS RentNumberAMT, SUM(ISNULL(EBus4KmAmt, 0)) AS UnionKMAMT, CAST((SELECT ISNULL(SUM(CASE WHEN b.ActualKM &gt;= c.FeedKm THEN c.FeedAmt2 ELSE c.FeedAmt1 END), 0) AS FeedAMT FROM RUNSHEETA AS a LEFT OUTER JOIN RunSheetB AS b ON b.AssignNo = a.ASSIGNNO LEFT OUTER JOIN Lines AS c ON c.LinesNo = b.LinesNo WHERE (a.DRIVER = RsEmpMonth.empno) AND (c.IsFeed = 'V') AND (ISNULL(b.ReduceReason, '') = '') AND (a.BUDATE BETWEEN @BuDate AND @EndDate)) + (SELECT SUM(ISNULL(CAST(RentNumber AS float), 0)) AS RentNumber FROM RUNSHEETA WHERE (DRIVER = RsEmpMonth.empno) AND (BUDATE BETWEEN @BuDate AND @EndDate)) * @RentTimesBounds AS decimal(10 , 0)) AS LinesAllowance, CAST(SUM(ISNULL(CBus6Km, 0)) * @RuleLinesBounds + SUM(ISNULL(CBus7Km, 0)) * @BusLinesBounds AS Decimal(10 , 0)) AS FarawayAllowance, CAST(((((((((SUM(ISNULL(CBus1Km, 0)) * @RuleBounds + SUM(ISNULL(CBus3Km, 0)) * @SPECBoundsKM) + SUM(ISNULL(CBus2Km, 0)) * @BusBounds) + SUM(ISNULL(CBus4Km, 0)) * @UnionKMBounds) + SUM(ISNULL(CRentTraKm, 0)) * @TransBoundsR) + SUM(ISNULL(CRentAKm, 0)) * @RentBoundsR) + SUM(ISNULL(CRentBKm, 0)) * @TourBoundsR) + SUM(ISNULL(CBus5Km, 0)) * @NoneBusiBoundsR) + SUM(ISNULL(CBus8Km, 0)) * @MiniBusBoundsR) + SUM(ISNULL(CBus9Km, 0)) * @MiniBusBoundsB AS Decimal(10 , 0)) AS TotalKMBounds, (SELECT ISNULL(cashnum04, 0) AS Expr1 FROM PAYREC AS PAYREC_3 WHERE (empno = RsEmpMonth.empno) AND (paydate = RsEmpMonth.budate) AND (paydur = '1')) AS TotalKMBounds2, (SELECT ISNULL(cashnum08, 0) AS Expr1 FROM PAYREC AS PAYREC_2 WHERE (empno = RsEmpMonth.empno) AND (paydate = RsEmpMonth.budate) AND (paydur = '1')) AS LinesAllowance2, (SELECT ISNULL(cashnum10, 0) AS Expr1 FROM PAYREC AS PAYREC_1 WHERE (empno = RsEmpMonth.empno) AND (paydate = RsEmpMonth.budate) AND (paydur = '1')) AS FarawayAllowance2, (SELECT LEAVEDAY FROM EMPLOYEE AS EMPLOYEE_1 WHERE (EMPNO = RsEmpMonth.empno)) AS LeaveDay FROM RsEmpMonth WHERE (budate = @BuDate) AND (empno = @EmpNo) AND (depno = @DepNo) GROUP BY budate, depno, empno">
        <SelectParameters>
            <asp:Parameter Name="CashBoundsR" />
            <asp:Parameter Name="CashBoundsB" />
            <asp:Parameter Name="NormalBoundsR" />
            <asp:Parameter Name="NormalBoundsB" />
            <asp:Parameter Name="ContractBoundsR" />
            <asp:Parameter Name="HighwayCashR" />
            <asp:Parameter Name="StudentBoundsR" />
            <asp:Parameter Name="StudentBoundsB" />
            <asp:Parameter Name="OlderBoundsR" />
            <asp:Parameter Name="OlderBoundsB" />
            <asp:Parameter Name="BuDate" />
            <asp:Parameter Name="EndDate" />
            <asp:Parameter Name="RentTimesBounds" />
            <asp:Parameter Name="RuleLinesBounds" />
            <asp:Parameter Name="BusLinesBounds" />
            <asp:Parameter Name="RuleBounds" />
            <asp:Parameter Name="SPECBoundsKM" />
            <asp:Parameter Name="BusBounds" />
            <asp:Parameter Name="UnionKMBounds" />
            <asp:Parameter Name="TransBoundsR" />
            <asp:Parameter Name="RentBoundsR" />
            <asp:Parameter Name="TourBoundsR" />
            <asp:Parameter Name="NoneBusiBoundsR" />
            <asp:Parameter Name="MiniBusBoundsR" />
            <asp:Parameter Name="MiniBusBoundsB" />
            <asp:Parameter Name="EmpNo" />
            <asp:Parameter Name="DepNo" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
