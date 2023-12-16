Imports ggcRetailSales
Imports ggcAppDriver

Public Class frmMain1
    'fixed image max count
    Private Const pxeMaxTable As Integer = 12
    Private Const pxeMaxCatgr As Integer = 4
    Private Const pxeMaxDtail As Integer = 29

    'image scroll handling
    Private pnTablePage As Integer = 0
    Private pnCategPage As Integer = 0
    Private pnDtailPage As Integer = 0
    Private pnTotalTble As Integer = 0

    Private pnToolTip As New ToolTip()

    Private pnLoadx As Integer
    Private pnActiveRow As Integer
    Private psCategrID As String

    Private p_sLogName As String

    Private WithEvents p_oTrans As New_Sales_Order

    Private Sub newOrder()
        Call InitializeDataGrid()

        Call initGrid()
        Call clearFields()

        Call initForm(xeEditMode.MODE_ADDNEW)

        pnCategPage = 0
        pnDtailPage = 0
        pnTablePage = 0

        p_oTrans = New New_Sales_Order(p_oAppDriver)

        p_oTrans.ValidDailySales = pbValidSales
        p_oTrans.PosDate = pdPOSDatex
        p_oTrans.ValidDailySales = True
        p_oTrans.SalesStatus = pnSaleStat
        p_oTrans.Cashier = psCashierx
        p_oTrans.LogName = modMain.p_sLogName

        Call initDetailImages()
        Call initCategoryImages()

        loadMaster()
        loadTable()

        If p_oTrans.LoadOrder Then loadOrder()
    End Sub

    Private Sub procValues(ByVal fnIndex As Integer)
        With p_oTrans
            If .ItemCount = 0 Then Exit Sub 'check if no detail added
            If .Detail(pnActiveRow, "cDetailxx") = 1 Then Exit Sub 'check if detail is a combo item
            If .Detail(pnActiveRow, "cReversed") = "1" Then Exit Sub 'check if it was tagged as reversed
            If .Detail(pnActiveRow, "cReversex") = "-" Then Exit Sub

            Dim lnValue As Double
            lnValue = 1

            .Quantity = CInt(Decimal.Truncate(IIf(lnValue = 0, 1, lnValue)))

            Select Case fnIndex
                Case 0 'reverse
                    If .ReverseOrder(pnActiveRow, .Detail(pnActiveRow, "nQuantity")) Then showComputation()
                Case 1 'deduct item
                    If .Quantity > .Detail(pnActiveRow, "nQuantity") Then Exit Sub

                    Dim lbRefresh As Boolean
                    lbRefresh = IIf(.Detail(pnActiveRow, "nQuantity") = 1, True, False)

                    If .DeductItem(pnActiveRow) Then showComputation()

                    If lbRefresh Then loadDetail()
                Case 2 'add item
                    If .AddItem(pnActiveRow) Then showComputation()
                Case 3 'change price
                    If .ChangePrice(pnActiveRow) Then showComputation()
                Case 4 'change quantity
                    If .ChangeQty(pnActiveRow) Then
                        loadDetail()
                        showComputation()
                    End If
            End Select
        End With
    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub

#Region "Form Events"
    Private Sub Form_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If pnLoadx = 0 Then
            pnLoadx = 1

            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)
            Call grpEventHandler(Me, GetType(Label), "lblTable", "Click", AddressOf lblTable_Click)
            Call grpEventHandler(Me, GetType(Label), "lblTable", "DoubleClick", AddressOf lblTable_DoubleClick)

            Call newOrder()

            Dim row As DataRow

            txtDetail00.AutoCompleteCustomSource.Clear()
            For Each row In p_oTrans.SearchItem.Rows
                txtDetail00.AutoCompleteCustomSource.Add(row.Item("sBriefDsc").ToString())
            Next

            txtDetail00.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtDetail00.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        End If
    End Sub

    Private Sub Form_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                'If MsgBox("Print Cashier Transaction?", _
                '              MsgBoxStyle.Question & MsgBoxStyle.YesNo, _
                '              "Confirm") = MsgBoxResult.Yes Then

                '    poSales.PrintCashierSales("20170127", Environment.GetEnvironmentVariable("RMS-CRM-No"), p_oAppDriver.UserID)
                'End If

                If MsgBox("You are going to Log out. Continue?", _
                              MsgBoxStyle.Question & MsgBoxStyle.YesNo, _
                              "Confirm") = MsgBoxResult.Yes Then

                    p_oAppDriver.SaveEvent("0027", "", p_oTrans.SerialNo)
                    Me.Close()
                End If

            Case Keys.F5
                Call newOrder()
            Case Keys.Add
                Call procValues(2)
                e.SuppressKeyPress = True
            Case Keys.Subtract
                Call procValues(1)
                e.SuppressKeyPress = True
            Case Keys.F9 'pay charge
                p_oTrans.PrintChargeOR()
                Call newOrder()
            Case Keys.F10 'tag charge
                p_oTrans.PayCharge()
                Call newOrder()
            Case Keys.F11 'browse order
                p_oTrans.BrowseOrder()
            Case Keys.F12 'reprint or
                p_oTrans.Reprint()
        End Select
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        lblCurTimex.Text = Format(p_oAppDriver.SysDate, "hh:mm:ss tt")
        lblCurTimex.Refresh()
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        With DataGridView1
            pnActiveRow = .CurrentCell.RowIndex
        End With
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        With p_oTrans
            Select Case lnIndex
                Case 0 'Logout
                    If MsgBox("Print Cashier Transaction?", _
                              MsgBoxStyle.Question & MsgBoxStyle.YesNo, _
                              "Confirm") = MsgBoxResult.Yes Then
                        poSales.PrintCashierSales(Format(p_oAppDriver.SysDate, "yyyyMMdd"), Environment.GetEnvironmentVariable("RMS-CRM-No"), p_oAppDriver.UserID)
                    End If

                    If MsgBox("You are going to Log out. Continue?", _
                                  MsgBoxStyle.Question & MsgBoxStyle.YesNo, _
                                  "Confirm") = MsgBoxResult.Yes Then
                        Me.Close()
                    End If
                Case 1 'Print Z Reading
                    If p_oTrans.ProcTZReading Then Me.Close()
                Case 2 'Save Order
                    If p_oTrans.SaveTransaction() Then newOrder()
                Case 3 'Merge Order
                    Call initPanel(0, False)

                    If .MergeOrder() Then newOrder()

                    Call initPanel(0, True)
                Case 4, 6, 7, 8
                    Call initPanel(1, False)

                    If lnIndex = 4 Then
                        If .PayOrder() Then newOrder()
                    ElseIf lnIndex = 6 Then
                        If .ChargeOrder() Then newOrder()
                    ElseIf lnIndex = 7 Then
                        If .Complementary() Then newOrder()
                    Else
                        If DataGridView1.Rows(0).Cells(1).Value <> "" Then
                            If .IssueDiscount() Then newOrder()
                        End If
                    End If

                    Call initPanel(1, True)
                Case 5, 9
                    Call initPanel(0, False)

                    If lnIndex = 5 Then
                        If .SplitOrder() Then newOrder()
                    Else
                        If .SalesReturn() Then newOrder()
                    End If

                    Call initPanel(0, True)
                Case 10 'cash pullout
                    Dim loForm As frmCashPullout

                    loForm = New frmCashPullout(p_oAppDriver)
                    loForm.Cashier = psCashierx
                    loForm.TopMost = True
                    loForm.ShowDialog()
                Case 11 'void
                    If .VoidOrder() Then newOrder()
                Case 12 'print bill
                    If Not .PrintBill Then
                        MsgBox("Unable to print billing statement.")
                    End If
                Case 13 'end shift
                    If .ProcTXReading() Then Me.Close()
                Case 14 'cancel receipt
                    .CancelOR()
                Case 15 'pay charge
                    .PayCharge()
                Case 16
                    Call newOrder()
            End Select
        End With
endProc:
        txtDetail00.Focus()
        Exit Sub
    End Sub

    Private Sub lblTable_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loLbl As Label
        loLbl = CType(sender, System.Windows.Forms.Label)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loLbl.Name, 9))

        If p_oTrans.LoadTable(loLbl.Text) Then
            loadMaster()
            loadDetail()
        End If

    End Sub

    Private Sub lblTable_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loLbl As Label
        loLbl = CType(sender, System.Windows.Forms.Label)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loLbl.Name, 9))

        If p_oTrans.ManageTable(CInt(loLbl.Text)) Then newOrder()
        showComputation()
    End Sub
#End Region

#Region "Load Data"
    Private Sub loadOrder()
        loadMaster()
        loadDetail()
    End Sub

    Private Sub loadMaster()
        With p_oTrans
            'lblMaster00.Text = Strings.Right(.Master("sTransNox"), 20)
            lblMaster00.Text = Strings.Right(.Master("sTransNox"), 16)
            lblMaster05.Text = getCashier(p_oAppDriver.UserID)

            lblMaster80.Text = .POSNumber
            lblMaster81.Text = .Accreditation
            lblMaster90.Text = Format(poSales.POSDate, xsDATE_LONG)
            lblMaster91.Text = "," & Format(poSales.POSDate, "dddd")
        End With

    End Sub

    Private Sub loadDetail(ByVal lnRow As Integer)
        Dim lnCtr As Integer

        With DataGridView1
            For lnCtr = lnRow To p_oTrans.ItemCount - 1
                .Item(0, lnCtr).Value = lnCtr + 1
                .Item(1, lnCtr).Value = p_oTrans.Detail(lnCtr, "sBarcodex")
                .Item(2, lnCtr).Value = p_oTrans.Detail(lnCtr, "sBriefDsc")
                .Item(3, lnCtr).Value = IIf(p_oTrans.Detail(lnCtr, "cReversex") = "-", p_oTrans.Detail(lnCtr, "cReversex"), "") & p_oTrans.Detail(lnCtr, "nQuantity")
                .Item(4, lnCtr).Value = Format(p_oTrans.Detail(lnCtr, "nUnitPrce") * _
                                                (100 - p_oTrans.Detail(lnCtr, "nDiscount")) / 100 -
                                                    p_oTrans.Detail(lnCtr, "nAddDiscx"), "#,##0.00")
                .Item(5, lnCtr).Value = Format(.Item(3, lnCtr).Value * .Item(4, lnCtr).Value, "#,##0.00")
            Next

            If .RowCount > 14 Then
                .Columns(2).Width = 160
                .FirstDisplayedScrollingRowIndex = .RowCount - 14
            Else
                .Columns(2).Width = 160
            End If
        End With

        showComputation()
    End Sub

    Private Sub loadDetail()
        Call initGrid()

        With DataGridView1
            Dim lnCtr As Integer

            .RowCount = IIf(p_oTrans.ItemCount = 0, 1, p_oTrans.ItemCount)

            If .RowCount > 14 Then
                .Columns(2).Width = 160
                .FirstDisplayedScrollingRowIndex = .RowCount - 14
            Else
                .Columns(2).Width = 160
            End If

            For lnCtr = 0 To p_oTrans.ItemCount - 1
                .Item(0, lnCtr).Value = lnCtr + 1
                .Item(1, lnCtr).Value = p_oTrans.Detail(lnCtr, "sBarcodex")
                .Item(2, lnCtr).Value = p_oTrans.Detail(lnCtr, "sBriefDsc")
                .Item(3, lnCtr).Value = IIf(p_oTrans.Detail(lnCtr, "cReversex") = "-", p_oTrans.Detail(lnCtr, "cReversex"), "") & p_oTrans.Detail(lnCtr, "nQuantity")
                '.Item(4, lnCtr).Value = Format(p_oTrans.Detail(lnCtr, "nUnitPrce") * _
                '                                (100 - p_oTrans.Detail(lnCtr, "nDiscount")) / 100 -
                '                                    p_oTrans.Detail(lnCtr, "nAddDiscx"), "#,##0.00")
                .Item(4, lnCtr).Value = Format(p_oTrans.Detail(lnCtr, "nUnitPrce"), "#,##0.00")
                .Item(5, lnCtr).Value = Format(.Item(3, lnCtr).Value * .Item(4, lnCtr).Value, "#,##0.00")
            Next

            showComputation()

            pnActiveRow = .RowCount - 1
            .ClearSelection()
            .Rows(pnActiveRow).Selected = True

            Call clearSeek()
        End With
    End Sub

    Private Sub showComputation()
        'kalyptus - 2017.01.20 01:21pm
        'Deduct nVoidTotl from nTranTotl to show the actual sales after the reversal of an ordered item...
        Dim lnSrvCrge As Double
        Dim lnAmntDuex As Double

        lblMaster04.Text = FormatNumber(p_oTrans.Master("nTranTotl") - p_oTrans.Master("nVoidTotl"), 2) 'sales total

        lblMaster13.Text = FormatNumber(p_oTrans.Master("nVATSales"), 2) 'vat sales
        lblMaster14.Text = FormatNumber(p_oTrans.Master("nVATAmtxx"), 2) 'vat amount
        lblMaster15.Text = FormatNumber(p_oTrans.Master("nNonVATxx"), 2) 'non vat 

        lblMaster17.Text = FormatNumber(p_oTrans.Master("nDiscount") + p_oTrans.Master("nVatDiscx") + p_oTrans.Master("nPWDDiscx"), 2) 'discounts
        'jovan 03-12-2021
        lnAmntDuex = FormatNumber(CDbl(lblMaster04.Text) - CDbl(lblMaster17.Text), 2) 'amount due
        'lblAmount.Text = FormatNumber(CDbl(lblMaster04.Text) - CDbl(lblMaster17.Text), 2) 'amount due

        lnSrvCrge = IFNull(p_oTrans.Master("nSChargex"), 0)
        lblSrvCrge.Text = FormatNumber(lnSrvCrge, 2)
        lblAmount.Text = FormatNumber(lnAmntDuex + lnSrvCrge, 2)
        p_oTrans.setTranTotal = lblAmount.Text
    End Sub

    'Private Sub showComputationNew()
    '    'jovan - 2020.10.15 revised presentation of discount in interface

    '    Dim lnSubTotal As Double
    '    Dim lnAddDiscount As Double
    '    Dim lnPWDiscount As Double
    '    Dim lnVATExcls As Double
    '    Dim lnDiscount As Double
    '    Dim lnVatRatex As Double = 1.12
    '    Dim lnSrvCrge As Double
    '    Dim lnAmntDuex As Double

    '    lnSubTotal = p_oTrans.Master("nTranTotl") - p_oTrans.Master("nVoidTotl")

    '    lblMaster04.Text = FormatNumber(lnSubTotal, 2) 'sales total
    '    lnVATExcls = lnSubTotal / lnVatRatex
    '    lblMaster13.Text = FormatNumber(lnVATExcls, 2) 'vat sales

    '    lnDiscount = p_oTrans.Master("nDiscount") / lnVatRatex

    '    lnAddDiscount = p_oTrans.Master("nVatDiscx") + p_oTrans.Master("nPWDDiscx")
    '    lnPWDiscount = p_oTrans.Master("nPWDDiscx")
    '    lblMaster17.Text = FormatNumber(lnDiscount, 2) 'discounts
    '    lblMaster14.Text = FormatNumber(p_oTrans.Master("nVATAmtxx"), 2) 'vat amount
    '    lblMaster15.Text = FormatNumber(lnPWDiscount, 2)   'non vat 
    '    lblNetSale.Text = FormatNumber(lnVATExcls - (lnDiscount + lnPWDiscount), 2)
    '    If lnPWDiscount = 0 Then
    '        lnAmntDuex = FormatNumber(lblNetSale.Text + p_oTrans.Master("nVATAmtxx"), 2)
    '    Else
    '        lnAmntDuex = FormatNumber(lnVATExcls - lnPWDiscount, 2)
    '    End If

    '    lnSrvCrge = IFNull(p_oTrans.Master("nSChargex"), 0)
    '    lblSrvCrge.Text = FormatNumber(lnSrvCrge, 2)
    '    lblAmount.Text = FormatNumber(lnAmntDuex + lnSrvCrge, 2)
    '    p_oTrans.setTranTotal = lblAmount.Text

    'End Sub
#End Region

#Region "PictureBox"
    Private Sub clearDetailImage(ByVal Index As Integer)
        Dim loPic As PictureBox

        loPic = CType(FindPictureBox(Me, "picDetail" & Format(Index, "00")), PictureBox)
        loPic.BackgroundImage = Nothing
    End Sub

    Private Sub clearCategoryImage(ByVal Index As Integer)
        Dim loPic As PictureBox

        loPic = CType(FindPictureBox(Me, "picCategr" & Format(Index, "00")), PictureBox)
        loPic.BackgroundImage = Nothing
    End Sub

    Private Sub initCategoryImages()
        Dim lnCtr As Integer
        Dim loPic As PictureBox
        Dim loDT As DataTable
        Dim lnTot As Integer
        Dim lnRow As Integer

        loDT = p_oTrans.GetCategory

        If loDT.Rows.Count > 0 Then
            lnTot = (pnCategPage * pxeMaxCatgr)

            For lnCtr = 0 + lnTot To lnTot + pxeMaxCatgr - 1
                lnRow = IIf(pnCategPage = 0, lnCtr, lnCtr - lnTot)

                If lnRow < pxeMaxCatgr Then
                    If lnCtr <= loDT.Rows.Count - 1 Then
                        loPic = CType(FindPictureBox(Me, "picCategr" & Format(lnCtr - lnTot, "00")), PictureBox)
                        loPic.BackColor = Color.Transparent
                        loPic.BackgroundImage = Nothing
                        loPic.BorderStyle = BorderStyle.Fixed3D
                        loPic.Tag = loDT(lnCtr)("sCategrCd")

                        loadCategoryImages(lnCtr - lnTot, loDT(lnCtr)("sImgePath"))
                    Else
                        clearCategoryImage(lnCtr - lnTot)
                    End If
                End If
            Next
        Else
            For lnCtr = 0 To pxeMaxCatgr - 1
                loPic = CType(FindPictureBox(Me, "picCategr" & Format(lnCtr, "00")), PictureBox)
                loPic.BackColor = Color.Transparent
                loPic.BackgroundImage = Nothing
                loPic.BorderStyle = BorderStyle.Fixed3D
                loPic.Tag = ""

                clearCategoryImage(lnCtr)
            Next
        End If
    End Sub

    Private Sub initDetailImages(Optional ByVal fsCategrCd As String = "")
        Dim lnCtr As Integer
        Dim loPic As PictureBox
        Dim loDT As DataTable

        For lnCtr = 0 To pxeMaxDtail - 1
            loPic = CType(FindPictureBox(Me, "picDetail" & Format(lnCtr, "00")), PictureBox)
            loPic.BackColor = Color.Transparent
            loPic.BorderStyle = BorderStyle.FixedSingle

            loPic.Tag = ""
            clearDetailImage(lnCtr)
        Next

        If fsCategrCd = "" Then
            psCategrID = ""
        Else
            loDT = p_oTrans.GetDetailImage(fsCategrCd)

            If loDT.Rows.Count > 0 Then
                For lnCtr = 0 To loDT.Rows.Count - 1
                    If lnCtr < pxeMaxDtail Then
                        loPic = CType(FindPictureBox(Me, "picDetail" & Format(lnCtr, "00")), PictureBox)
                        loPic.BackColor = Color.Transparent
                        loPic.BackgroundImage = Nothing
                        loPic.BorderStyle = BorderStyle.FixedSingle
                        loPic.Tag = loDT(lnCtr)("sBriefDsc")

                        loadDetailImages(lnCtr, loDT(lnCtr)("sImgePath"))
                    End If
                Next
                psCategrID = fsCategrCd
            End If
        End If
    End Sub

    Private Sub loadDetailImages(ByVal Index As Integer, ByVal lsDirectory As String)
        Dim loPic As PictureBox

        loPic = CType(FindPictureBox(Me, "picDetail" & Format(Index, "00")), PictureBox)
        loPic.BackgroundImage = Image.FromFile(lsDirectory) 'load from location

    End Sub

    Private Sub loadCategoryImages(ByVal Index As Integer, ByVal lsDirectory As String)
        Dim loPic As PictureBox

        loPic = CType(FindPictureBox(Me, "picCategr" & Format(Index, "00")), PictureBox)
        loPic.BackgroundImage = Image.FromFile(lsDirectory) 'load from location
    End Sub

    Private Sub picCategr06_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picCategr04.MouseDown, picCategr05.MouseDown
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 6
                loPic.BackgroundImage = My.Resources.prev_mouse_over1
            Case 7
                loPic.BackgroundImage = My.Resources.next_mouse_over1
        End Select
    End Sub

    Private Sub picCategr06_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picCategr04.MouseUp, picCategr05.MouseUp
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 6
                loPic.BackgroundImage = My.Resources.prev
            Case 7
                loPic.BackgroundImage = My.Resources._next
        End Select
    End Sub

    Private Sub picDetail55_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picDetail29.MouseDown, picDetail30.MouseDown
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 55
                loPic.BackgroundImage = My.Resources.prev_mouse_over1
            Case 56
                loPic.BackgroundImage = My.Resources.next_mouse_over1
        End Select
    End Sub

    Private Sub picDetail55_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picDetail29.MouseUp, picDetail30.MouseUp
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 55
                loPic.BackgroundImage = My.Resources.prev
            Case 56
                loPic.BackgroundImage = My.Resources._next
        End Select
    End Sub

    Private Sub picTable00_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picTable00.MouseDown, picTable01.MouseDown
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 0
                loPic.BackgroundImage = My.Resources.prev_mouse_over1
            Case 1
                loPic.BackgroundImage = My.Resources.next_mouse_over1
        End Select
    End Sub

    Private Sub picTable00_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picTable00.MouseUp, picTable01.MouseUp
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 0
                loPic.BackgroundImage = My.Resources.prev
            Case 1
                loPic.BackgroundImage = My.Resources._next
        End Select
    End Sub

    Private Sub picButton00_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picButton00.Click, picButton01.Click, picButton02.Click, picButton03.Click, picButton04.Click
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Call procValues(lnIndex)
    End Sub

    Private Sub picButton00_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picButton00.MouseDown, picButton01.MouseDown, picButton02.MouseDown, picButton03.MouseDown, picButton04.MouseDown
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 0
                loPic.BackgroundImage = My.Resources.Mouse_over_x
            Case 1
                loPic.BackgroundImage = My.Resources.Mouse_over_minus
            Case 2
                loPic.BackgroundImage = My.Resources.Mouse_over_plus
            Case 3
                loPic.BackgroundImage = My.Resources.change_price_mouse_ove
            Case 4
                loPic.BackgroundImage = My.Resources.Q_mouse_over
        End Select
    End Sub

    Private Sub picButton00_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        picButton00.MouseHover, picButton01.MouseHover, picButton02.MouseHover, picButton03.MouseHover, picButton04.MouseHover
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))
        With pnToolTip
            Select Case lnIndex
                Case 0
                    .SetToolTip(loPic, "Cancel Order")
                Case 1
                    .SetToolTip(loPic, "Deduct 1 Quantity")
                Case 2
                    .SetToolTip(loPic, "Add 1 Quantity")
                Case 3
                    .SetToolTip(loPic, "Price Update")
                Case 4
                    .SetToolTip(loPic, "QUantity Update")
            End Select
        End With

    End Sub

    Private Sub picButton00_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picButton00.MouseUp, picButton01.MouseUp, picButton02.MouseUp, picButton03.MouseUp, picButton04.MouseUp
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 0
                loPic.BackgroundImage = My.Resources.Mouse_Up_x
            Case 1
                loPic.BackgroundImage = My.Resources.Mouse_Up_minus
            Case 2
                loPic.BackgroundImage = My.Resources.Mouse_Up_plus
            Case 3
                loPic.BackgroundImage = My.Resources.change_price_mouse_up
            Case 4
                loPic.BackgroundImage = My.Resources.Q
        End Select
    End Sub

    Private Sub picScroll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picTable00.Click, picTable01.Click
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 9))

        Select Case lnIndex
            Case 0 'table up
                If pnTablePage = 0 Then Exit Sub

                pnTablePage -= 1
                Call loadTable()
            Case 1 'table down
                If (pnTablePage * pxeMaxTable) + pxeMaxTable > p_oTrans.GetTables.Rows.Count Then Exit Sub

                pnTablePage += 1
                Call loadTable()
        End Select
    End Sub

    Private Sub picCategScroll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picCategr04.Click, picCategr05.Click
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 4 'table up
                If pnCategPage = 0 Then Exit Sub

                pnCategPage -= 1
                Call initCategoryImages()
                Call initDetailImages()
            Case 5 'table down
                If (pnCategPage * pxeMaxCatgr) + pxeMaxCatgr > p_oTrans.GetCategory.Rows.Count Then Exit Sub

                pnCategPage += 1
                Call initCategoryImages()
                Call initDetailImages()
        End Select
    End Sub
#End Region

#Region "Tables"
    Private Sub loadTable()
        Dim loDT As DataTable
        Dim lnCtr As Integer
        Dim lnRow As Integer
        Dim lnTot As Integer
        Dim loLbl As Label

        loDT = p_oTrans.GetTables
        pnTotalTble = loDT.Rows.Count

        Call initTable()

        lnTot = (pnTablePage * pxeMaxTable)
        For lnCtr = 0 + lnTot To lnTot + pxeMaxTable - 1
            lnRow = IIf(pnTablePage = 0, lnCtr, lnCtr - lnTot)

            loLbl = CType(FindLabel(Me, "lblTable" & Format(lnRow, "00")), Label)

            If lnRow < pxeMaxTable Then
                If lnCtr <= loDT.Rows.Count - 1 Then
                    loLbl.Text = loDT(lnCtr)("nTableNox")
                    markTable(lnCtr, CInt(loDT(lnCtr)("cStatusxx")))
                Else
                    markTable(lnCtr, xeTableStatus.xeNONE)
                End If
            End If
        Next
    End Sub

    'this will set table numbers and set them to empty
    Private Sub initTable()
        Dim lnCtr As Integer
        Dim loLbl As Label

        For lnCtr = 0 To pxeMaxTable - 1
            loLbl = CType(FindLabel(Me, "lblTable" & Format(lnCtr, "00")), Label)
            loLbl.Text = ""
            markTable(lnCtr, xeTableStatus.xeNONE)
        Next
    End Sub

    'fsTables is based on label index 0 to max table, not the given table number
    Private Sub markTable(ByVal fsTables As String, ByVal fsStatus As xeTableStatus)
        Dim lvBackColor As Color
        Dim lvForeColor As Color

        Select Case fsStatus
            Case xeTableStatus.xeEmpty
                lvBackColor = Color.White
                lvForeColor = Color.Green
            Case xeTableStatus.xeOccupied
                lvBackColor = Color.Green
                lvForeColor = Color.White
            Case xeTableStatus.xeReserved
                lvBackColor = Color.Yellow
                lvForeColor = Color.White
            Case xeTableStatus.xeDirty
                lvBackColor = Color.Red
                lvForeColor = Color.White
            Case xeTableStatus.xeNONE
                lvBackColor = Color.Black
                lvForeColor = Color.Black
        End Select

        Dim lasDetail() As String = Split(fsTables, ",")
        Dim lnCtr As Integer
        Dim loLbl As Label
        For lnCtr = 0 To UBound(lasDetail)
            loLbl = CType(FindLabel(Me, "lblTable" & Format(CInt(lasDetail(lnCtr)), "00")), Label)

            loLbl.BackColor = lvBackColor
            loLbl.ForeColor = lvForeColor
        Next
    End Sub
#End Region

#Region "TextBox"

    Private Sub txtDetail00_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDetail00.Click
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loTxt.Name, 10))

        With loTxt
            .BackColor = Color.OldLace

            .SelectionStart = 0
            .SelectionLength = .Text.Length
        End With
    End Sub

    Private Sub txtDetail00_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDetail00.GotFocus
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loTxt.Name, 10))

        With loTxt
            .BackColor = Color.OldLace

            .SelectionStart = 0
            .SelectionLength = .Text.Length
        End With
    End Sub

    Private Sub txtDetail_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDetail00.KeyDown
        Dim lsValue As String
        Dim lnQty As Integer

        With p_oTrans
            Select Case e.KeyCode
                Case Keys.Enter
                    lsValue = txtDetail00.Text
                    lnQty = InStr(lsValue, "x", CompareMethod.Text)

                    If lnQty = 0 Then
                        lnQty = 1
                    Else
                        If Not IsNumeric(Strings.Right(lsValue, Len(lsValue) - lnQty)) Then
                            lnQty = 1
                        Else
                            lnQty = CInt(Strings.Right(lsValue, Len(lsValue) - lnQty))

                            lsValue = Strings.Left(lsValue, Len(lsValue) - Len(Str(lnQty)))
                        End If
                    End If

                    .Quantity = lnQty
                    If lsValue.Equals("%") Or lsValue.Equals("") Then Exit Sub
                    If .SearchItem(lsValue, psCategrID) Then loadDetail()
            End Select
        End With
    End Sub

    Private Sub txtDetail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDetail00.LostFocus
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loTxt.Name, 10))

        loTxt.BackColor = SystemColors.Window
    End Sub

    Private Sub txtDetail01_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".")
        End If
    End Sub
#End Region

#Region "Interface Control"
    Private Sub InitializeDataGrid()
        With DataGridView1
            ' Initialize basic DataGridView properties.
            .Dock = DockStyle.Fill
            .BackgroundColor = Color.LightGray
            .BorderStyle = BorderStyle.Fixed3D

            ' Set property values appropriate for read-only display and 
            ' limited interactivity. 
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToOrderColumns = False
            .ReadOnly = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            .AllowUserToResizeColumns = False
            .ColumnHeadersHeightSizeMode = _
                DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .AllowUserToResizeRows = False
            .RowHeadersWidthSizeMode = _
                DataGridViewRowHeadersWidthSizeMode.DisableResizing

            ' Set the selection background color for all the cells.
            .DefaultCellStyle.SelectionBackColor = Color.Empty
            .DefaultCellStyle.SelectionForeColor = Color.Black

            ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default
            ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
            .RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty 'Color.White

            ' Set the background color for all rows and for alternating rows. 
            ' The value for alternating rows overrides the value for all rows. 
            .RowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro

            ' Set the row and column header styles.
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.BackColor = Color.Black
            .RowHeadersDefaultCellStyle.BackColor = Color.Black

            .Font = New Font("Tahoma", 11)
            .RowTemplate.Height = 25
            .ColumnHeadersHeight = 30
        End With

        With DataGridView1.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Navy
            .ForeColor = Color.White
            .Font = New Font(DataGridView1.Font, FontStyle.Bold)
        End With
    End Sub

    Private Sub initGrid()
        With DataGridView1
            .RowCount = 0

            'Set No of Columns
            .ColumnCount = 6

            'Set Column Headers
            .Columns(0).HeaderText = "No"
            .Columns(1).HeaderText = "Barrcode"
            .Columns(2).HeaderText = "Description"
            .Columns(3).HeaderText = "Qty"
            .Columns(4).HeaderText = "SRP"
            .Columns(5).HeaderText = "Total"

            'Set Column Sizes
            .Columns(0).Width = 30
            .Columns(1).Width = 95
            .Columns(2).Width = 160
            .Columns(3).Width = 40
            .Columns(4).Width = 75
            .Columns(5).Width = 80

            .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            .Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            'Set No of Rows
            .RowCount = 1

            pnActiveRow = 0
        End With
    End Sub

    Private Sub initForm(ByVal fvValue As xeEditMode)
        Dim lbShow As Boolean
        Select Case fvValue
            Case xeEditMode.MODE_ADDNEW
                lbShow = True
            Case Else
                lbShow = False
        End Select

        pnlMenu.Visible = lbShow
        pnlTable.Visible = lbShow
        pnlCategory.Visible = lbShow
    End Sub

    Private Sub initPanel(ByVal fnValue As Integer, ByVal fbShow As Boolean)
        Select Case fnValue
            Case 0 'whole body panel
                Panel2.Visible = fbShow
            Case 1 'right side of body section
                pnlCategory.Visible = fbShow
                pnlMenu.Visible = fbShow
                pnlTable.Visible = fbShow
            Case 2 'left side of body section
            Case 3 'buttons section
        End Select
    End Sub

    Private Sub clearFields()
        lblCurTimex.Text = Format(p_oAppDriver.SysDate, "hh:mm:ss tt")
        lblCurTimex.Refresh()
        lblCurDatex.Text = Format(p_oAppDriver.SysDate, "MMM-dd-yyyy")
        lblCurDatex.Refresh()
        lblCurDay.Text = Format(p_oAppDriver.SysDate, "dddd") & ","
        lblCurDay.Refresh()

        lblAmount.Text = "0.00"
        lblMaster04.Text = "0.00"
        lblMaster13.Text = "0.00"
        lblMaster14.Text = "0.00"
        lblMaster15.Text = "0.00"
        lblMaster17.Text = "0.00"
        'lblNetSale.Text = "0.00"
        lblSrvCrge.Text = "0.00"

        lblMaster00.Text = ""
        lblMaster05.Text = ""
        lblMaster80.Text = ""
        lblMaster81.Text = ""

        Call clearSeek()
    End Sub

    Private Sub clearSeek()
        txtDetail00.Text = ""
        txtDetail00.Focus()
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub PreventFlicker()
        With Me
            .SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            .SetStyle(ControlStyles.UserPaint, True)
            .SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            .UpdateStyles()
        End With
    End Sub
#End Region

    Private Sub p_oTrans_DisplayFromRow(ByVal Row As Integer) Handles p_oTrans.DisplayFromRow
        If p_oTrans.ItemCount = 0 Then
            Call initGrid()
        Else
            Call loadDetail(Row)
        End If
    End Sub

    Private Sub p_oTrans_DisplayRow(ByVal Row As Integer) Handles p_oTrans.DisplayRow
        With DataGridView1
            If .Rows.Count < p_oTrans.ItemCount Then
                .RowCount = .RowCount + (p_oTrans.ItemCount - .Rows.Count)
            End If

            If .RowCount > 14 Then
                .Columns(2).Width = 160
                .FirstDisplayedScrollingRowIndex = .RowCount - 14
            Else
                .Columns(2).Width = 160
            End If

            If Trim(p_oTrans.Detail(Row, "sStockIDx")) <> "" Then
                .Item(0, Row).Value = Row + 1
                .Item(1, Row).Value = p_oTrans.Detail(Row, "sBarcodex")
                .Item(2, Row).Value = p_oTrans.Detail(Row, "sBriefDsc")
                .Item(3, Row).Value = IIf(p_oTrans.Detail(Row, "cReversex") = "-", _
                                              p_oTrans.Detail(Row, "cReversex"), "") & p_oTrans.Detail(Row, "nQuantity")
                .Item(4, Row).Value = Format(p_oTrans.Detail(Row, "nUnitPrce") * _
                                                (100 - p_oTrans.Detail(Row, "nDiscount")) / 100 -
                                                    p_oTrans.Detail(Row, "nAddDiscx"), "#,##0.00")

                .Item(5, Row).Value = Format(.Item(3, Row).Value * .Item(4, Row).Value, "#,##0.00")

                'set active row
                pnActiveRow = Row
                .ClearSelection()
                .Rows(pnActiveRow).Selected = True
            End If
        End With
    End Sub

    Private Sub picCategr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picCategr00.Click, picCategr01.Click,
        picCategr02.Click, picCategr03.Click

        Dim loLbl As PictureBox

        loLbl = CType(sender, System.Windows.Forms.PictureBox)

        initDetailImages(loLbl.Tag)

        If loLbl.Tag <> "" Then txtDetail00.Text = ""
    End Sub

    Private Sub picDetailImages_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picDetail29.Click, picDetail30.Click
        Dim loPic As PictureBox
        loPic = CType(sender, System.Windows.Forms.PictureBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loPic.Name, 10))

        Select Case lnIndex
            Case 54 'table up
                If pnDtailPage = 0 Then Exit Sub

                pnDtailPage -= 1
                Call initDetailImages(psCategrID)
            Case 55 'table down
                If (pnDtailPage * pxeMaxDtail) + pxeMaxDtail > p_oTrans.GetDetailImage(psCategrID).Rows.Count Then Exit Sub

                pnDtailPage += 1
                Call initDetailImages(psCategrID)
        End Select
    End Sub

    Private Sub picDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        picDetail00.Click, picDetail01.Click, picDetail02.Click, picDetail03.Click, picDetail04.Click,
        picDetail05.Click, picDetail06.Click, picDetail07.Click,
        picDetail08.Click, picDetail09.Click, picDetail10.Click,
        picDetail11.Click, picDetail12.Click, picDetail20.Click, picDetail21.Click,
        picDetail22.Click, picDetail23.Click, picDetail13.Click, picDetail14.Click,
        picDetail15.Click, picDetail16.Click, picDetail17.Click, picDetail18.Click,
        picDetail19.Click, picDetail24.Click,
        picDetail25.Click, picDetail26.Click, picDetail27.Click, picDetail28.Click

        Dim loLbl As PictureBox

        loLbl = CType(sender, System.Windows.Forms.PictureBox)
        If loLbl.Tag <> "" Then
            With txtDetail00
                .Text = loLbl.Tag
                .SelectionStart = .Text.Length
                .SelectionLength = .Text.Length
            End With
        End If
    End Sub


    Private Sub p_oTrans_MasterRetreived(ByVal Index As Integer, ByVal Value As Object) Handles p_oTrans.MasterRetreived
        Select Case Index
            Case 0
                lblMaster00.Text = Strings.Right(Value, 16)
        End Select
    End Sub

    Private Sub picDetail00_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        picDetail00.MouseHover, picDetail01.MouseHover, picDetail02.MouseHover, picDetail03.MouseHover, picDetail04.MouseHover,
        picDetail05.MouseHover, picDetail06.MouseHover, picDetail07.MouseHover,
        picDetail08.MouseHover, picDetail09.MouseHover, picDetail10.MouseHover,
        picDetail11.MouseHover, picDetail12.MouseHover, picDetail13.MouseHover, picDetail14.MouseHover,
        picDetail15.MouseHover, picDetail16.MouseHover, picDetail17.MouseHover, picDetail18.MouseHover,
        picDetail19.MouseHover, picDetail20.MouseHover, picDetail21.MouseHover,
        picDetail22.MouseHover, picDetail23.MouseHover, picDetail24.MouseHover,
        picDetail25.MouseHover, picDetail26.MouseHover, picDetail27.MouseHover, picDetail28.MouseHover, picDetail29.MouseHover,
        picDetail30.MouseHover

        Dim loLbl As PictureBox

        loLbl = CType(sender, System.Windows.Forms.PictureBox)
        Dim loIndex As Integer

        loIndex = Val(Mid(loLbl.Name, 10))

        If Mid(loLbl.Name, 1, 9) = "picDetail" Then
            If loLbl.Tag <> "" Then
                With pnToolTip
                    Select Case loIndex
                        Case 0 To pxeMaxDtail
                            .SetToolTip(loLbl, loLbl.Tag)
                    End Select
                End With
            Else
                Select Case loIndex
                    Case 29
                        pnToolTip.SetToolTip(loLbl, "Previous")
                    Case 30
                        pnToolTip.SetToolTip(loLbl, "Next")
                End Select
            End If

        End If

    End Sub

    Private Sub picCategr00_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        picCategr00.MouseHover, picCategr01.MouseHover, picCategr02.MouseHover, picCategr03.MouseHover,
         picCategr04.MouseHover, picCategr05.MouseHover

        Dim loLbl As PictureBox

        loLbl = CType(sender, System.Windows.Forms.PictureBox)
        Dim loIndex As Integer

        loIndex = Val(Mid(loLbl.Name, 10))

        If Mid(loLbl.Name, 1, 9) = "picCategr" Then
            If loLbl.Tag <> "" Then
                With pnToolTip
                    Select Case loIndex
                        Case 0 To pxeMaxDtail
                            .SetToolTip(loLbl, loadCateg(loLbl.Tag))
                    End Select
                End With
            Else
                Select Case loIndex
                    Case 4
                        pnToolTip.SetToolTip(loLbl, "Previous")
                    Case 5
                        pnToolTip.SetToolTip(loLbl, "Next")
                End Select

            End If

        End If
    End Sub

    Public Function loadCateg(ByVal fsCategrCd As String) As String
        Dim lsSQL As String
        Dim loDT As DataTable
        Dim result As String

        lsSQL = "SELECT sDescript FROM Product_Category WHERE sCategrCd =" + fsCategrCd + ";"

        loDT = p_oAppDriver.ExecuteQuery(lsSQL)

        If loDT.Rows.Count > 0 Then
            result = loDT(0)("sDescript")
        End If

        Return result
    End Function

    Private Sub cmdButton00_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        cmdButton00.MouseHover, cmdButton01.MouseHover, cmdButton02.MouseHover, cmdButton03.MouseHover, cmdButton04.MouseHover,
        cmdButton05.MouseHover, cmdButton07.MouseHover, cmdButton08.MouseHover, cmdButton09.MouseHover,
        cmdButton10.MouseHover, cmdButton11.MouseHover, cmdButton12.MouseHover, cmdButton13.MouseHover, cmdButton14.MouseHover

        Dim cmdButton As Button

        cmdButton = CType(sender, System.Windows.Forms.Button)
        Dim loIndex As Integer

        loIndex = Val(Mid(cmdButton.Name, 10))

        If Mid(cmdButton.Name, 1, 9) = "cmdButton" Then
            With pnToolTip
                Select Case loIndex
                    Case 0
                        .SetToolTip(cmdButton, "Logout\n Escape")
                    Case 1
                        .SetToolTip(cmdButton, "Print Z- Reading")
                    Case 2
                        .SetToolTip(cmdButton, "Save Order")
                    Case 3
                        .SetToolTip(cmdButton, "Merge Order")
                    Case 4
                        .SetToolTip(cmdButton, "Pay")
                    Case 5
                        .SetToolTip(cmdButton, "Split")
                    Case 6
                        .SetToolTip(cmdButton, "Charge")
                    Case 7
                        .SetToolTip(cmdButton, "Complement")
                    Case 8
                        .SetToolTip(cmdButton, "Discount")
                    Case 9
                        .SetToolTip(cmdButton, "Sales Return")
                    Case 10
                        .SetToolTip(cmdButton, "Cash Pull Out")
                    Case 11
                        .SetToolTip(cmdButton, "Void Order")
                    Case 12
                        .SetToolTip(cmdButton, "Print Bill")
                    Case 13
                        .SetToolTip(cmdButton, "End Shift")
                    Case 14
                        .SetToolTip(cmdButton, "Cancel O.R.")
                    Case 15
                        .SetToolTip(cmdButton, "Pay Charge")
                End Select
            End With

        End If
    End Sub
End Class
