Imports ggcAppDriver
Imports ggcRetailSales
Imports System.Reflection

' PRINTER NAME USED IN OUR PRINTING...
'        RMS_PRN_CS = CASHIER PRINTER
'                   = Ex. \\192.168.10.14\EPSON LX-310 ESC/P
'        RMP_PRN_TK = ORDER TAKER PRINTER
'                   = Ex. \\192.168.10.14\EPSON LX-310 ESC/P
'        RMS_PRN_KN = KITCHEN PRINTER
'COMPANY INFORMATION
'        p_sPOSNo  = MIN
'        p_sVATReg = VAT REG No.
'        p_sCompny = Mother Company of the Branch

Module modMain
    Public p_oAppDriver As GRider
    Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As UInteger, ByVal dwExtraInfo As UInteger)

    Private Const xsSignature As String = "08220326"
    Public pbValidSales As Boolean
    Public pnSaleStat As Integer
    Public psCashierx As String = ""
    Public p_sLogName As String = ""
    Public pdPOSDatex As Date
    Public poSales As DailySales
    Enum xeTableStatus
        xeEmpty = 0
        xeOccupied = 1
        xeReserved = 2
        xeDirty = 3
        xeNONE = 4
    End Enum

    Public Sub Main(ByVal args As String())
        'Enable XP visual style/skin
        Application.EnableVisualStyles()

        Dim lsUserIDxx As String

        p_oAppDriver = New GRider("RetMgtSys")
        If Not p_oAppDriver.LoadEnv() Then
            MsgBox("Unable to load configuration file!")
            Exit Sub
        End If

        lsUserIDxx = ""
        If args.Length = 2 Then
            lsUserIDxx = args(1)
            If Not p_oAppDriver.LogUser(lsUserIDxx) Then
                MsgBox("User unable to log!")
                Exit Sub
            End If
        Else
            If Not p_oAppDriver.LogUser() Then
                MsgBox("User unable to log!")
                Exit Sub
            Else
                p_sLogName = Decrypt(p_oAppDriver.LogName, xsSignature)
            End If
        End If

        'lsUserIDxx = "M001180004"
        'If Not p_oAppDriver.LogUser(lsUserIDxx) Then
        '    MsgBox("User unable to log!")
        '    Exit Sub
        'End If

        '0->With Open Sales Order From Previous Sale;
        '1->Sales for the Day was already closed;
        '2->Sales for the Day is Ok;
        '3->Error Printing TXReading 
        '4->User is not allowed to enter Sales Transaction
        pnSaleStat = procDailySales()

        Select Case pnSaleStat
            Case -1, 1, 3, 4
                'If pc is not registered then do not load the program
                Exit Sub
            Case Else
                pbValidSales = pnSaleStat = 2
        End Select

        frmMain.ShowDialog()
    End Sub

    Public Function getCashier(ByVal sCashierx As String) As String
        Dim lsSQL As String
        Dim lsCashierNm As String
        Dim loDta As DataTable

        lsSQL = "SELECT" & _
                    " a.sUserName" & _
                    " FROM xxxSysUser a" & _
                    " WHERE a.sUserIDxx = " & strParm(sCashierx)

        loDta = p_oAppDriver.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            lsCashierNm = ""
        Else
            lsCashierNm = Decrypt(loDta(0).Item("sUserName"), xsSignature)
        End If

        loDta = Nothing

        Return lsCashierNm

    End Function

    Private Function procDailySales() As Integer
        Try
            poSales = New DailySales(p_oAppDriver)

            If Not poSales.initMachine Then
                MsgBox("Work Station is not Registered.", MsgBoxStyle.Exclamation, "Warning")
                Return -1
            End If

            poSales.Cashier = p_oAppDriver.UserID

            If Not poSales.NewTransaction Then
                pdPOSDatex = poSales.POSDate
                Return poSales.SalesStatus
            End If

            pdPOSDatex = poSales.POSDate
            psCashierx = poSales.Cashier

            If poSales.EditMode = xeEditMode.MODE_READY Then Return poSales.SalesStatus

            Dim frmDailySales As frmDailySales
            frmDailySales = New frmDailySales(p_oAppDriver)
            With frmDailySales
                .Sales = poSales
                .TopMost = True
                .ShowDialog()
                If .Cancel Then Return 4

                poSales.Master("nOpenBalx") = .InitialCash
                poSales.Master("nCPullOut") = .PullOut
                poSales.Master("dOpenedxx") = p_oAppDriver.SysDate

                poSales.SaveTransaction()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
            Return -1
        End Try

        Return poSales.SalesStatus
    End Function

    'Private Function procDailySales() As Boolean
    '    Dim frmDailySales As frmDailySales
    '    frmDailySales = New frmDailySales(p_oAppDriver)
    '    With frmDailySales
    '        If .EditMode = xeEditMode.MODE_READY Then
    '            'Daily sales record was already created for this cashier....
    '            .Close()
    '            frmDailySales = Nothing
    '            Return True
    '        End If

    '        .TopMost = True
    '        .ShowDialog()

    '        If .Cancel Then Return False
    '    End With

    '    Return True
    'End Function

    'This method can handle all events using EventHandler
    Public Sub grpEventHandler(ByVal foParent As Control, ByVal foType As Type, ByVal fsGroupNme As String, ByVal fsEvent As String, ByVal foAddress As EventHandler)
        Dim loTxt As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = foType Then
                'Handle events for this controls only
                If LCase(Mid(loTxt.Name, 1, Len(fsGroupNme))) = LCase(fsGroupNme) Then
                    If foType = GetType(TextBox) Then
                        Dim loObj = DirectCast(loTxt, TextBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(CheckBox) Then
                        Dim loObj = DirectCast(loTxt, CheckBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Button) Then
                        Dim loObj = DirectCast(loTxt, Button)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Label) Then
                        Dim loObj = DirectCast(loTxt, Label)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(PictureBox) Then
                        Dim loObj = DirectCast(loTxt, PictureBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    End If
                End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
            Else
                If loTxt.HasChildren Then
                    Call grpEventHandler(loTxt, foType, fsGroupNme, fsEvent, foAddress)
                End If
            End If
        Next 'loTxt In loControl.Controls
    End Sub

    'This method can handle all events using CancelEventHandler
    Public Sub grpCancelHandler(ByVal foParent As Control, ByVal foType As Type, ByVal fsGroupNme As String, ByVal fsEvent As String, ByVal foAddress As System.ComponentModel.CancelEventHandler)
        Dim loTxt As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = foType Then
                'Handle events for this controls only
                If LCase(Mid(loTxt.Name, 1, Len(fsGroupNme))) = LCase(fsGroupNme) Then
                    If foType = GetType(TextBox) Then
                        Dim loObj = DirectCast(loTxt, TextBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(CheckBox) Then
                        Dim loObj = DirectCast(loTxt, CheckBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Button) Then
                        Dim loObj = DirectCast(loTxt, Button)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    End If
                End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
            Else
                If loTxt.HasChildren Then
                    Call grpCancelHandler(loTxt, foType, fsGroupNme, fsEvent, foAddress)
                End If
            End If
        Next 'loTxt In loControl.Controls
    End Sub

    'This method can handle all events using KeyEventHandler
    Public Sub grpKeyHandler(ByVal foParent As Control, ByVal foType As Type, ByVal fsGroupNme As String, ByVal fsEvent As String, ByVal foAddress As KeyEventHandler)
        Dim loTxt As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = foType Then
                'Handle events for this controls only
                If LCase(Mid(loTxt.Name, 1, Len(fsGroupNme))) = LCase(fsGroupNme) Then
                    If foType = GetType(TextBox) Then
                        Dim loObj = DirectCast(loTxt, TextBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(CheckBox) Then
                        Dim loObj = DirectCast(loTxt, CheckBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Button) Then
                        Dim loObj = DirectCast(loTxt, Button)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    End If
                End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
            Else
                If loTxt.HasChildren Then
                    Call grpKeyHandler(loTxt, foType, fsGroupNme, fsEvent, foAddress)
                End If
            End If
        Next 'loTxt In loControl.Controls
    End Sub

    'This method can handle all events using EventHandler
    Public Function FindPictureBox(ByVal foParent As Control, ByVal fsName As String) As Control
        Dim loTxt As Control
        Static loRet As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = GetType(PictureBox) Then
                'Handle events for this controls only
                If LCase(loTxt.Name) = LCase(fsName) Then
                    loRet = loTxt
                End If
            Else
                If loTxt.HasChildren Then
                    Call FindPictureBox(loTxt, fsName)
                End If
            End If
        Next 'loTxt In loControl.Controls

        Return loRet
    End Function

    'This method can handle all events using EventHandler
    Public Function FindTextBox(ByVal foParent As Control, ByVal fsName As String) As Control
        Dim loTxt As Control
        Static loRet As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = GetType(TextBox) Then
                'Handle events for this controls only
                If LCase(loTxt.Name) = LCase(fsName) Then
                    loRet = loTxt
                End If
            Else
                If loTxt.HasChildren Then
                    Call FindTextBox(loTxt, fsName)
                End If
            End If
        Next 'loTxt In loControl.Controls

        Return loRet
    End Function

    'This method can handle all events using EventHandler
    Public Function FindLabel(ByVal foParent As Control, ByVal fsName As String) As Control
        Dim loTxt As Control
        Static loRet As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = GetType(Label) Then
                'Handle events for this controls only
                If LCase(loTxt.Name) = LCase(fsName) Then
                    loRet = loTxt
                End If
            Else
                If loTxt.HasChildren Then
                    Call FindLabel(loTxt, fsName)
                End If
            End If
        Next 'loTxt In loControl.Controls

        Return loRet
    End Function

    Public Sub SetNextFocus()
        keybd_event(&H9, 0, 0, 0)
        keybd_event(&H9, 0, &H2, 0)
    End Sub

    Public Sub SetPreviousFocus()
        keybd_event(&H10, 0, 0, 0)
        keybd_event(&H9, 0, 0, 0)
        keybd_event(&H10, 0, &H2, 0)
    End Sub
End Module
