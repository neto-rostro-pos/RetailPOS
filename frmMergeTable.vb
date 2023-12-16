Imports System.Threading
Public Class frmMergeTable
    Private Const pxeOccupiedTag As String = "Y"
    Private Const pxeUnoccupdTag As String = "N"
    Private Const pxeGrpdTbleTag As String = "G"
    Private Const pxeAddedTblTag As String = "A"

    Private pnLoadx As Integer
    Private poControl As Control

    Private Sub frmPayGC_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

    Private Sub frmPay_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        setVisible(False)

        If pnLoadx = 0 Then
            clearFields()
            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)
            Call grpEventHandler(Me, GetType(Label), "lblTable", "Click", AddressOf lblTable_Click)

            pnLoadx = 1
        End If

        'pass accoupied table nos separated by »
        'Call Occupied(ByVal lsValue As String, ByVal lbOccupied As Boolean, Optional ByVal lsRefVal As String = "")
        Call Occupied("1»2»3»4»5", True, "000001")
    End Sub

    Private Sub frmPay_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        setVisible(True)
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 'SAVE
            Case 1 'CANCEL
        End Select
endProc:
        Exit Sub
    End Sub

    Private Sub setVisible(ByVal lbShow As Boolean)
        If lbShow Then
            Me.BackgroundImage = My.Resources.BACKGROUND_1
            Me.BackgroundImageLayout = ImageLayout.Stretch

            pnlMain.Visible = True

            Me.Opacity = 100
        Else
            Me.Opacity = 0
            Me.BackColor = Color.Orange
            Me.TransparencyKey = Me.BackColor
            Me.TopMost = False
        End If
    End Sub

    Private Sub showGroup(ByVal lsValue As String)
        Dim lasValue() As String
        Dim lnCtr As Integer
        Dim loLbl As Label

        lasValue = Split(lsValue, "»")

        For lnCtr = 0 To UBound(lasValue)
            loLbl = CType(FindLabel(Me, "lblTable" & Format(Int(lasValue(lnCtr)), "00")), Label)
            loLbl.BackColor = Color.Green
            loLbl.ForeColor = Color.White
            loLbl.Tag = pxeGrpdTbleTag
        Next
    End Sub

    Private Sub addToGroup(ByVal lnIndex As Integer)
        Dim loLbl As Label
        Dim loRef As Label

        loLbl = CType(FindLabel(Me, "lblTable" & Format(lnIndex, "00")), Label)
        loRef = CType(FindLabel(Me, "lblRefer" & Format(lnIndex, "00")), Label)

        If loLbl.Tag <> pxeOccupiedTag Then
            If loLbl.Tag <> pxeAddedTblTag Then
                loLbl.BackColor = Color.Green
                loLbl.ForeColor = Color.White
                loLbl.Tag = pxeAddedTblTag

                loRef.BackColor = Color.Green
                loRef.ForeColor = Color.White
                loRef.Text = "ADDED"
            Else
                Occupied(lnIndex, False)
            End If
        End If
    End Sub

    Private Sub Occupied(ByVal lsValue As String, ByVal lbOccupied As Boolean, Optional ByVal lsRefVal As String = "")
        Dim lasValue() As String
        Dim loLbl As Label
        Dim loRef As Label
        Dim lnCtr As Integer

        lasValue = Split(lsValue, "»")

        For lnCtr = 0 To UBound(lasValue)
            loLbl = CType(FindLabel(Me, "lblTable" & Format(Int(lasValue(lnCtr)), "00")), Label)
            loRef = CType(FindLabel(Me, "lblRefer" & Format(Int(lasValue(lnCtr)), "00")), Label)

            If lbOccupied Then
                loLbl.BackColor = Color.Gray
                loLbl.ForeColor = Color.White
                loLbl.Tag = pxeOccupiedTag

                loRef.BackColor = Color.Gray
                loRef.ForeColor = Color.White
            Else
                loLbl.BackColor = Color.White
                loLbl.ForeColor = Color.Green
                loLbl.Tag = pxeUnoccupdTag

                loRef.BackColor = Color.White
                loRef.ForeColor = Color.Green
            End If

            loRef.Text = lsRefVal
        Next
    End Sub

    Private Sub Occupied(ByVal lnIndex As Integer, ByVal lbOccupied As Boolean, Optional ByVal lsRefVal As String = "")
        Dim loLbl As Label
        Dim loRef As Label

        loLbl = CType(FindLabel(Me, "lblTable" & Format(lnIndex, "00")), Label)
        loRef = CType(FindLabel(Me, "lblRefer" & Format(lnIndex, "00")), Label)

        If lbOccupied Then
            loLbl.BackColor = Color.Gray
            loLbl.ForeColor = Color.White
            loLbl.Tag = pxeOccupiedTag

            loRef.BackColor = Color.Gray
            loRef.ForeColor = Color.White
        Else
            loLbl.BackColor = Color.White
            loLbl.ForeColor = Color.Green
            loLbl.Tag = pxeUnoccupdTag

            loRef.BackColor = Color.White
            loRef.ForeColor = Color.Green
        End If

        loRef.Text = lsRefVal
    End Sub

    Private Sub clearFields()
        Dim lnCtr As Integer

        For lnCtr = 0 To 31
            Occupied(lnCtr, False)
        Next
    End Sub

    Private Sub lblTable_Click(sender As System.Object, e As System.EventArgs)
        Dim loLbl As Label
        loLbl = CType(sender, System.Windows.Forms.Label)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loLbl.Name, 9))

        'upon click check if the table was merged to others
        'pass value separated by »
        'Call showGroup(ByVal lsValue As String)

        'higlight table no
        addToGroup(lnIndex)
    End Sub
End Class