Imports System.Threading
Public Class frmPay
    Private pnLoadx As Integer
    Private poControl As Control

    Private Sub frmPayGC_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                showForm(1, False)
                showForm(2, False)
                showForm(3, False)
                showForm(4, False)
        End Select
    End Sub

    Private Sub frmPay_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        setVisible()
        If pnLoadx = 0 Then
            showDetail(False)
            clearFields()
            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)
            pnLoadx = 1
        End If
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 'OK
                'pay code here
            Case 1 'CASH
            Case 2, 3, 4
                showForm(lnIndex, True)
        End Select
endProc:
        Exit Sub
    End Sub

    Private Sub showDetail(ByVal lbShow As Boolean)
        Dim lvDetailLoc As New Point(1, 199)
        Dim lvButtonLoc As New Point(1, 333)
        Dim lvMPnelOrgx As New Size(344, 377)
        Dim lvMPnelNewx As New Size(344, 243)
        Dim lvFormOrgxx As New Size(628, 645)
        Dim lvFormNewxx As New Size(628, 645)

        'If lbShow Then
        '    Me.Size = lvFormOrgxx
        '    pnlDetail.Visible = True
        '    pnlMain.Size = lvMPnelOrgx
        '    pnlButtons.Location = lvButtonLoc
        'Else
        '    Me.Size = lvFormNewxx
        '    pnlDetail.Visible = False
        '    pnlMain.Size = lvMPnelNewx
        '    pnlButtons.Location = lvDetailLoc
        'End If
    End Sub

    Private Sub setVisible()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        Me.Visible = False
        Me.BackColor = Color.Transparent
        Me.Location = New Point(3, 3)
        txtAmount.Focus()
        Me.Visible = True
    End Sub

    Private Sub clearFields()
        lblBill.Text = Format(0.0#, "#,##0.00")
        lblChange.Text = Format(0.0#, "#,##0.00")
        txtAmount.Text = Format(0.0#, "#,##0.00")
    End Sub
End Class