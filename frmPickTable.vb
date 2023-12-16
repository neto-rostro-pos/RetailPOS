Imports System.Threading
Public Class frmPickTable
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
            pnLoadx = 1
        End If
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
            Case 0 'OK
                'pay code here
            Case 1 'CASH
            Case 2, 3, 4
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
            txtField00.Focus()
        Else
            Me.Opacity = 0
            Me.BackColor = Color.Orange
            Me.TransparencyKey = Me.BackColor
            Me.TopMost = False
        End If
    End Sub

    Private Sub clearFields()
        txtField00.Text = ""
        txtField01.Text = ""
        txtField02.Text = ""
    End Sub
End Class