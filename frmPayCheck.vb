Public Class frmPayCheck
    Private pnLoadx As Integer
    Private poControl As Control

    Private Sub frmPayGC_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                showForm(1, True)
        End Select
    End Sub

    Private Sub frmPay_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        setVisible(False)

        If pnLoadx = 0 Then
            showDetail(False)
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
            Case 0
                showForm(1, True)
            Case 1, 2, 4
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
        Dim lvFormOrgxx As New Size(356, 495)
        Dim lvFormNewxx As New Size(356, 360)

        If lbShow Then
            Me.Size = lvFormOrgxx
            pnlDetail.Visible = True
            pnlMain.Size = lvMPnelOrgx
            pnlButtons.Location = lvButtonLoc
        Else
            Me.Size = lvFormNewxx
            pnlDetail.Visible = False
            pnlMain.Size = lvMPnelNewx
            pnlButtons.Location = lvDetailLoc
        End If
    End Sub

    Private Sub setVisible(ByVal lbShow As Boolean)
        If lbShow Then
            Me.BackgroundImage = My.Resources.BACKGROUND_1
            Me.BackgroundImageLayout = ImageLayout.Stretch

            pnlBill.Visible = True
            pnlMain.Visible = True

            Me.Opacity = 100
            txtField00.Focus()
        Else
            Me.Opacity = 0
            Me.BackColor = Color.Orange
            Me.TransparencyKey = Me.BackColor
            Me.TopMost = False

            cmdButton03.Enabled = False
        End If
    End Sub

    Private Sub clearFields()
        lblBill.Text = Format(0.0#, "#,##0.00")
        txtField05.Text = Format(0.0#, "#,##0.00")
        txtField00.Text = ""
        txtField01.Text = ""
        txtField02.Text = ""
        txtField03.Text = ""
        txtField04.Text = ""
    End Sub
End Class