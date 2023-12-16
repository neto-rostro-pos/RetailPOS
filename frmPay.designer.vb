<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPay
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPay))
        Me.lblBill = New System.Windows.Forms.Label()
        Me.pnlBill = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblChange = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.cmdButton03 = New System.Windows.Forms.Button()
        Me.cmdButton00 = New System.Windows.Forms.Button()
        Me.cmdButton01 = New System.Windows.Forms.Button()
        Me.cmdButton04 = New System.Windows.Forms.Button()
        Me.cmdButton02 = New System.Windows.Forms.Button()
        Me.pnlDetail = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.pnlAmount = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.pnlBill.SuspendLayout()
        Me.pnlMain.SuspendLayout()
        Me.pnlButtons.SuspendLayout()
        Me.pnlDetail.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAmount.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblBill
        '
        Me.lblBill.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblBill.Font = New System.Drawing.Font("Microsoft Sans Serif", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBill.ForeColor = System.Drawing.Color.White
        Me.lblBill.Location = New System.Drawing.Point(3, 4)
        Me.lblBill.Name = "lblBill"
        Me.lblBill.Size = New System.Drawing.Size(603, 48)
        Me.lblBill.TabIndex = 1
        Me.lblBill.Text = "00,000.00"
        Me.lblBill.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlBill
        '
        Me.pnlBill.BackColor = System.Drawing.Color.Transparent
        Me.pnlBill.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlBill.Controls.Add(Me.Label3)
        Me.pnlBill.Controls.Add(Me.lblChange)
        Me.pnlBill.Controls.Add(Me.Label2)
        Me.pnlBill.Controls.Add(Me.lblBill)
        Me.pnlBill.Location = New System.Drawing.Point(5, 23)
        Me.pnlBill.Name = "pnlBill"
        Me.pnlBill.Size = New System.Drawing.Size(611, 91)
        Me.pnlBill.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(42, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "CHANGE"
        '
        'lblChange
        '
        Me.lblChange.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblChange.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChange.ForeColor = System.Drawing.Color.Red
        Me.lblChange.Location = New System.Drawing.Point(3, 52)
        Me.lblChange.Name = "lblChange"
        Me.lblChange.Size = New System.Drawing.Size(603, 33)
        Me.lblChange.TabIndex = 3
        Me.lblChange.Text = "00,000.00"
        Me.lblChange.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label2.Location = New System.Drawing.Point(3, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "TOTAL BILL"
        '
        'pnlMain
        '
        Me.pnlMain.BackColor = System.Drawing.Color.Transparent
        Me.pnlMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlMain.Controls.Add(Me.pnlButtons)
        Me.pnlMain.Controls.Add(Me.pnlDetail)
        Me.pnlMain.Controls.Add(Me.pnlAmount)
        Me.pnlMain.Location = New System.Drawing.Point(5, 117)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(611, 516)
        Me.pnlMain.TabIndex = 2
        '
        'pnlButtons
        '
        Me.pnlButtons.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlButtons.Controls.Add(Me.cmdButton03)
        Me.pnlButtons.Controls.Add(Me.cmdButton00)
        Me.pnlButtons.Controls.Add(Me.cmdButton01)
        Me.pnlButtons.Controls.Add(Me.cmdButton04)
        Me.pnlButtons.Controls.Add(Me.cmdButton02)
        Me.pnlButtons.Location = New System.Drawing.Point(3, 470)
        Me.pnlButtons.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Size = New System.Drawing.Size(603, 40)
        Me.pnlButtons.TabIndex = 7
        '
        'cmdButton03
        '
        Me.cmdButton03.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.cmdButton03.FlatAppearance.BorderSize = 0
        Me.cmdButton03.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton03.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton03.ForeColor = System.Drawing.Color.White
        Me.cmdButton03.Location = New System.Drawing.Point(468, 3)
        Me.cmdButton03.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdButton03.Name = "cmdButton03"
        Me.cmdButton03.Size = New System.Drawing.Size(64, 32)
        Me.cmdButton03.TabIndex = 10
        Me.cmdButton03.Text = "CHECK"
        Me.cmdButton03.UseVisualStyleBackColor = False
        '
        'cmdButton00
        '
        Me.cmdButton00.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.cmdButton00.FlatAppearance.BorderSize = 0
        Me.cmdButton00.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton00.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton00.ForeColor = System.Drawing.Color.White
        Me.cmdButton00.Location = New System.Drawing.Point(270, 3)
        Me.cmdButton00.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdButton00.Name = "cmdButton00"
        Me.cmdButton00.Size = New System.Drawing.Size(64, 32)
        Me.cmdButton00.TabIndex = 7
        Me.cmdButton00.Text = "OK"
        Me.cmdButton00.UseVisualStyleBackColor = False
        '
        'cmdButton01
        '
        Me.cmdButton01.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.cmdButton01.FlatAppearance.BorderSize = 0
        Me.cmdButton01.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton01.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton01.ForeColor = System.Drawing.Color.White
        Me.cmdButton01.Location = New System.Drawing.Point(336, 3)
        Me.cmdButton01.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdButton01.Name = "cmdButton01"
        Me.cmdButton01.Size = New System.Drawing.Size(64, 32)
        Me.cmdButton01.TabIndex = 8
        Me.cmdButton01.Text = "CASH"
        Me.cmdButton01.UseVisualStyleBackColor = False
        '
        'cmdButton04
        '
        Me.cmdButton04.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.cmdButton04.FlatAppearance.BorderSize = 0
        Me.cmdButton04.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton04.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton04.ForeColor = System.Drawing.Color.White
        Me.cmdButton04.Location = New System.Drawing.Point(534, 3)
        Me.cmdButton04.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdButton04.Name = "cmdButton04"
        Me.cmdButton04.Size = New System.Drawing.Size(64, 32)
        Me.cmdButton04.TabIndex = 11
        Me.cmdButton04.Text = "GC"
        Me.cmdButton04.UseVisualStyleBackColor = False
        '
        'cmdButton02
        '
        Me.cmdButton02.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.cmdButton02.FlatAppearance.BorderSize = 0
        Me.cmdButton02.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton02.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton02.ForeColor = System.Drawing.Color.White
        Me.cmdButton02.Location = New System.Drawing.Point(402, 3)
        Me.cmdButton02.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdButton02.Name = "cmdButton02"
        Me.cmdButton02.Size = New System.Drawing.Size(64, 32)
        Me.cmdButton02.TabIndex = 9
        Me.cmdButton02.Text = "CREDIT"
        Me.cmdButton02.UseVisualStyleBackColor = False
        '
        'pnlDetail
        '
        Me.pnlDetail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlDetail.Controls.Add(Me.DataGridView1)
        Me.pnlDetail.Location = New System.Drawing.Point(3, 280)
        Me.pnlDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlDetail.Name = "pnlDetail"
        Me.pnlDetail.Size = New System.Drawing.Size(603, 187)
        Me.pnlDetail.TabIndex = 6
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(595, 179)
        Me.DataGridView1.TabIndex = 21
        '
        'pnlAmount
        '
        Me.pnlAmount.BackColor = System.Drawing.Color.Transparent
        Me.pnlAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlAmount.Controls.Add(Me.Label4)
        Me.pnlAmount.Controls.Add(Me.txtAmount)
        Me.pnlAmount.Location = New System.Drawing.Point(3, 3)
        Me.pnlAmount.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlAmount.Name = "pnlAmount"
        Me.pnlAmount.Size = New System.Drawing.Size(603, 274)
        Me.pnlAmount.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(229, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(102, 48)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Cash Tendered"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAmount
        '
        Me.txtAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 25.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount.ForeColor = System.Drawing.Color.Green
        Me.txtAmount.Location = New System.Drawing.Point(336, 9)
        Me.txtAmount.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(255, 45)
        Me.txtAmount.TabIndex = 4
        Me.txtAmount.Text = "100.00"
        Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(2, 2)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(123, 16)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Payment (CASH)"
        '
        'frmPay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.CausesValidation = False
        Me.ClientSize = New System.Drawing.Size(621, 638)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.pnlMain)
        Me.Controls.Add(Me.pnlBill)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPay"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TransparencyKey = System.Drawing.Color.Transparent
        Me.pnlBill.ResumeLayout(False)
        Me.pnlBill.PerformLayout()
        Me.pnlMain.ResumeLayout(False)
        Me.pnlButtons.ResumeLayout(False)
        Me.pnlDetail.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAmount.ResumeLayout(False)
        Me.pnlAmount.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblBill As System.Windows.Forms.Label
    Friend WithEvents pnlBill As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pnlMain As System.Windows.Forms.Panel
    Friend WithEvents pnlAmount As System.Windows.Forms.Panel
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents cmdButton04 As System.Windows.Forms.Button
    Friend WithEvents cmdButton03 As System.Windows.Forms.Button
    Friend WithEvents cmdButton02 As System.Windows.Forms.Button
    Friend WithEvents cmdButton01 As System.Windows.Forms.Button
    Friend WithEvents cmdButton00 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlDetail As System.Windows.Forms.Panel
    Friend WithEvents pnlButtons As System.Windows.Forms.Panel
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblChange As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
