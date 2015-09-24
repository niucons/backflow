<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CompanyInfoForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CompanyInfoForm))
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtCCCDI = New System.Windows.Forms.TextBox
        Me.txtPhone = New System.Windows.Forms.TextBox
        Me.dtpCalibration = New System.Windows.Forms.DateTimePicker
        Me.lblCompanyName = New System.Windows.Forms.Label
        Me.lblCalibrationDate = New System.Windows.Forms.Label
        Me.lblCCCDIXC = New System.Windows.Forms.Label
        Me.lblPhone = New System.Windows.Forms.Label
        Me.btnCompanyInfo = New System.Windows.Forms.Button
        Me.txtTestKit = New System.Windows.Forms.TextBox
        Me.lblTestKit = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(134, 21)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(199, 20)
        Me.txtName.TabIndex = 0
        '
        'txtCCCDI
        '
        Me.txtCCCDI.Location = New System.Drawing.Point(134, 127)
        Me.txtCCCDI.Name = "txtCCCDI"
        Me.txtCCCDI.Size = New System.Drawing.Size(199, 20)
        Me.txtCCCDI.TabIndex = 4
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(134, 47)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(199, 20)
        Me.txtPhone.TabIndex = 1
        '
        'dtpCalibration
        '
        Me.dtpCalibration.CalendarForeColor = System.Drawing.SystemColors.GrayText
        Me.dtpCalibration.CalendarTitleBackColor = System.Drawing.Color.LightSlateGray
        Me.dtpCalibration.Cursor = System.Windows.Forms.Cursors.Default
        Me.dtpCalibration.Location = New System.Drawing.Point(134, 100)
        Me.dtpCalibration.Name = "dtpCalibration"
        Me.dtpCalibration.Size = New System.Drawing.Size(200, 20)
        Me.dtpCalibration.TabIndex = 3
        '
        'lblCompanyName
        '
        Me.lblCompanyName.AutoSize = True
        Me.lblCompanyName.Location = New System.Drawing.Point(46, 24)
        Me.lblCompanyName.Name = "lblCompanyName"
        Me.lblCompanyName.Size = New System.Drawing.Size(82, 13)
        Me.lblCompanyName.TabIndex = 4
        Me.lblCompanyName.Text = "Company Name"
        '
        'lblCalibrationDate
        '
        Me.lblCalibrationDate.AutoSize = True
        Me.lblCalibrationDate.Location = New System.Drawing.Point(47, 106)
        Me.lblCalibrationDate.Name = "lblCalibrationDate"
        Me.lblCalibrationDate.Size = New System.Drawing.Size(82, 13)
        Me.lblCalibrationDate.TabIndex = 5
        Me.lblCalibrationDate.Text = "Calibration Date"
        '
        'lblCCCDIXC
        '
        Me.lblCCCDIXC.AutoSize = True
        Me.lblCCCDIXC.Location = New System.Drawing.Point(47, 133)
        Me.lblCCCDIXC.Name = "lblCCCDIXC"
        Me.lblCCCDIXC.Size = New System.Drawing.Size(63, 13)
        Me.lblCCCDIXC.TabIndex = 6
        Me.lblCCCDIXC.Text = "CCCDI# XC"
        '
        'lblPhone
        '
        Me.lblPhone.AutoSize = True
        Me.lblPhone.Location = New System.Drawing.Point(47, 50)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(85, 13)
        Me.lblPhone.TabIndex = 7
        Me.lblPhone.Text = "Company Phone"
        '
        'btnCompanyInfo
        '
        Me.btnCompanyInfo.Location = New System.Drawing.Point(258, 164)
        Me.btnCompanyInfo.Name = "btnCompanyInfo"
        Me.btnCompanyInfo.Size = New System.Drawing.Size(75, 23)
        Me.btnCompanyInfo.TabIndex = 5
        Me.btnCompanyInfo.Text = "OK"
        Me.btnCompanyInfo.UseVisualStyleBackColor = True
        '
        'txtTestKit
        '
        Me.txtTestKit.Location = New System.Drawing.Point(134, 74)
        Me.txtTestKit.Name = "txtTestKit"
        Me.txtTestKit.Size = New System.Drawing.Size(199, 20)
        Me.txtTestKit.TabIndex = 2
        '
        'lblTestKit
        '
        Me.lblTestKit.AutoSize = True
        Me.lblTestKit.Location = New System.Drawing.Point(49, 80)
        Me.lblTestKit.Name = "lblTestKit"
        Me.lblTestKit.Size = New System.Drawing.Size(43, 13)
        Me.lblTestKit.TabIndex = 10
        Me.lblTestKit.Text = "Test Kit"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(172, 164)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'CompanyInfoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(389, 207)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblTestKit)
        Me.Controls.Add(Me.txtTestKit)
        Me.Controls.Add(Me.btnCompanyInfo)
        Me.Controls.Add(Me.lblPhone)
        Me.Controls.Add(Me.lblCCCDIXC)
        Me.Controls.Add(Me.lblCalibrationDate)
        Me.Controls.Add(Me.lblCompanyName)
        Me.Controls.Add(Me.dtpCalibration)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.txtCCCDI)
        Me.Controls.Add(Me.txtName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CompanyInfoForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Company Information"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtCCCDI As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents dtpCalibration As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblCompanyName As System.Windows.Forms.Label
    Friend WithEvents lblCalibrationDate As System.Windows.Forms.Label
    Friend WithEvents lblCCCDIXC As System.Windows.Forms.Label
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents btnCompanyInfo As System.Windows.Forms.Button
    Friend WithEvents txtTestKit As System.Windows.Forms.TextBox
    Friend WithEvents lblTestKit As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
