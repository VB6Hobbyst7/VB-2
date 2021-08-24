<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmUserSet
#Region "Windows Form 디자이너에서 생성한 코드 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'이 호출은 Windows Form 디자이너에 필요합니다.
		InitializeComponent()
	End Sub
	'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows Form 디자이너에 필요합니다.
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOk As System.Windows.Forms.Button
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents txtPasswd As System.Windows.Forms.TextBox
	Public WithEvents txtUser As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents _lblStep3_0 As System.Windows.Forms.Label
	Public WithEvents labMsg As System.Windows.Forms.Label
	Public WithEvents _lblStep3_4 As System.Windows.Forms.Label
	Public WithEvents lblStep3 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
	'Windows Form 디자이너를 사용하여 수정할 수 있습니다.
	'코드 편집기를 사용하여 수정하지 마십시오.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmUserSet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOk = New System.Windows.Forms.Button
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
		Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog
		Me.CommonDialog1Font = New System.Windows.Forms.FontDialog
		Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog
		Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog
		Me.txtPasswd = New System.Windows.Forms.TextBox
		Me.txtUser = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me._lblStep3_0 = New System.Windows.Forms.Label
		Me.labMsg = New System.Windows.Forms.Label
		Me._lblStep3_4 = New System.Windows.Forms.Label
		Me.lblStep3 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblStep3, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "User Setting"
		Me.ClientSize = New System.Drawing.Size(374, 218)
		Me.Location = New System.Drawing.Point(190, 117)
		Me.ControlBox = False
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmUserSet"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Font = New System.Drawing.Font("굴림", 9!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdCancel.Size = New System.Drawing.Size(81, 27)
		Me.cmdCancel.Location = New System.Drawing.Point(238, 162)
		Me.cmdCancel.TabIndex = 8
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOk.Text = "OK"
		Me.cmdOk.Font = New System.Drawing.Font("굴림", 9!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdOk.Size = New System.Drawing.Size(81, 27)
		Me.cmdOk.Location = New System.Drawing.Point(134, 162)
		Me.cmdOk.TabIndex = 7
		Me.cmdOk.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOk.CausesValidation = True
		Me.cmdOk.Enabled = True
		Me.cmdOk.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOk.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOk.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOk.TabStop = True
		Me.cmdOk.Name = "cmdOk"
		Me.Frame2.Size = New System.Drawing.Size(372, 8)
		Me.Frame2.Location = New System.Drawing.Point(2, 136)
		Me.Frame2.TabIndex = 5
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtPasswd.AutoSize = False
		Me.txtPasswd.Font = New System.Drawing.Font("굴림", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.txtPasswd.Size = New System.Drawing.Size(191, 24)
		Me.txtPasswd.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPasswd.Location = New System.Drawing.Point(130, 90)
		Me.txtPasswd.PasswordChar = ChrW(42)
		Me.txtPasswd.TabIndex = 1
		Me.txtPasswd.Text = "Passwd"
		Me.txtPasswd.AcceptsReturn = True
		Me.txtPasswd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPasswd.BackColor = System.Drawing.SystemColors.Window
		Me.txtPasswd.CausesValidation = True
		Me.txtPasswd.Enabled = True
		Me.txtPasswd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPasswd.HideSelection = True
		Me.txtPasswd.ReadOnly = False
		Me.txtPasswd.Maxlength = 0
		Me.txtPasswd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPasswd.MultiLine = False
		Me.txtPasswd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPasswd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPasswd.TabStop = True
		Me.txtPasswd.Visible = True
		Me.txtPasswd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtPasswd.Name = "txtPasswd"
		Me.txtUser.AutoSize = False
		Me.txtUser.Font = New System.Drawing.Font("굴림", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.txtUser.Size = New System.Drawing.Size(191, 24)
		Me.txtUser.Location = New System.Drawing.Point(129, 60)
		Me.txtUser.TabIndex = 0
		Me.txtUser.Text = "User"
		Me.txtUser.AcceptsReturn = True
		Me.txtUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUser.BackColor = System.Drawing.SystemColors.Window
		Me.txtUser.CausesValidation = True
		Me.txtUser.Enabled = True
		Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUser.HideSelection = True
		Me.txtUser.ReadOnly = False
		Me.txtUser.Maxlength = 0
		Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUser.MultiLine = False
		Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUser.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUser.TabStop = True
		Me.txtUser.Visible = True
		Me.txtUser.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtUser.Name = "txtUser"
		Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.Label1.Text = "  사용자 등록"
		Me.Label1.Font = New System.Drawing.Font("굴림", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.Label1.Size = New System.Drawing.Size(371, 39)
		Me.Label1.Location = New System.Drawing.Point(2, 2)
		Me.Label1.TabIndex = 6
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me._lblStep3_0.Text = "암호(&B):"
		Me._lblStep3_0.Font = New System.Drawing.Font("굴림", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._lblStep3_0.Size = New System.Drawing.Size(57, 15)
		Me._lblStep3_0.Location = New System.Drawing.Point(67, 92)
		Me._lblStep3_0.TabIndex = 4
		Me._lblStep3_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStep3_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblStep3_0.Enabled = True
		Me._lblStep3_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStep3_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStep3_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStep3_0.UseMnemonic = True
		Me._lblStep3_0.Visible = True
		Me._lblStep3_0.AutoSize = True
		Me._lblStep3_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStep3_0.Name = "_lblStep3_0"
		Me.labMsg.Size = New System.Drawing.Size(4, 12)
		Me.labMsg.Location = New System.Drawing.Point(6, 167)
		Me.labMsg.TabIndex = 3
		Me.labMsg.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.labMsg.BackColor = System.Drawing.SystemColors.Control
		Me.labMsg.Enabled = True
		Me.labMsg.ForeColor = System.Drawing.SystemColors.ControlText
		Me.labMsg.Cursor = System.Windows.Forms.Cursors.Default
		Me.labMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.labMsg.UseMnemonic = True
		Me.labMsg.Visible = True
		Me.labMsg.AutoSize = True
		Me.labMsg.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.labMsg.Name = "labMsg"
		Me._lblStep3_4.Text = "사용자명(&U):"
		Me._lblStep3_4.Font = New System.Drawing.Font("굴림", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._lblStep3_4.Size = New System.Drawing.Size(87, 15)
		Me._lblStep3_4.Location = New System.Drawing.Point(37, 63)
		Me._lblStep3_4.TabIndex = 2
		Me._lblStep3_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStep3_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblStep3_4.Enabled = True
		Me._lblStep3_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStep3_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStep3_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStep3_4.UseMnemonic = True
		Me._lblStep3_4.Visible = True
		Me._lblStep3_4.AutoSize = True
		Me._lblStep3_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStep3_4.Name = "_lblStep3_4"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOk)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(txtPasswd)
		Me.Controls.Add(txtUser)
		Me.Controls.Add(Label1)
		Me.Controls.Add(_lblStep3_0)
		Me.Controls.Add(labMsg)
		Me.Controls.Add(_lblStep3_4)
		Me.lblStep3.SetIndex(_lblStep3_0, CType(0, Short))
		Me.lblStep3.SetIndex(_lblStep3_4, CType(4, Short))
		CType(Me.lblStep3, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class