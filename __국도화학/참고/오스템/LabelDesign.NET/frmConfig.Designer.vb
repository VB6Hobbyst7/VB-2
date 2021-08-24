<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmConfig
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
	Public WithEvents lblHiddenView As System.Windows.Forms.Label
	Public WithEvents Picture1 As System.Windows.Forms.Panel
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdConfirm As System.Windows.Forms.Button
	Public WithEvents cmdDel As System.Windows.Forms.Button
	Public WithEvents cmdEdit As System.Windows.Forms.Button
	Public WithEvents cmdSet As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents txtHLayOut As System.Windows.Forms.TextBox
	Public WithEvents txtWLayOut As System.Windows.Forms.TextBox
	Public WithEvents cboLayout As System.Windows.Forms.ComboBox
	Public WithEvents _txtConfig_8 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_7 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_6 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtConfig_0 As System.Windows.Forms.TextBox
	Public WithEvents _Label1_11 As System.Windows.Forms.Label
	Public WithEvents _Label1_10 As System.Windows.Forms.Label
	Public WithEvents _Label1_9 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents _Label1_8 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_4 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtConfig As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
	'Windows Form 디자이너를 사용하여 수정할 수 있습니다.
	'코드 편집기를 사용하여 수정하지 마십시오.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmConfig))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Picture1 = New System.Windows.Forms.Panel
		Me.lblHiddenView = New System.Windows.Forms.Label
		Me.cmdExit = New System.Windows.Forms.Button
		Me.cmdConfirm = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cmdDel = New System.Windows.Forms.Button
		Me.cmdEdit = New System.Windows.Forms.Button
		Me.cmdSet = New System.Windows.Forms.Button
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.txtHLayOut = New System.Windows.Forms.TextBox
		Me.txtWLayOut = New System.Windows.Forms.TextBox
		Me.cboLayout = New System.Windows.Forms.ComboBox
		Me._txtConfig_8 = New System.Windows.Forms.TextBox
		Me._txtConfig_7 = New System.Windows.Forms.TextBox
		Me._txtConfig_6 = New System.Windows.Forms.TextBox
		Me._txtConfig_5 = New System.Windows.Forms.TextBox
		Me._txtConfig_4 = New System.Windows.Forms.TextBox
		Me._txtConfig_3 = New System.Windows.Forms.TextBox
		Me._txtConfig_2 = New System.Windows.Forms.TextBox
		Me._txtConfig_1 = New System.Windows.Forms.TextBox
		Me._txtConfig_0 = New System.Windows.Forms.TextBox
		Me._Label1_11 = New System.Windows.Forms.Label
		Me._Label1_10 = New System.Windows.Forms.Label
		Me._Label1_9 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me._Label1_8 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_4 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.txtConfig = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Picture1.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtConfig, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "라벨디자이너 설정"
		Me.ClientSize = New System.Drawing.Size(335, 518)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.Font = New System.Drawing.Font("굴림체", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.Icon = CType(resources.GetObject("frmConfig.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmConfig"
		Me.Picture1.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.Picture1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Picture1.Size = New System.Drawing.Size(327, 37)
		Me.Picture1.Location = New System.Drawing.Point(4, 0)
		Me.Picture1.TabIndex = 11
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.Visible = True
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture1.Name = "Picture1"
		Me.lblHiddenView.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblHiddenView.Text = "Environment"
		Me.lblHiddenView.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHiddenView.ForeColor = System.Drawing.SystemColors.highlightText
		Me.lblHiddenView.Size = New System.Drawing.Size(315, 29)
		Me.lblHiddenView.Location = New System.Drawing.Point(6, 4)
		Me.lblHiddenView.TabIndex = 15
		Me.lblHiddenView.BackColor = System.Drawing.Color.Transparent
		Me.lblHiddenView.Enabled = True
		Me.lblHiddenView.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblHiddenView.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblHiddenView.UseMnemonic = True
		Me.lblHiddenView.Visible = True
		Me.lblHiddenView.AutoSize = False
		Me.lblHiddenView.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblHiddenView.Name = "lblHiddenView"
		Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExit.Text = "닫기"
		Me.cmdExit.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdExit.Size = New System.Drawing.Size(61, 25)
		Me.cmdExit.Location = New System.Drawing.Point(264, 458)
		Me.cmdExit.TabIndex = 10
		Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExit.CausesValidation = True
		Me.cmdExit.Enabled = True
		Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExit.TabStop = True
		Me.cmdExit.Name = "cmdExit"
		Me.cmdConfirm.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdConfirm.Text = "확인"
		Me.cmdConfirm.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdConfirm.Size = New System.Drawing.Size(61, 25)
		Me.cmdConfirm.Location = New System.Drawing.Point(196, 458)
		Me.cmdConfirm.TabIndex = 9
		Me.cmdConfirm.BackColor = System.Drawing.SystemColors.Control
		Me.cmdConfirm.CausesValidation = True
		Me.cmdConfirm.Enabled = True
		Me.cmdConfirm.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdConfirm.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdConfirm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdConfirm.TabStop = True
		Me.cmdConfirm.Name = "cmdConfirm"
		Me.Frame1.Text = "Program Path Setting"
		Me.Frame1.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.Frame1.ForeColor = System.Drawing.SystemColors.Highlight
		Me.Frame1.Size = New System.Drawing.Size(325, 409)
		Me.Frame1.Location = New System.Drawing.Point(4, 40)
		Me.Frame1.TabIndex = 12
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cmdDel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDel.Text = "삭제"
		Me.cmdDel.Enabled = False
		Me.cmdDel.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdDel.Size = New System.Drawing.Size(37, 47)
		Me.cmdDel.Location = New System.Drawing.Point(290, 368)
		Me.cmdDel.TabIndex = 33
		Me.cmdDel.Visible = False
		Me.cmdDel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDel.CausesValidation = True
		Me.cmdDel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDel.TabStop = True
		Me.cmdDel.Name = "cmdDel"
		Me.cmdEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEdit.Text = "수정"
		Me.cmdEdit.Enabled = False
		Me.cmdEdit.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdEdit.Size = New System.Drawing.Size(45, 47)
		Me.cmdEdit.Location = New System.Drawing.Point(210, 258)
		Me.cmdEdit.TabIndex = 32
		Me.cmdEdit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEdit.CausesValidation = True
		Me.cmdEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEdit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEdit.TabStop = True
		Me.cmdEdit.Name = "cmdEdit"
		Me.cmdSet.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSet.Text = "적용"
		Me.cmdSet.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdSet.Size = New System.Drawing.Size(47, 47)
		Me.cmdSet.Location = New System.Drawing.Point(154, 258)
		Me.cmdSet.TabIndex = 31
		Me.cmdSet.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSet.CausesValidation = True
		Me.cmdSet.Enabled = True
		Me.cmdSet.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSet.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSet.TabStop = True
		Me.cmdSet.Name = "cmdSet"
		Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdd.Text = "추가"
		Me.cmdAdd.Enabled = False
		Me.cmdAdd.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.cmdAdd.Size = New System.Drawing.Size(45, 47)
		Me.cmdAdd.Location = New System.Drawing.Point(264, 258)
		Me.cmdAdd.TabIndex = 34
		Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdd.CausesValidation = True
		Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdd.TabStop = True
		Me.cmdAdd.Name = "cmdAdd"
		Me.txtHLayOut.AutoSize = False
		Me.txtHLayOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtHLayOut.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.txtHLayOut.Size = New System.Drawing.Size(53, 23)
		Me.txtHLayOut.Location = New System.Drawing.Point(96, 282)
		Me.txtHLayOut.TabIndex = 30
		Me.txtHLayOut.AcceptsReturn = True
		Me.txtHLayOut.BackColor = System.Drawing.SystemColors.Window
		Me.txtHLayOut.CausesValidation = True
		Me.txtHLayOut.Enabled = True
		Me.txtHLayOut.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHLayOut.HideSelection = True
		Me.txtHLayOut.ReadOnly = False
		Me.txtHLayOut.Maxlength = 0
		Me.txtHLayOut.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHLayOut.MultiLine = False
		Me.txtHLayOut.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHLayOut.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHLayOut.TabStop = True
		Me.txtHLayOut.Visible = True
		Me.txtHLayOut.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtHLayOut.Name = "txtHLayOut"
		Me.txtWLayOut.AutoSize = False
		Me.txtWLayOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtWLayOut.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me.txtWLayOut.Size = New System.Drawing.Size(53, 23)
		Me.txtWLayOut.Location = New System.Drawing.Point(96, 256)
		Me.txtWLayOut.TabIndex = 28
		Me.txtWLayOut.AcceptsReturn = True
		Me.txtWLayOut.BackColor = System.Drawing.SystemColors.Window
		Me.txtWLayOut.CausesValidation = True
		Me.txtWLayOut.Enabled = True
		Me.txtWLayOut.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWLayOut.HideSelection = True
		Me.txtWLayOut.ReadOnly = False
		Me.txtWLayOut.Maxlength = 0
		Me.txtWLayOut.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWLayOut.MultiLine = False
		Me.txtWLayOut.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWLayOut.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWLayOut.TabStop = True
		Me.txtWLayOut.Visible = True
		Me.txtWLayOut.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtWLayOut.Name = "txtWLayOut"
		Me.cboLayout.Size = New System.Drawing.Size(215, 21)
		Me.cboLayout.Location = New System.Drawing.Point(96, 226)
		Me.cboLayout.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLayout.TabIndex = 26
		Me.cboLayout.BackColor = System.Drawing.SystemColors.Window
		Me.cboLayout.CausesValidation = True
		Me.cboLayout.Enabled = True
		Me.cboLayout.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboLayout.IntegralHeight = True
		Me.cboLayout.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLayout.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLayout.Sorted = False
		Me.cboLayout.TabStop = True
		Me.cboLayout.Visible = True
		Me.cboLayout.Name = "cboLayout"
		Me._txtConfig_8.AutoSize = False
		Me._txtConfig_8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._txtConfig_8.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_8.Size = New System.Drawing.Size(35, 23)
		Me._txtConfig_8.Location = New System.Drawing.Point(96, 196)
		Me._txtConfig_8.TabIndex = 6
		Me._txtConfig_8.AcceptsReturn = True
		Me._txtConfig_8.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_8.CausesValidation = True
		Me._txtConfig_8.Enabled = True
		Me._txtConfig_8.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_8.HideSelection = True
		Me._txtConfig_8.ReadOnly = False
		Me._txtConfig_8.Maxlength = 0
		Me._txtConfig_8.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_8.MultiLine = False
		Me._txtConfig_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_8.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_8.TabStop = True
		Me._txtConfig_8.Visible = True
		Me._txtConfig_8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_8.Name = "_txtConfig_8"
		Me._txtConfig_7.AutoSize = False
		Me._txtConfig_7.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_7.Size = New System.Drawing.Size(69, 23)
		Me._txtConfig_7.Location = New System.Drawing.Point(96, 346)
		Me._txtConfig_7.TabIndex = 8
		Me._txtConfig_7.Visible = False
		Me._txtConfig_7.AcceptsReturn = True
		Me._txtConfig_7.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_7.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_7.CausesValidation = True
		Me._txtConfig_7.Enabled = True
		Me._txtConfig_7.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_7.HideSelection = True
		Me._txtConfig_7.ReadOnly = False
		Me._txtConfig_7.Maxlength = 0
		Me._txtConfig_7.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_7.MultiLine = False
		Me._txtConfig_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_7.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_7.TabStop = True
		Me._txtConfig_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_7.Name = "_txtConfig_7"
		Me._txtConfig_6.AutoSize = False
		Me._txtConfig_6.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_6.Size = New System.Drawing.Size(35, 23)
		Me._txtConfig_6.Location = New System.Drawing.Point(96, 318)
		Me._txtConfig_6.TabIndex = 7
		Me._txtConfig_6.Visible = False
		Me._txtConfig_6.AcceptsReturn = True
		Me._txtConfig_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_6.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_6.CausesValidation = True
		Me._txtConfig_6.Enabled = True
		Me._txtConfig_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_6.HideSelection = True
		Me._txtConfig_6.ReadOnly = False
		Me._txtConfig_6.Maxlength = 0
		Me._txtConfig_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_6.MultiLine = False
		Me._txtConfig_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_6.TabStop = True
		Me._txtConfig_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_6.Name = "_txtConfig_6"
		Me._txtConfig_5.AutoSize = False
		Me._txtConfig_5.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_5.Size = New System.Drawing.Size(215, 23)
		Me._txtConfig_5.Location = New System.Drawing.Point(96, 168)
		Me._txtConfig_5.TabIndex = 5
		Me._txtConfig_5.AcceptsReturn = True
		Me._txtConfig_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_5.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_5.CausesValidation = True
		Me._txtConfig_5.Enabled = True
		Me._txtConfig_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_5.HideSelection = True
		Me._txtConfig_5.ReadOnly = False
		Me._txtConfig_5.Maxlength = 0
		Me._txtConfig_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_5.MultiLine = False
		Me._txtConfig_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_5.TabStop = True
		Me._txtConfig_5.Visible = True
		Me._txtConfig_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_5.Name = "_txtConfig_5"
		Me._txtConfig_4.AutoSize = False
		Me._txtConfig_4.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_4.Size = New System.Drawing.Size(215, 23)
		Me._txtConfig_4.Location = New System.Drawing.Point(96, 140)
		Me._txtConfig_4.TabIndex = 4
		Me._txtConfig_4.AcceptsReturn = True
		Me._txtConfig_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_4.CausesValidation = True
		Me._txtConfig_4.Enabled = True
		Me._txtConfig_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_4.HideSelection = True
		Me._txtConfig_4.ReadOnly = False
		Me._txtConfig_4.Maxlength = 0
		Me._txtConfig_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_4.MultiLine = False
		Me._txtConfig_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_4.TabStop = True
		Me._txtConfig_4.Visible = True
		Me._txtConfig_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_4.Name = "_txtConfig_4"
		Me._txtConfig_3.AutoSize = False
		Me._txtConfig_3.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_3.Size = New System.Drawing.Size(215, 23)
		Me._txtConfig_3.Location = New System.Drawing.Point(96, 112)
		Me._txtConfig_3.TabIndex = 3
		Me._txtConfig_3.AcceptsReturn = True
		Me._txtConfig_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_3.CausesValidation = True
		Me._txtConfig_3.Enabled = True
		Me._txtConfig_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_3.HideSelection = True
		Me._txtConfig_3.ReadOnly = False
		Me._txtConfig_3.Maxlength = 0
		Me._txtConfig_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_3.MultiLine = False
		Me._txtConfig_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_3.TabStop = True
		Me._txtConfig_3.Visible = True
		Me._txtConfig_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_3.Name = "_txtConfig_3"
		Me._txtConfig_2.AutoSize = False
		Me._txtConfig_2.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_2.Size = New System.Drawing.Size(215, 23)
		Me._txtConfig_2.Location = New System.Drawing.Point(96, 83)
		Me._txtConfig_2.TabIndex = 2
		Me._txtConfig_2.AcceptsReturn = True
		Me._txtConfig_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_2.CausesValidation = True
		Me._txtConfig_2.Enabled = True
		Me._txtConfig_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_2.HideSelection = True
		Me._txtConfig_2.ReadOnly = False
		Me._txtConfig_2.Maxlength = 0
		Me._txtConfig_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_2.MultiLine = False
		Me._txtConfig_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_2.TabStop = True
		Me._txtConfig_2.Visible = True
		Me._txtConfig_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_2.Name = "_txtConfig_2"
		Me._txtConfig_1.AutoSize = False
		Me._txtConfig_1.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_1.Size = New System.Drawing.Size(215, 23)
		Me._txtConfig_1.Location = New System.Drawing.Point(96, 55)
		Me._txtConfig_1.TabIndex = 1
		Me._txtConfig_1.AcceptsReturn = True
		Me._txtConfig_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_1.CausesValidation = True
		Me._txtConfig_1.Enabled = True
		Me._txtConfig_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_1.HideSelection = True
		Me._txtConfig_1.ReadOnly = False
		Me._txtConfig_1.Maxlength = 0
		Me._txtConfig_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_1.MultiLine = False
		Me._txtConfig_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_1.TabStop = True
		Me._txtConfig_1.Visible = True
		Me._txtConfig_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_1.Name = "_txtConfig_1"
		Me._txtConfig_0.AutoSize = False
		Me._txtConfig_0.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._txtConfig_0.Size = New System.Drawing.Size(215, 23)
		Me._txtConfig_0.Location = New System.Drawing.Point(96, 26)
		Me._txtConfig_0.TabIndex = 0
		Me._txtConfig_0.AcceptsReturn = True
		Me._txtConfig_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtConfig_0.BackColor = System.Drawing.SystemColors.Window
		Me._txtConfig_0.CausesValidation = True
		Me._txtConfig_0.Enabled = True
		Me._txtConfig_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtConfig_0.HideSelection = True
		Me._txtConfig_0.ReadOnly = False
		Me._txtConfig_0.Maxlength = 0
		Me._txtConfig_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtConfig_0.MultiLine = False
		Me._txtConfig_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtConfig_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtConfig_0.TabStop = True
		Me._txtConfig_0.Visible = True
		Me._txtConfig_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._txtConfig_0.Name = "_txtConfig_0"
		Me._Label1_11.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_11.Text = "넓이"
		Me._Label1_11.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_11.Size = New System.Drawing.Size(26, 13)
		Me._Label1_11.Location = New System.Drawing.Point(64, 284)
		Me._Label1_11.TabIndex = 35
		Me._Label1_11.BackColor = System.Drawing.Color.Transparent
		Me._Label1_11.Enabled = True
		Me._Label1_11.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_11.UseMnemonic = True
		Me._Label1_11.Visible = True
		Me._Label1_11.AutoSize = False
		Me._Label1_11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_11.Name = "_Label1_11"
		Me._Label1_10.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_10.Text = "높이"
		Me._Label1_10.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_10.Size = New System.Drawing.Size(30, 13)
		Me._Label1_10.Location = New System.Drawing.Point(60, 258)
		Me._Label1_10.TabIndex = 29
		Me._Label1_10.BackColor = System.Drawing.Color.Transparent
		Me._Label1_10.Enabled = True
		Me._Label1_10.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_10.UseMnemonic = True
		Me._Label1_10.Visible = True
		Me._Label1_10.AutoSize = False
		Me._Label1_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_10.Name = "_Label1_10"
		Me._Label1_9.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_9.Text = "라벨용지 :"
		Me._Label1_9.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_9.Size = New System.Drawing.Size(80, 13)
		Me._Label1_9.Location = New System.Drawing.Point(8, 230)
		Me._Label1_9.TabIndex = 27
		Me._Label1_9.BackColor = System.Drawing.Color.Transparent
		Me._Label1_9.Enabled = True
		Me._Label1_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_9.UseMnemonic = True
		Me._Label1_9.Visible = True
		Me._Label1_9.AutoSize = False
		Me._Label1_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_9.Name = "_Label1_9"
		Me.Label5.Text = "배"
		Me.Label5.Size = New System.Drawing.Size(39, 13)
		Me.Label5.Location = New System.Drawing.Point(138, 202)
		Me.Label5.TabIndex = 25
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me._Label1_8.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_8.Text = "기본배율 :"
		Me._Label1_8.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_8.Size = New System.Drawing.Size(80, 13)
		Me._Label1_8.Location = New System.Drawing.Point(9, 202)
		Me._Label1_8.TabIndex = 24
		Me._Label1_8.BackColor = System.Drawing.Color.Transparent
		Me._Label1_8.Enabled = True
		Me._Label1_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_8.UseMnemonic = True
		Me._Label1_8.Visible = True
		Me._Label1_8.AutoSize = False
		Me._Label1_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_8.Name = "_Label1_8"
		Me.Label4.Text = "픽셀대비 트윕값 1:15"
		Me.Label4.Size = New System.Drawing.Size(141, 13)
		Me.Label4.Location = New System.Drawing.Point(170, 352)
		Me.Label4.TabIndex = 23
		Me.Label4.Visible = False
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "1=트윕, 3:픽셀"
		Me.Label3.Size = New System.Drawing.Size(163, 13)
		Me.Label3.Location = New System.Drawing.Point(142, 324)
		Me.Label3.TabIndex = 22
		Me.Label3.Visible = False
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_7.Text = "Scale Cal :"
		Me._Label1_7.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_7.Size = New System.Drawing.Size(80, 13)
		Me._Label1_7.Location = New System.Drawing.Point(8, 352)
		Me._Label1_7.TabIndex = 21
		Me._Label1_7.Visible = False
		Me._Label1_7.BackColor = System.Drawing.Color.Transparent
		Me._Label1_7.Enabled = True
		Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_7.UseMnemonic = True
		Me._Label1_7.AutoSize = False
		Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_7.Name = "_Label1_7"
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_6.Text = "Scale Mode :"
		Me._Label1_6.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_6.Size = New System.Drawing.Size(80, 13)
		Me._Label1_6.Location = New System.Drawing.Point(8, 324)
		Me._Label1_6.TabIndex = 20
		Me._Label1_6.Visible = False
		Me._Label1_6.BackColor = System.Drawing.Color.Transparent
		Me._Label1_6.Enabled = True
		Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.AutoSize = False
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_5.Text = "Log Path :"
		Me._Label1_5.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_5.Size = New System.Drawing.Size(80, 13)
		Me._Label1_5.Location = New System.Drawing.Point(9, 174)
		Me._Label1_5.TabIndex = 19
		Me._Label1_5.BackColor = System.Drawing.Color.Transparent
		Me._Label1_5.Enabled = True
		Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_5.UseMnemonic = True
		Me._Label1_5.Visible = True
		Me._Label1_5.AutoSize = False
		Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_5.Name = "_Label1_5"
		Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_4.Text = "Work Path :"
		Me._Label1_4.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_4.Size = New System.Drawing.Size(80, 13)
		Me._Label1_4.Location = New System.Drawing.Point(9, 146)
		Me._Label1_4.TabIndex = 18
		Me._Label1_4.BackColor = System.Drawing.Color.Transparent
		Me._Label1_4.Enabled = True
		Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_4.UseMnemonic = True
		Me._Label1_4.Visible = True
		Me._Label1_4.AutoSize = False
		Me._Label1_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_4.Name = "_Label1_4"
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_3.Text = "Scan Path :"
		Me._Label1_3.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_3.Size = New System.Drawing.Size(80, 13)
		Me._Label1_3.Location = New System.Drawing.Point(9, 118)
		Me._Label1_3.TabIndex = 17
		Me._Label1_3.BackColor = System.Drawing.Color.Transparent
		Me._Label1_3.Enabled = True
		Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_3.UseMnemonic = True
		Me._Label1_3.Visible = True
		Me._Label1_3.AutoSize = False
		Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_3.Name = "_Label1_3"
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_2.Text = "Logo Path :"
		Me._Label1_2.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_2.Size = New System.Drawing.Size(80, 13)
		Me._Label1_2.Location = New System.Drawing.Point(9, 89)
		Me._Label1_2.TabIndex = 16
		Me._Label1_2.BackColor = System.Drawing.Color.Transparent
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.Text = "Image Path :"
		Me._Label1_0.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_0.Size = New System.Drawing.Size(80, 13)
		Me._Label1_0.Location = New System.Drawing.Point(9, 32)
		Me._Label1_0.TabIndex = 14
		Me._Label1_0.BackColor = System.Drawing.Color.Transparent
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_1.Text = "Layout Path :"
		Me._Label1_1.Font = New System.Drawing.Font("굴림체", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
		Me._Label1_1.Size = New System.Drawing.Size(80, 13)
		Me._Label1_1.Location = New System.Drawing.Point(9, 61)
		Me._Label1_1.TabIndex = 13
		Me._Label1_1.BackColor = System.Drawing.Color.Transparent
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me.Controls.Add(Picture1)
		Me.Controls.Add(cmdExit)
		Me.Controls.Add(cmdConfirm)
		Me.Controls.Add(Frame1)
		Me.Picture1.Controls.Add(lblHiddenView)
		Me.Frame1.Controls.Add(cmdDel)
		Me.Frame1.Controls.Add(cmdEdit)
		Me.Frame1.Controls.Add(cmdSet)
		Me.Frame1.Controls.Add(cmdAdd)
		Me.Frame1.Controls.Add(txtHLayOut)
		Me.Frame1.Controls.Add(txtWLayOut)
		Me.Frame1.Controls.Add(cboLayout)
		Me.Frame1.Controls.Add(_txtConfig_8)
		Me.Frame1.Controls.Add(_txtConfig_7)
		Me.Frame1.Controls.Add(_txtConfig_6)
		Me.Frame1.Controls.Add(_txtConfig_5)
		Me.Frame1.Controls.Add(_txtConfig_4)
		Me.Frame1.Controls.Add(_txtConfig_3)
		Me.Frame1.Controls.Add(_txtConfig_2)
		Me.Frame1.Controls.Add(_txtConfig_1)
		Me.Frame1.Controls.Add(_txtConfig_0)
		Me.Frame1.Controls.Add(_Label1_11)
		Me.Frame1.Controls.Add(_Label1_10)
		Me.Frame1.Controls.Add(_Label1_9)
		Me.Frame1.Controls.Add(Label5)
		Me.Frame1.Controls.Add(_Label1_8)
		Me.Frame1.Controls.Add(Label4)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame1.Controls.Add(_Label1_7)
		Me.Frame1.Controls.Add(_Label1_6)
		Me.Frame1.Controls.Add(_Label1_5)
		Me.Frame1.Controls.Add(_Label1_4)
		Me.Frame1.Controls.Add(_Label1_3)
		Me.Frame1.Controls.Add(_Label1_2)
		Me.Frame1.Controls.Add(_Label1_0)
		Me.Frame1.Controls.Add(_Label1_1)
		Me.Label1.SetIndex(_Label1_11, CType(11, Short))
		Me.Label1.SetIndex(_Label1_10, CType(10, Short))
		Me.Label1.SetIndex(_Label1_9, CType(9, Short))
		Me.Label1.SetIndex(_Label1_8, CType(8, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_4, CType(4, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.txtConfig.SetIndex(_txtConfig_8, CType(8, Short))
		Me.txtConfig.SetIndex(_txtConfig_7, CType(7, Short))
		Me.txtConfig.SetIndex(_txtConfig_6, CType(6, Short))
		Me.txtConfig.SetIndex(_txtConfig_5, CType(5, Short))
		Me.txtConfig.SetIndex(_txtConfig_4, CType(4, Short))
		Me.txtConfig.SetIndex(_txtConfig_3, CType(3, Short))
		Me.txtConfig.SetIndex(_txtConfig_2, CType(2, Short))
		Me.txtConfig.SetIndex(_txtConfig_1, CType(1, Short))
		Me.txtConfig.SetIndex(_txtConfig_0, CType(0, Short))
		CType(Me.txtConfig, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Picture1.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class