Option Strict Off
Option Explicit On
Friend Class frmUserSet
	Inherits System.Windows.Forms.Form
	
	Dim bConnected As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
		End
	End Sub
	
	Private Sub cmdOk_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
		
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		If Trim(txtUser.Text) = "" Then
			MsgBox(" ����ڸ��� �Է� �Ͻÿ�.")
			Exit Sub
		Else
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			Call SaveString(HKEY_CURRENT_USER, REG_POSITION, REG_USER_ID, Trim(txtUser.Text))
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			Call SaveString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD, Trim(txtPasswd.Text))
			
			If DBConnect_MDS Then
				Me.Close()
			Else
				MsgBox("�Է������� Ʋ���ϴ�. �ٽ� �õ� �Ͻʽÿ�.")
			End If
		End If
	End Sub
	
	Private Sub frmUserSet_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Escape
				Call cmdCancel_Click(cmdCancel, New System.EventArgs())
			Case System.Windows.Forms.Keys.Return
				Call cmdOk_Click(cmdOk, New System.EventArgs())
			Case Else
				
		End Select
	End Sub
	
	Private Sub frmUserSet_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_WARNING: Screen �Ӽ� Screen.MousePointer�� �� ������ �ֽ��ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
		txtUser.Text = ""
		txtPasswd.Text = ""
		
		cmdOk.Enabled = True
		
	End Sub
	
	
	Private Sub txtPasswd_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPasswd.Enter
		txtPasswd.SelectionStart = 0
		txtPasswd.SelectionLength = Len(txtPasswd.Text)
	End Sub
	
	Private Sub txtPasswd_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPasswd.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then
			Call cmdOk_Click(cmdOk, New System.EventArgs())
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub txtUser_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUser.Enter
		txtUser.SelectionStart = 0
		txtUser.SelectionLength = Len(txtUser.Text)
	End Sub
	
	Private Sub txtUser_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUser.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then
			System.Windows.Forms.SendKeys.Send("{TAB}")
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class