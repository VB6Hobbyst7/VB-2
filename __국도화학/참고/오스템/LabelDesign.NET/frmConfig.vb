Option Strict Off
Option Explicit On
Friend Class frmConfig
	Inherits System.Windows.Forms.Form
	'===============================================================================
	'  ���α׷� : ������ ���ö�Ʈ ���� ��
	'  �� �� �� : frmConfig.frm
	'  �� �� �� : 2011.09.21
	'  �� �� �� : ������
	'  Ȩ������ : http://www.didiminfoinfo.co.kr
	'  ��    �� :
	'  �����̷� :
	'===============================================================================
	
	
	'UPGRADE_WARNING: ���� �ʱ�ȭ�� �� cboLayout.SelectedIndexChanged �̺�Ʈ�� �߻��մϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLayout_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLayout.SelectedIndexChanged
		'FIXIT: 'strTmp'��(��) �ʱ⿡ ���ε��Ǵ� ������ �������� �����Ͻʽÿ�.                                            FixIT90210ae-R1672-R1B8ZE
		Dim strTmp As Object
		
		'UPGRADE_WARNING: strTmp ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strTmp = cboLayout.Text
		
		'UPGRADE_WARNING: strTmp ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If strTmp = "�߰�" Then
			cmdAdd.Enabled = True
			cmdSet.Enabled = False
			cmdEdit.Enabled = False
			cmdDel.Enabled = False
			txtWLayOut.Text = ""
			txtHLayOut.Text = ""
			txtWLayOut.Focus()
		Else
			cmdAdd.Enabled = False
			cmdSet.Enabled = True
			cmdEdit.Enabled = True
			cmdDel.Enabled = True
			'FIXIT: 'Mid' �Լ��� 'Mid$' �Լ��� �ٲٽʽÿ�.                                                        FixIT90210ae-R9757-R1B8ZE
			'UPGRADE_WARNING: strTmp ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			txtWLayOut.Text = Mid(strTmp, 1, InStr(strTmp, ":") - 1)
			'FIXIT: 'Mid' �Լ��� 'Mid$' �Լ��� �ٲٽʽÿ�.                                                        FixIT90210ae-R9757-R1B8ZE
			'UPGRADE_WARNING: strTmp ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			txtHLayOut.Text = Mid(strTmp, InStr(strTmp, ":") + 1)
		End If
		
		
	End Sub
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		
		If txtWLayOut.Text = "" Then
			MsgBox("���̸� �Է��ϼ���.", MsgBoxStyle.Information, Me.Text)
			txtWLayOut.Focus()
			Exit Sub
		End If
		
		If txtHLayOut.Text = "" Then
			MsgBox("���̸� �Է��ϼ���.", MsgBoxStyle.Information, Me.Text)
			txtHLayOut.Focus()
			Exit Sub
		End If
		
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		If Not IsNumeric(Trim(txtWLayOut.Text)) Then
			MsgBox("���̴� ���ڸ� �Է��� �����մϴ�.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
			txtWLayOut.Focus()
			Exit Sub
		End If
		
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		If Not IsNumeric(Trim(txtHLayOut.Text)) Then
			MsgBox("���̴� ���ڸ� �Է��� �����մϴ�.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
			txtHLayOut.Focus()
			Exit Sub
		End If
		
		gLayOutUse = CStr(cboLayout.SelectedIndex)
		
		Call PutSetup("LAYOUT", "Cnt", CStr(UBound(gLayOutValue) + 1))
		
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		Call PutSetup("LAYOUT", CStr(UBound(gLayOutValue) + CDbl("1")), Trim(txtWLayOut.Text) & ":" & Trim(txtHLayOut.Text))
		
		Call GetSetup()
		
		Call LoadConfig()
		
	End Sub
	
	'-- ���� ����
	Private Sub cmdConfirm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConfirm.Click
		Dim Parity As String
		Dim sEquipNo As String
		
		On Error GoTo ErrorHandler
		
		If MsgBox("������ �����Ͻðڽ��ϱ�?", MsgBoxStyle.Critical + MsgBoxStyle.OKCancel + MsgBoxStyle.DefaultButton2, "Ȯ��!") = MsgBoxResult.Cancel Then
			Me.Close()
			Exit Sub
		Else
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gSetup.Image = Trim(txtConfig(0).Text) : Call PutSetup("CONFIG", "ImagePath", gSetup.Image) : gImage = gSetup.Image
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gSetup.Layout = Trim(txtConfig(1).Text) : Call PutSetup("CONFIG", "LayoutPath", gSetup.Layout) : gLayOut = gSetup.Layout
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gSetup.Logo = Trim(txtConfig(2).Text) : Call PutSetup("CONFIG", "LogoPath", gSetup.Logo) : gLogo = gSetup.Logo
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gSetup.Scan = Trim(txtConfig(3).Text) : Call PutSetup("CONFIG", "ScanPath", gSetup.Scan) : gScan = gSetup.Scan
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gSetup.Work = Trim(txtConfig(4).Text) : Call PutSetup("CONFIG", "WorkPath", gSetup.Work) : gWork = gSetup.Work
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gSetup.Log = Trim(txtConfig(5).Text) : Call PutSetup("CONFIG", "LogPath", gSetup.Log) : gLog = gSetup.Log
			
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gScaleMode = Trim(txtConfig(6).Text) : Call PutSetup("MODE", "ScaleMode", gScaleMode)
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gScaleCal = Trim(txtConfig(7).Text) : Call PutSetup("MODE", "ScaleCal", gScaleCal)
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			gDevide = Trim(txtConfig(8).Text) : Call PutSetup("MODE", "Devide", gDevide)
			
			Me.Close()
		End If
		
		Exit Sub
		
ErrorHandler: 
		Resume Next
		
	End Sub
	
	'-- ����
	Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click
		
		gLayOutUse = CStr(cboLayout.SelectedIndex)
		
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		Call PutSetup("LAYOUT", gLayOutUse, Trim(txtWLayOut.Text) & ":" & Trim(txtHLayOut.Text))
		
		Call GetSetup()
		
		Call LoadConfig()
		
	End Sub
	
	'-- �ݱ�
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		Me.Close()
	End Sub
	
	'-- ����
	Private Sub cmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSet.Click
		
		gLayOutUse = CStr(cboLayout.SelectedIndex)
		
		Call PutSetup("LAYOUT", "Use", gLayOutUse)
		
		Call GetSetup()
		
		Call LoadConfig()
		
	End Sub
	
	'-- �ҷ�����
	Private Sub frmConfig_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Short
		
		Me.Width = VB6.TwipsToPixelsX(4995)
		Me.Height = VB6.TwipsToPixelsY(7035) '6555 '6510
		
		Call LoadConfig()
		
	End Sub
	
	Private Sub LoadConfig()
		Dim i As Short
		
		txtConfig(0).Text = gImage
		txtConfig(1).Text = gLayOut
		txtConfig(2).Text = gLogo
		txtConfig(3).Text = gScan
		txtConfig(4).Text = gWork
		txtConfig(5).Text = gLog
		
		txtConfig(6).Text = gScaleMode
		txtConfig(7).Text = gScaleCal
		txtConfig(8).Text = gDevide
		
		cboLayout.Items.Clear()
		cboLayout.Items.Add("�߰�")
		For i = 1 To UBound(gLayOutValue)
			cboLayout.Items.Add(gLayOutValue(i))
		Next 
		
		cboLayout.SelectedIndex = gLayOutUse '- 1
		
		'FIXIT: 'Mid' �Լ��� 'Mid$' �Լ��� �ٲٽʽÿ�.                                                        FixIT90210ae-R9757-R1B8ZE
		txtWLayOut.Text = Mid(gLayOutValue(CInt(gLayOutUse)), 1, InStr(gLayOutValue(CInt(gLayOutUse)), ":") - 1)
		'FIXIT: 'Mid' �Լ��� 'Mid$' �Լ��� �ٲٽʽÿ�.                                                        FixIT90210ae-R9757-R1B8ZE
		txtHLayOut.Text = Mid(gLayOutValue(CInt(gLayOutUse)), InStr(gLayOutValue(CInt(gLayOutUse)), ":") + 1)
		
	End Sub
	
	'-- Hidden ���� ���̱�/�Ⱥ��̱�
	Private Sub lblHiddenView_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblHiddenView.DoubleClick
		If Label1(6).Visible = True Then
			Label1(6).Visible = False
			Label1(7).Visible = False
			txtConfig(6).Visible = False
			txtConfig(7).Visible = False
			Label3.Visible = False
			Label4.Visible = False
		Else
			Label1(6).Visible = True
			Label1(7).Visible = True
			txtConfig(6).Visible = True
			txtConfig(7).Visible = True
			Label3.Visible = True
			Label4.Visible = True
		End If
	End Sub
End Class