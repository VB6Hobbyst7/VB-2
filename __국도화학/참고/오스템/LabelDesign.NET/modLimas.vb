Option Strict Off
Option Explicit On
Module modLimas
	
	Public Function DBConnect_MDS() As Boolean ' MS Acess2000 데이터 베이스 붙일때
		Dim DB_Name As String
		Dim UserName As String
		Dim Password As String
		Dim blnWinNTAuth As Boolean
		
		On Error GoTo ConnectError
		
		'UPGRADE_WARNING: GetString() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		UserName = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_USER_ID)
		'UPGRADE_WARNING: GetString() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Password = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD)
		
		If (UserName = "admin") And (Password = "20990101") Then
			DBConnect_MDS = True
		Else
			DBConnect_MDS = False
			Exit Function
		End If
		'UPGRADE_WARNING: Screen 속성 Screen.MousePointer에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		DBConnect_MDS = True
		
		Exit Function
		
ConnectError: 
		
		MsgBox("   Error No. : " & Err.Number & vbCrLf & " Description : " & Err.Description & vbCrLf & "      Source : " & Err.Source & vbCrLf & vbCrLf, MsgBoxStyle.Critical, " DB Open Error")
		
		
	End Function
	
	'UPGRADE_WARNING: Sub Main()이 끝나면 응용 프로그램이 종료됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"'
	Public Sub Main()
		Dim strMsg As String
		Dim rv As Integer
		Dim LocalPath As String
		Dim strLicense As String
		Dim strKey As String
		
		'두번 실행 하지 않음
		'    If App.PrevInstance Then
		'       MsgBox "라벨러 프로그램이 이미 실행중입니다.", vbExclamation
		'       End
		'    End If
		
		If Len(GetString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD)) = 0 Then
			frmUserSet.ShowDialog()
		End If
		
		'메인 Form 나타남
		frmLabelDesign.Show()
		
	End Sub
	
	'레지스 트리에 문자열값 저장
	Public Sub SaveString(ByRef hKey As Integer, ByRef strPath As String, ByRef strValue As String, ByRef strdata As String)
		Dim keyhand As Integer
		Dim r As Integer
		r = RegCreateKey(hKey, strPath, keyhand)
		'UPGRADE_ISSUE: vbFromUnicode 상수가 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: LenB 함수는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, strdata, LenB(StrConv(strdata, vbFromUnicode)))
		r = RegCloseKey(keyhand)
	End Sub
End Module