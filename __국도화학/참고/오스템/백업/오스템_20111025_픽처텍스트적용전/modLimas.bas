Attribute VB_Name = "modLimas"
Option Explicit

Public Function DBConnect_MDS() As Boolean ' MS Acess2000 ������ ���̽� ���϶�
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean

On Error GoTo ConnectError

    UserName = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_USER_ID)
    Password = GetString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD)

    If (UserName = "admin") And (Password = "20990101") Then
        DBConnect_MDS = True
    Else
        DBConnect_MDS = False
        Exit Function
    End If
    Screen.MousePointer = vbDefault
    DBConnect_MDS = True
 
 Exit Function

ConnectError:

    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"


End Function

Sub Main()
    Dim strMsg      As String
    Dim rv          As Long
    Dim LocalPath   As String
    Dim strLicense As String
    Dim strKey  As String
    
    '�ι� ���� ���� ����
'    If App.PrevInstance Then
'       MsgBox "�󺧷� ���α׷��� �̹� �������Դϴ�.", vbExclamation
'       End
'    End If
        
    If Len(GetString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD)) = 0 Then
        frmUserSet.Show vbModal
    End If

    '���� Form ��Ÿ��
    frmLabelDesign.Show
    
End Sub

'������ Ʈ���� ���ڿ��� ����
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, LenB(StrConv(strdata, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

