Attribute VB_Name = "modLimas"
Option Explicit

Public Function DBConnect_MDS() As Boolean ' MS Acess2000 데이터 베이스 붙일때
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
''''
    'Create the instance of the common wrapper
    Set moCommonWrapper = New clsCommonWrapper
    'moCommonWrapper.ActivateLibrary ("DUMMY")     'DUMMY대신 key(BDL83Y75H7AU83E87Y7AC74N83E76W73RBFQ) 넣어서 테스트 해봐
    
    moCommonWrapper.ActivateLibrary ("BDL83Y75H7AU83E87Y7AC74N83E76W73RBFQo")     'DUMMY대신 key(BDL83Y75H7AU83E87Y7AC74N83E76W73RBFQ) 넣어서 테스트 해봐
                                                                                              '  BDL83Y75H7AU83E87Y7AC74N83E76W73RBFQo
    'Initialize common Dialogs
'    moCommonDialog.InitDialogs

    'Initialize language
'    moMultiLngSupport.Start App.Path & "\arabic.lng", "Arial Unicode MS"

    'Change the default button style
'    moCommonWrapper.DefaultButtonStyle = iCtlBtnStyle_Vista4
''''
''''    'Start the first form
''''    'frmSplash.Show 1

    Dim strMsg      As String
    Dim rv          As Long
    Dim LocalPath   As String
    Dim strLicense As String
    Dim strKey  As String
    
    
    
    '두번 실행 하지 않음
'    If App.PrevInstance Then
'       MsgBox "라벨러 프로그램이 이미 실행중입니다.", vbExclamation
'       End
'    End If
        
'    If Len(GetString(HKEY_CURRENT_USER, REG_POSITION, REG_PASSWD)) = 0 Then
'        frmUserSet.Show vbModal
'    End If

    '메인 Form 나타남
    frmLabelDesign.Show
    
End Sub

'레지스 트리에 문자열값 저장
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, LenB(StrConv(strdata, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

