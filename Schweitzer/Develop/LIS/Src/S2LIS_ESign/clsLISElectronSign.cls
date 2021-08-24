VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISElectronSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSignId                  As String
Private mvarSignNm                  As String
Private mvarErrMsg                  As String
Private mvarElectronSignPath        As String
Private mvarElectronSignFileName    As String
Private mvarElectronSingPassword    As String
Private mvarFormVisible             As Boolean
Private mvarElectronSingOk          As Boolean
Private mvarIsElectronSignature     As Boolean
Private mvarImageTrue               As Boolean

Private mvarForm                    As frmESign
Private objMe                       As clsLISElectronSign
Private objFolder                   As Scripting.FileSystemObject

Friend Property Let ElectronSingOk(ByVal vData As Boolean)
    '���ڼ����� Ȯ���Ͽ��°� �ϴ� �Ǵ�(�н����� Ȯ�ο���)
    '
    mvarElectronSingOk = vData
    '
End Property

Public Property Get ElectronSignPrintOk() As Boolean
    '���ڼ����̹����� ����� �� �ִ����� �Ǵ�.
    Dim objFolder As New Scripting.FileSystemObject
    '
    If objFolder.FileExists(mvarElectronSignPath & "\" & mvarElectronSignFileName) = True Then
        ElectronSignPrintOk = True
    Else
        ElectronSignPrintOk = False
    End If
    '
End Property

Public Property Get ElectronSingOk() As Boolean
    ElectronSingOk = mvarElectronSingOk
End Property

Public Property Get ErrMsg() As String
    ErrMsg = mvarErrMsg
End Property

Friend Property Let ImageTrue(ByVal vData As Boolean)
    mvarImageTrue = vData
End Property

Public Property Get ImageTrue() As Boolean
    ImageTrue = mvarImageTrue
End Property

Public Property Get SignId() As String
    SignId = mvarSignId
End Property

Public Property Get SignNm() As String
    SignNm = mvarSignNm
End Property

Public Property Get ElectronSingPassword() As String
    ElectronSingPassword = mvarElectronSingPassword
End Property

Public Property Get ElectronSignPath() As String
    ElectronSignPath = mvarElectronSignPath
End Property

Public Property Get ElectronSignFileName() As String
    ElectronSignFileName = mvarElectronSignFileName
End Property

Public Sub ShowESignForm()
Dim objMe As New clsLISElectronSign
Dim objFolder As New Scripting.FileSystemObject
    '
    Set mvarForm = New frmESign
    With mvarForm
        .Tag = ""
        .lblEsinNm = mvarSignNm
        .lblSignId = mvarSignId
        .lblPass = mvarElectronSingPassword
        Set objMe = Me
        .SetESign objMe
        
        If mvarImageTrue = True Then
            .lblAuthorization.Visible = True
            .lblWarnig.Visible = False
            .lblNonVerify.Visible = False
            .cmdOk.Enabled = True
            .imgSign.Tag = mvarElectronSignPath & "\" & mvarElectronSignFileName
            .imgSign.Picture = LoadPicture(mvarElectronSignPath & "\" & mvarElectronSignFileName)
            .cmdAuthoCancel.Enabled = True
            .lblPassNm.Enabled = True
            .txtPass.Enabled = True
            .txtPass.BackColor = vbWhite

        Else
            .cmdOk.Enabled = False
            .lblAuthorization.Visible = False
            .lblWarnig.Visible = True
            .lblNonVerify.Visible = True
            .imgSign.Tag = ""
            .imgSign.Picture = LoadPicture()
            .cmdAuthoCancel.Enabled = False
            .lblPassNm.Enabled = False
            .txtPass.BackColor = &HDBE6E6
            .txtPass.Enabled = False
            
        End If
    End With
    '
    mvarForm.Show vbModal
    mvarFormVisible = True
    DoEvents
    '
End Sub

Public Function LoadElectronSign(ByVal pEmpId As String, ByVal pClientPath As String) As Boolean
    
    Dim strSql As String
    Dim strPath As String
    Dim RS As Recordset
    Dim ii As Long
    
    '
    mvarImageTrue = False
    Set objFolder = New Scripting.FileSystemObject
    '
    strSql = "select * from " & T_LAB032 & _
             " where " & DBW("cdindex", LC3_ElectronicSign, 2) & _
             "   and " & DBW("cdval1", pEmpId, 2)
    
    Set RS = New Recordset
    RS.Open strSql, DBConn
    
    If RS.RecordCount > 0 Then
        
        mvarSignId = pEmpId
        mvarSignNm = RS.Fields("field2").Value & ""
        mvarElectronSignPath = ""
        mvarElectronSignFileName = mvarSignNm & ".jpg"
        mvarElectronSingPassword = RS.Fields("field1").Value & ""
        
        Set RS = Nothing
        '
        '7.1 Home\Schweitzer\Common\Help : ���� Help
        '7.2 Home\Schweitzer\Common\DLL : ���� DLL
        '7.3 Home\Schweitzer\(Laboratory,Anatomic Pth,BloodBank)\Bin : ������, DLL
        '7.4 Home\Schweitzer\(Laboratory,Anatomic Pth,BloodBank)\Report : Crystal Report
        '7.5 Home\Schweitzer\(Laboratory,Anatomic Pth,BloodBank)\Etc : ini����, image�� -> \���\���ڼ����̹���
        
        '���ڼ��� �̹��������� validation check
        If objFolder.FolderExists(pClientPath) = True Then
            If objFolder.FolderExists(pClientPath & "\" & "Etc") = False Then
                strPath = objFolder.CreateFolder(pClientPath & "\" & "Etc")
            Else
                strPath = pClientPath & "\" & "Etc"
            End If
            '
            If objFolder.FolderExists(strPath & "\" & mvarSignId) = False Then
                strPath = objFolder.CreateFolder(strPath & "\" & mvarSignId)
            Else
                strPath = strPath & "\" & mvarSignId
            End If
            If objFolder.FileExists(strPath & "\" & mvarElectronSignFileName) = False Then
                mvarErrMsg = "���ڼ��� �̹����� ��ϵǾ� ���� �ʽ��ϴ�." & _
                    vbNewLine & "���ڼ��� �̹����� ������ ������ ����Ͻʽÿ�."
                mvarImageTrue = False
            Else
                mvarImageTrue = True
            End If
            mvarElectronSignPath = strPath
        Else
            mvarErrMsg = "���ڼ��� �̹����� ������ �߸��Ǿ� �ֽ��ϴ�." & _
                vbNewLine & "���α׷� �¾����¸� Ȯ���Ͽ� �ֽʽÿ�."
            Exit Function
        End If
    Else
        mvarErrMsg = "���ڼ��� ������ �����ϴ�." & vbNewLine & _
            "���ڼ��� [�����������]�� �Ͻʽÿ�."
        Exit Function
    End If
    '
    LoadElectronSign = True
    '
    Set objFolder = Nothing
    '
End Function

Private Sub Class_Initialize()
    '
    Clear
    '
End Sub


Private Sub Clear()
    '
    mvarErrMsg = ""
    mvarSignId = ""
    mvarSignNm = ""
    mvarElectronSignPath = ""
    mvarElectronSignFileName = ""
    mvarElectronSingPassword = ""
    mvarImageTrue = False
    '
End Sub

Private Sub Class_Terminate()
    '
    Clear
    If mvarFormVisible = True Then
        Set mvarForm = Nothing
    End If
    mvarFormVisible = False
    '
    Set objMe = Nothing
    '
End Sub