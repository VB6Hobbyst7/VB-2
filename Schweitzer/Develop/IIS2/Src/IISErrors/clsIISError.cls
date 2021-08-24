VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISError.cls
'   �ۼ���  : �̻��
'   ��  ��  : Error Ŭ����
'   �ۼ���  : 2003-12-03
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mFileNm         As String       '������ �߻��� ���ϸ�
Private mClassNm        As String       '������ �߻��� Ŭ������
Private mProcedureNm    As String       '������ �߻��� ���ν�����
Private mDescription    As String       '������ ���� ����
Private mErrorTime      As Date         '������ �߻��� �ð�

Public Property Get FileNm() As String
    FileNm = mFileNm
End Property

Public Property Let FileNm(ByVal vData As String)
    mFileNm = vData
End Property

Public Property Get ClassNm() As String
    ClassNm = mClassNm
End Property

Public Property Let ClassNm(ByVal vData As String)
    mClassNm = vData
End Property

Public Property Get ProcedureNm() As String
    ProcedureNm = mProcedureNm
End Property

Public Property Let ProcedureNm(ByVal vData As String)
    mProcedureNm = vData
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

Public Property Get ErrorTime() As Date
    ErrorTime = mErrorTime
End Property

Public Property Let ErrorTime(ByVal vData As Date)
    mErrorTime = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���������� �α����Ͽ� ���
'   �μ� :
'       - pFileNm       : ������ �߻��� ���ϸ�
'       - pClassNm      : ������ �߻��� Ŭ������
'       - pProcedureNm  : ������ �߻��� ���ν�����
'       - pDescription  : ������ ���� ����
'       - pErrorTime    : ������ �߻��� �ð�
'-----------------------------------------------------------------------------'
Public Sub SetLog(ByVal pFileNm As String, ByVal pClassNm As String, ByVal pProcedureNm As String, _
                    ByVal pDescription As String, ByVal pErrorTime As Date)
    Dim strFileNm       As String   '���ϸ�
    Dim intFileNo       As Integer  '���Ϲ�ȣ
    
    mFileNm = pFileNm:              mClassNm = pClassNm
    mProcedureNm = pProcedureNm:    mDescription = pDescription
    mErrorTime = pErrorTime
    
On Error Resume Next
    '## Log ���丮 ������ ����
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    '## ������ �����ϸ� Append, ������ Create
    intFileNo = FreeFile
    strFileNm = App.Path & "\Log\Error(" & Format(mErrorTime, "YYYY-MM-DD") & ").log"
    If Dir(strFileNm) = "" Then
        Open strFileNm For Output As #intFileNo
    Else
        Open strFileNm For Append As #intFileNo
    End If
    
    '## �������� ���
    Print #intFileNo, "[" & CStr(mErrorTime) & "]"
    Print #intFileNo, "Project Name   : " & mFileNm
    Print #intFileNo, "Class Name     : " & mClassNm
    Print #intFileNo, "Procedure Name : " & mProcedureNm
    Print #intFileNo, "Description    : " & mDescription & vbCrLf
    
    Close intFileNo
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub Clear()
    mFileNm = "":       mClassNm = ""
    mProcedureNm = "":  mDescription = ""
End Sub