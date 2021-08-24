VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTemplate.cls
'   �ۼ���  : �̻��
'   ��  ��  : ���ø� ������ Ŭ����
'   �ۼ���  : 2004-03-10
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - ���ø� �ε���
'   - ���ø� �ڵ�
'   - ���ø� ����

'## �޼ҵ�
'   - ���ø� �Է�
'   - ���ø� ����
'   - ���ø� ����
'   - ���ø� ��ȸ

Option Explicit

Private mSQL        As clsIISMasterSql

Private mCdIndex    As String   '���ø� �ε���
Private mCode       As String   '���ø� �ڵ�
Private mContent    As String   '���ø� ����

Public Property Get CdIndex() As String
    CdIndex = mCdIndex
End Property

Public Property Let CdIndex(ByVal vData As String)
    mCdIndex = vData
End Property

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal vData As String)
    mCode = vData
End Property

Public Property Get Content() As String
    Content = mContent
End Property

Public Property Let Content(ByVal vData As String)
    mContent = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���ø� �Է�
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddTemplate() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertTemplate(mCdIndex, mCode, mContent), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTemplate = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTemlate", "AddTemplate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���ø� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyTemplate() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.UpdateTemplate(mCdIndex, mCode, mContent), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyTemplate = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTemlate", "ModifyTemplate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���ø� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelTemplate() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteTemplate(mCdIndex, mCode), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTemplate = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTemlate", "DelTemplate", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���ø� ��ȸ
'   �μ� :
'       1.pCdIndex : ���ø� �ε���
'   ��ȯ : ���ø� Recordset
'-----------------------------------------------------------------------------'
Public Function GetTemplates(ByVal pCdIndex As String) As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectTemplate(pCdIndex), , adCmdText)
    Set GetTemplates = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISTemlate", "GetTemplates", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub