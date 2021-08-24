VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSMAuthority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'//COM009
Private mvarDeptFg As String
Private mvarFormID As String
Private mvarReadFg As String        '// '0','1',
Private mvarWriteFg As String
Private mvarPrintFg  As String

Private mvarCanRead As Boolean
Private mvarCanWrite As Boolean
Private mvarCanPrint As Boolean


'[�Ӽ�] - �μ�(������Ʈ) ����
Public Property Let DeptFg(ByVal vData As String)
    mvarDeptFg = vData
End Property
Public Property Get DeptFg() As String
    DeptFg = mvarDeptFg
End Property

'[�Ӽ�] - �� �Ƶ�
Public Property Let FormId(ByVal vData As String)
    mvarFormID = vData
End Property
Public Property Get FormId() As String
    FormId = mvarFormID
End Property

'[�Ӽ�] - �б���� ����
Public Property Let ReadFg(ByVal vData As String)
    mvarReadFg = vData
    mvarCanRead = IIf(mvarReadFg = "1", True, False)
End Property
Public Property Get ReadFg() As String
    ReadFg = mvarReadFg
End Property

'[�Ӽ�] - ������� ����
Public Property Let WriteFg(ByVal vData As String)
    mvarWriteFg = vData
    mvarCanWrite = IIf(mvarWriteFg = "1", True, False)
End Property
Public Property Get WriteFg() As String
    WriteFg = mvarWriteFg
End Property

'[�Ӽ�] - ��±��� ����
Public Property Let PrintFg(ByVal vData As String)
    mvarPrintFg = vData
    mvarCanPrint = IIf(mvarPrintFg = "1", True, False)
End Property
Public Property Get PrintFg() As String
    PrintFg = mvarPrintFg
End Property

'[�Ӽ�] - �б����
Public Property Get CanRead() As Boolean
    CanRead = mvarCanRead
End Property

'[�Ӽ�] - �������
Public Property Get CanWrite() As String
    CanWrite = mvarCanWrite
End Property

'[�Ӽ�] - ��±���
Public Property Get CanPrint() As String
    CanPrint = mvarCanPrint
End Property

