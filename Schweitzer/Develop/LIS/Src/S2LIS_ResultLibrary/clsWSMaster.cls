VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWSMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LAB008"
Attribute VB_Ext_KEY = "Member1" ,"LAB031"
Option Explicit

'/* Data Fields Property */
Private mvarWorkCd As String
Private mvarTestCd As String
Private mvarSpcCd As String
Private mvarWorkInfo As String
Private mvarEmpId As String
'
'/* Class Default Property */
Private mvarRecordCount As Long
Private mvarLAB008 As LAB008
Private mvarEOF As Boolean
Private mvarBOF As Boolean
Private mvarKey As String
Private mvarErrNo As Long
Private mvarErrText As String
Private m_INDEX As Long

'/* Class Default Property */
Public Property Get ErrNo() As Long
    mvarErrNo = OraErr.Err
    ErrNo = mvarErrNo
End Property

Public Property Get ErrText() As String
    mvarErrText = OraErr.ErrText
    ErrText = mvarErrText
End Property

Public Property Let Key(ByVal vData As String)
   'Collection Class Key Setting
    vData = mvarWorkCd & mvarTestCd & mvarSpcCd
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Get BOF() As Boolean
    BOF = mvarBOF
End Property

Public Property Get EOF() As Boolean
    EOF = mvarEOF
End Property

Public Sub MoveFirst()
'Copy First Data Collection to Class Property
   CopyDataToProperty 1
   '
End Sub

Public Sub MoveLast()
'Copy Last Data Collection to Class Propertys
   CopyDataToProperty Me.LAB008.Count
   '
End Sub

Public Sub MoveNext()
'Copy Next Data Collection to Class Propertys
   With Me
      CopyDataToProperty (m_INDEX + 1)
   End With
   '
End Sub

Public Sub MovePrevious()
'Copy Previous Data Collection to Class Propertys
   With Me
      CopyDataToProperty (m_INDEX - 1)
   End With
   '
End Sub

Public Property Get LAB008() As LAB008
    If mvarLAB008 Is Nothing Then
        Set mvarLAB008 = New LAB008
    End If
    Set LAB008 = mvarLAB008
End Property

Public Property Set LAB008(vData As LAB008)
    Set mvarLAB008 = vData
End Property

Private Sub Class_Initialize()
   '
   m_INDEX = 0
End Sub

Private Sub Class_Terminate()
   Set mvarLAB008 = Nothing
End Sub

Public Property Let RecordCount(ByVal vData As Long)
   If Me.LAB008.Count > 0 Then
      mvarRecordCount = Me.LAB008.Count
   Else
      mvarRecordCount = vData
   End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = mvarRecordCount
End Property

'/* Data Fields Property */
Public Property Let WorkInfo(ByVal vData As String)
    mvarWorkInfo = vData
End Property

Public Property Get WorkInfo() As String
    WorkInfo = mvarWorkInfo
End Property

Public Property Let EmpId(ByVal vData As String)
    mvarEmpId = vData
End Property

Public Property Get EmpId() As String
    EmpId = mvarEmpId
End Property

Public Property Let WorkCd(ByVal vData As String)
    mvarWorkCd = vData
End Property

Public Property Get WorkCd() As String
    WorkCd = mvarWorkCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mvarSpcCd = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mvarSpcCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mvarTestCd = vData
End Property

Public Property Get TestCd() As String
    TestCd = mvarTestCd
End Property

'/* Data Method */
Public Function LoadTable(Optional pOraSession As Object, _
   Optional pDatabase As Variant, _
   Optional pWorkCd As String = "Default", _
   Optional pTestCd As String = "Default", _
   Optional pSpcCd As String = "Default")
'Session open & Data Load
Dim DS As Recordset
   
'   MsgBox "LoadTable clsWSMaster ", vbCritical
   
'   If pDatabase Is Nothing Then
'      OraDS.InitConnect
'   Else
'      OraDS.InitConnect pOraSession, pDatabase
'   End If
'   'Oracle_Connection Fail Check
'   If gblnDBConnection = False Then Exit Function
   '
   If pWorkCd = "Default" Then
      pWorkCd = vbNullString
   End If
   If pTestCd = "Default" Then
      pTestCd = vbNullString
   End If
   If pSpcCd = "Default" Then
      pSpcCd = vbNullString
   End If
   On Error GoTo ErrorLoadTable
'   OraDS.GetLAB008 DS, pWorkCd, pTestCd, pSpcCd
   Set mvarLAB008 = New LAB008
   '
   If DS.RecordCount = 0 Then
      mvarBOF = False
      mvarEOF = False
      Exit Function
   End If
   '
   DS.MoveFirst
   Do Until DS.EOF
      mvarWorkCd = "" & DS.Fields("WorkCd").Value
      mvarTestCd = "" & DS.Fields("TestCd").Value
      mvarSpcCd = "" & DS.Fields("SpcCd").Value
      mvarWorkInfo = "" & DS.Fields("WorkInfo").Value
      mvarEmpId = "" & DS.Fields("EmpId").Value
      '
      'Collection Class Key Setting
      mvarKey = mvarWorkCd & mvarTestCd & mvarSpcCd
      '
      mvarRecordCount = DS.RecordCount
      mvarLAB008.Add mvarWorkCd, mvarTestCd, mvarSpcCd, _
         mvarWorkInfo, mvarEmpId, mvarRecordCount, mvarLAB008, mvarKey
      DS.MoveNext
   Loop
   '
   MoveFirst
   OraErr.Reset
   Exit Function
   '
ErrorLoadTable:
   OraErr.Reset
   If OraDS.LastServerErr <> 0 Then
      glngErrorNo = OraDS.LastServerErr
      gstrErrorMsg = OraDS.LastServerErrText
   Else
      If IsObject(Err) Then
         glngErrorNo = Err.Number
         gstrErrorMsg = Err.Description
      End If
   End If
   '
End Function

Private Sub CopyDataToProperty(ByVal index As Long)
   'With LAB008.Item(index)
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_INDEX = 0
         Else
            m_INDEX = Me.RecordCount + 1
         End If
         mvarWorkCd = 0
         mvarTestCd = vbNullString
         mvarSpcCd = vbNullString
         mvarWorkInfo = vbNullString
         mvarEmpId = vbNullString
         mvarKey = vbNullString                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB008.Count
         m_INDEX = index
   Else
      With Me.LAB008.Item(index)
         mvarWorkCd = .WorkCd
         mvarTestCd = .TestCd
         mvarSpcCd = .SpcCd
         mvarWorkCd = .WorkCd
         mvarWorkInfo = .WorkInfo
         mvarEmpId = .EmpId
         mvarKey = .Key                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB008.Count
         m_INDEX = index
      End With
   End If
   '
   With Me.LAB008
      If .Count > 0 Then
         Select Case m_INDEX
            Case Is < 1
               mvarBOF = True
               mvarEOF = False
            Case Is > .Count
               mvarEOF = True
               mvarBOF = False
            Case Else
               mvarBOF = False
               mvarEOF = False
         End Select
      Else
         mvarEOF = True
         mvarBOF = True
      End If
   End With
   '
End Sub

Public Function GetString(ByVal FieldName As String) As String
Dim ii As Long
Dim aryTmp() As String
Dim objTmp As clsWSMaster
   '
   If Me.LAB008.Count = 0 Then
      GetString = ""
      Exit Function
   End If
   ii = 0
   ReDim aryTmp(Me.LAB008.Count)
   For Each objTmp In Me.LAB008
      With objTmp
         Select Case UCase(FieldName)
            Case "TestCd"
               aryTmp(ii) = .TestCd
            Case "SpcCd"
               aryTmp(ii) = .SpcCd
            Case "WorkInfo"
               aryTmp(ii) = .WorkInfo
            Case "EmpId"
               aryTmp(ii) = .EmpId
         End Select
      End With
      ii = ii + 1
   Next
   GetString = Join(aryTmp, "$")
   ReDim aryTmp(0)
End Function
