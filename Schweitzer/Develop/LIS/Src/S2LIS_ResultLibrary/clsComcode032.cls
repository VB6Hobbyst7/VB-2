VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComcode032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LAB031"
Attribute VB_Ext_KEY = "Member1" ,"LAB032"
Option Explicit

'/* Data Fields Property */
Private mvarCdIndex As String
Private mvarCdVal1 As String
Private mvarField1 As String
Private mvarField2 As String
Private mvarField3 As String
Private mvarField4 As String
Private mvarField5 As String
Private mvarText1 As String
Private mvarText2 As String
'
'/* Class Default Property */
Private mvarRecordCount As Long
Private mvarLAB032 As LAB032
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
    vData = mvarCdIndex & mvarCdVal1
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
   CopyDataToProperty Me.LAB032.Count
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

Public Property Get LAB032() As LAB032
    If mvarLAB032 Is Nothing Then
        Set mvarLAB032 = New LAB032
    End If
    Set LAB032 = mvarLAB032
End Property

Public Property Set LAB032(vData As LAB032)
    Set mvarLAB032 = vData
End Property

Private Sub Class_Initialize()
   '
   m_INDEX = 0
End Sub

Private Sub Class_Terminate()
   Set mvarLAB032 = Nothing
End Sub

Public Property Let RecordCount(ByVal vData As Long)
   If Me.LAB032.Count > 0 Then
      mvarRecordCount = Me.LAB032.Count
   Else
      mvarRecordCount = vData
   End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = mvarRecordCount
End Property

'/* Data Fields Property */
Public Property Let Field1(ByVal vData As String)
    mvarField1 = vData
End Property

Public Property Get Field1() As String
    Field1 = mvarField1
End Property

Public Property Let Field2(ByVal vData As String)
    mvarField2 = vData
End Property

Public Property Get Field2() As String
    Field2 = mvarField2
End Property

Public Property Let Field3(ByVal vData As String)
    mvarField3 = vData
End Property

Public Property Get Field3() As String
    Field3 = mvarField3
End Property

Public Property Let Field4(ByVal vData As String)
    mvarField4 = vData
End Property

Public Property Get Field4() As String
    Field4 = mvarField4
End Property

Public Property Let Field5(ByVal vData As String)
    mvarField5 = vData
End Property

Public Property Get Field5() As String
    Field5 = mvarField5
End Property

Public Property Let CdIndex(ByVal vData As String)
    mvarCdIndex = vData
End Property

Public Property Get CdIndex() As String
    CdIndex = mvarCdIndex
End Property

Public Property Let CdVal1(ByVal vData As String)
    mvarCdVal1 = vData
End Property

Public Property Get CdVal1() As String
    CdVal1 = mvarCdVal1
End Property

Public Property Let Text2(ByVal vData As String)
    mvarText2 = vData
End Property

Public Property Get Text2() As String
    Text2 = mvarText2
End Property

Public Property Let Text1(ByVal vData As String)
    mvarText1 = vData
End Property

Public Property Get Text1() As String
    Text1 = mvarText1
End Property

'/* Data Method */
Public Function LoadTable(Optional pCdIndex_A As String = "Default", _
                        Optional pCdIndex_B As String = "Default", _
                        Optional pCdVal1 As String = "Default")
'Public Function LoadTable(Optional pOraSession As Object, _
'   Optional pDatabase As Variant, _
'   Optional pCdIndex_A As String = "Default", _
'   Optional pCdIndex_B As String = "Default", _
'   Optional pCdVal1 As String = "Default")
'Session open & Data Load
Dim DS As Recordset

'   MsgBox "LoadTable clsComcode032 ", vbCritical
'   If pDatabase Is Nothing Then
'      OraDS.InitConnect
'   Else
'      OraDS.InitConnect pOraSession, pDatabase
'   End If
'   'Oracle_Connection Fail Check
'   If gblnDBConnection = False Then Exit Function
   '
   If pCdIndex_A = "Default" Then
      pCdIndex_A = vbNullString
   End If
   If pCdIndex_B = "Default" Then
      pCdIndex_B = vbNullString
   End If
   If pCdVal1 = "Default" Then
      pCdVal1 = vbNullString
   End If
   On Error GoTo ErrorLoadTable
   OraDS.GetLAB032 DS, pCdIndex_A, pCdIndex_B, pCdVal1
   Set mvarLAB032 = New LAB032
   '
   If DS.RecordCount = 0 Then
      mvarBOF = False
      mvarEOF = False
      Exit Function
   End If
   '
   DS.MoveFirst
   Do Until DS.EOF
      mvarCdIndex = "" & DS.Fields("CDINDEX").Value
      mvarCdVal1 = "" & DS.Fields("CDVAL1").Value
      mvarField1 = "" & DS.Fields("Field1").Value
      mvarField2 = "" & DS.Fields("Field2").Value
      mvarField3 = "" & DS.Fields("Field3").Value
      mvarField4 = "" & DS.Fields("Field4").Value
      mvarField5 = "" & DS.Fields("Field5").Value
      mvarText1 = "" & DS.Fields("TEXT1").Value
      mvarText2 = "" & DS.Fields("TEXT2").Value
      '
      'Collection Class Key Setting
      mvarKey = mvarCdIndex & mvarCdVal1
      '
      mvarRecordCount = DS.RecordCount
      mvarLAB032.Add mvarCdIndex, mvarCdVal1, _
         mvarField1, mvarField2, mvarField3, mvarField4, mvarField5, mvarText1, mvarText2, _
         mvarRecordCount, mvarLAB032, mvarKey
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
   'With LAB032.Item(index)
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_INDEX = 0
         Else
            m_INDEX = Me.RecordCount + 1
         End If
         mvarCdIndex = 0
         mvarCdVal1 = vbNullString
         mvarField1 = vbNullString
         mvarField2 = vbNullString
         mvarField3 = vbNullString
         mvarField4 = vbNullString
         mvarField5 = vbNullString
         mvarText1 = vbNullString
         mvarText2 = vbNullString
         mvarKey = vbNullString                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB032.Count
         m_INDEX = index
   Else
      With Me.LAB032.Item(index)
         mvarCdIndex = .CdIndex
         mvarCdVal1 = .CdVal1
         mvarCdIndex = .CdIndex
         mvarField1 = .Field1
         mvarField2 = .Field2
         mvarField3 = .Field3
         mvarField4 = .Field4
         mvarField5 = .Field5
         mvarText1 = .Text1
         mvarText2 = .Text2
         mvarKey = .Key                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB032.Count
         m_INDEX = index
      End With
   End If
   '
   With Me.LAB032
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
Dim objTmp As clsComcode032
   '
   If Me.LAB032.Count = 0 Then
      GetString = ""
      Exit Function
   End If
   ii = 0
   ReDim aryTmp(Me.LAB032.Count)
   For Each objTmp In Me.LAB032
      With objTmp
         Select Case UCase(FieldName)
            Case "CDVAL1"
               aryTmp(ii) = .CdVal1
            Case "FIELD1"
               aryTmp(ii) = .Field1
            Case "FIELD2"
               aryTmp(ii) = .Field2
            Case "FIELD3"
               aryTmp(ii) = .Field3
            Case "FIELD4"
               aryTmp(ii) = .Field4
            Case "FIELD5"
               aryTmp(ii) = .Field5
            Case "TEXT1"
               aryTmp(ii) = .Text1
            Case "TEXT2"
               aryTmp(ii) = .Text2
         End Select
      End With
      ii = ii + 1
   Next
   GetString = Join(aryTmp, "$")
   ReDim aryTmp(0)
End Function
