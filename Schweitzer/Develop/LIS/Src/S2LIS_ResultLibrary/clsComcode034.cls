VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComcode034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LAB031"
Attribute VB_Ext_KEY = "Member1" ,"LAB032"
Attribute VB_Ext_KEY = "Member2" ,"LAB034"
Option Explicit

'/* Data Fields Property */
Private mvarCdIndex As String
Private mvarCdVal1 As String
Private mvarField1 As String
Private mvarText1 As String
Private mvarText2 As String
'
'/* Class Default Property */
Private mvarRecordCount As Long
Private mvarLAB034 As LAB034
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
   CopyDataToProperty Me.LAB034.Count
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

Public Property Get LAB034() As LAB034
    If mvarLAB034 Is Nothing Then
        Set mvarLAB034 = New LAB034
    End If
    Set LAB034 = mvarLAB034
End Property

Public Property Set LAB034(vData As LAB034)
    Set mvarLAB034 = vData
End Property

Private Sub Class_Initialize()
   '
   m_INDEX = 0
End Sub

Private Sub Class_Terminate()
   Set mvarLAB034 = Nothing
End Sub

Public Property Let RecordCount(ByVal vData As Long)
   If Me.LAB034.Count > 0 Then
      mvarRecordCount = Me.LAB034.Count
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
   
'   MsgBox "LoadTable clsComcode034 ", vbCritical
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
   OraDS.GetLAB034 DS, pCdIndex_A, pCdIndex_B, pCdVal1
   Set mvarLAB034 = New LAB034
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
      mvarText1 = "" & DS.Fields("TEXT1").Value
      mvarText2 = "" & DS.Fields("TEXT2").Value
      '
      'Collection Class Key Setting
      mvarKey = mvarCdIndex & mvarCdVal1
      '
      mvarRecordCount = DS.RecordCount
      mvarLAB034.Add mvarCdIndex, mvarCdVal1, _
               mvarField1, mvarText1, mvarText2, _
               mvarRecordCount, mvarLAB034, mvarKey
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
   'With LAB034.Item(index)
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_INDEX = 0
         Else
            m_INDEX = Me.RecordCount + 1
         End If
         mvarCdIndex = 0
         mvarCdVal1 = vbNullString
         mvarField1 = vbNullString
         mvarText1 = vbNullString
         mvarText2 = vbNullString
         mvarKey = vbNullString                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB034.Count
         m_INDEX = index
   Else
      With Me.LAB034.Item(index)
         mvarCdIndex = .CdIndex
         mvarCdVal1 = .CdVal1
         mvarCdIndex = .CdIndex
         mvarField1 = .Field1
         mvarText1 = .Text1
         mvarText2 = .Text2
         mvarKey = .Key                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB034.Count
         m_INDEX = index
      End With
   End If
   '
   With Me.LAB034
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
    Dim objTmp As clsComcode034
   '
   If Me.LAB034.Count = 0 Then
      GetString = ""
      Exit Function
   End If
   ii = 0
   ReDim aryTmp(Me.LAB034.Count - 1)
   For Each objTmp In Me.LAB034
      With objTmp
         Select Case UCase(FieldName)
            Case "CDVAL1"
               aryTmp(ii) = .CdVal1
            Case "FIELD1"
               aryTmp(ii) = .Field1
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

'/* Data Method */
Public Function LoadTableFootLote(ByVal CdIndex As String, ByVal Field1 As String)
    
    Dim RS      As Recordset
    Dim sSQL    As String
    
    sSQL = " SELECT * FROM " & T_LAB034 & " WHERE " & DBW("CDINDEX=", CdIndex) & " AND " & DBW("FIELD1=", Field1) & " order by cdval1 "
    Set RS = New Recordset
    RS.Open sSQL, DBConn
    If RS.EOF Then
        sSQL = " SELECT * FROM " & T_LAB034 & " WHERE " & DBW("CDINDEX=", CdIndex) & " order by cdval1 "
    End If
    
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open sSQL, DBConn
    
    If RS.EOF Then
        mvarBOF = False
        mvarEOF = False
        Set RS = Nothing
        Exit Function
    End If
    
   On Error GoTo ErrorLoadTable
   
   Set mvarLAB034 = New LAB034
   '
   RS.MoveFirst
   Do Until RS.EOF
      mvarCdIndex = "" & RS.Fields("CDINDEX").Value
      mvarCdVal1 = "" & RS.Fields("CDVAL1").Value
      mvarField1 = "" & RS.Fields("Field1").Value
      mvarText1 = "" & RS.Fields("TEXT1").Value
      mvarText2 = "" & RS.Fields("TEXT2").Value
      '
      'Collection Class Key Setting
      mvarKey = mvarCdIndex & mvarCdVal1
      '
      mvarRecordCount = RS.RecordCount
      mvarLAB034.Add mvarCdIndex, mvarCdVal1, _
                     mvarField1, mvarText1, mvarText2, _
                     mvarRecordCount, mvarLAB034, mvarKey
      RS.MoveNext
   Loop
   '
   MoveFirst
   OraErr.Reset
   Set RS = Nothing
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

