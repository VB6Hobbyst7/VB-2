VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LAB306"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'/* Data Fields Property */

Private mvarTransDt As String
Private mvarEqpCd As String
Private mvarTransTm As String
Private mvarTransNo As String
Private mvarSpcYy As String
Private mvarSpcNo As String
Private mvarWorkArea As String
Private mvarAccDt As String
Private mvarAccNo As String
Private mvarStatFg As String
'
'/* Class Default Property */
Private mvarRecordCount As Long
Private mvarLAB306 As LAB306
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
    vData = mvarEqpCd & mvarTransDt & mvarTransTm
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
   CopyDataToProperty Me.LAB306.Count
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

Public Property Get LAB306() As LAB306
    If mvarLAB306 Is Nothing Then
        Set mvarLAB306 = New LAB306
    End If
    Set LAB306 = mvarLAB306
End Property

Public Property Set LAB306(vData As LAB306)
    Set mvarLAB306 = vData
End Property

Private Sub Class_Initialize()
   '
   m_INDEX = 0
End Sub

Private Sub Class_Terminate()
   Set mvarLAB306 = Nothing
End Sub

Public Property Let RecordCount(ByVal vData As Long)
   If Me.LAB306.Count > 0 Then
      mvarRecordCount = Me.LAB306.Count
   Else
      mvarRecordCount = vData
   End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = mvarRecordCount
End Property

'/* Data Fields Property */
Public Property Let TransDt(ByVal vData As String)
    mvarTransDt = vData
End Property

Public Property Get TransDt() As String
    TransDt = mvarTransDt
End Property

Public Property Let TransTm(ByVal vData As String)
    mvarTransTm = vData
End Property

Public Property Get TransTm() As String
    TransTm = mvarTransTm
End Property

Public Property Let TransNo(ByVal vData As String)
    mvarTransNo = vData
End Property

Public Property Get TransNo() As String
    TransNo = mvarTransNo
End Property

Public Property Let EqpCd(ByVal vData As String)
    mvarEqpCd = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mvarEqpCd
End Property

Public Property Get SpcNum() As String
'Read Only Property
    SpcNum = mvarSpcYy & "-" & Format(mvarSpcNo, "0#######")
End Property

Public Property Let SpcYy(ByVal vData As String)
    mvarSpcYy = vData
End Property

Public Property Get SpcYy() As String
    SpcYy = mvarSpcYy
End Property

Public Property Let SpcNo(ByVal vData As String)
    mvarSpcNo = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mvarSpcNo
End Property

Public Property Let WorkArea(ByVal vData As String)
    mvarWorkArea = vData
End Property

Public Property Get WorkArea() As String
    WorkArea = mvarWorkArea
End Property

Public Property Let AccDt(ByVal vData As String)
    mvarAccDt = vData
End Property

Public Property Get AccDt() As String
    AccDt = mvarAccDt
End Property

Public Property Let AccNo(ByVal vData As String)
    mvarAccNo = vData
End Property

Public Property Get AccNo() As String
    AccNo = mvarAccNo
End Property

Public Property Let StatFg(ByVal vData As String)
    mvarStatFg = vData
End Property

Public Property Get StatFg() As String
    StatFg = mvarStatFg
End Property


'/* Data Method */
Public Function LoadTable(Optional pEqpCd As String = "Default", _
                        Optional pSpcNo As String = "Default", _
                        Optional pStatFg As String = "Default")
'Public Function LoadTable(Optional pOraSession As Object, _
'   Optional pDatabase As Variant, _
'   Optional pEqpCd As String = "Default", _
'   Optional pSpcNo As String = "Default", _
'   Optional pStatFg As String = "Default")
'Session open & Data Load
Dim DS As Recordset
Dim strTmp As String
Dim aryTrans() As String
   '
'   MsgBox "LoadTable clsEquipTransfer ", vbCritical
'   If pDatabase Is Nothing Then
'      OraDS.InitConnect
'   Else
'      OraDS.InitConnect pOraSession, pDatabase
'   End If
'   'Oracle_Connection Fail Check
'   If gblnDBConnection = False Then Exit Function
   '
   On Error GoTo ErrorLoadTable
   strTmp = OraDS.GetEquipTransDtTm(DS, pEqpCd, pSpcNo)
   If strTmp = "" Then
      mvarBOF = False
      mvarEOF = False
      Exit Function
   Else
      aryTrans = Split(strTmp, vbTab)
   End If
   '
   OraDS.GetLAB306 DS, pEqpCd, aryTrans(0), aryTrans(1), pStatFg
   If DS.RecordCount = 0 Then
      mvarBOF = False
      mvarEOF = False
      Exit Function
   End If
   Set mvarLAB306 = New LAB306
   '
   DS.MoveFirst
   Do Until DS.EOF
      mvarTransDt = "" & DS.Fields("TRANSDT").Value
      mvarEqpCd = "" & DS.Fields("EQPCD").Value
      mvarTransTm = "" & DS.Fields("TRANSTM").Value
      mvarTransNo = "" & DS.Fields("TRANSNO").Value
      mvarSpcYy = "" & DS.Fields("SPCYY").Value
      mvarSpcNo = "" & DS.Fields("SPCNO").Value
      mvarWorkArea = "" & DS.Fields("WORKAREA").Value
      mvarAccDt = "" & DS.Fields("ACCDT").Value
      mvarAccNo = "" & DS.Fields("ACCSEQ").Value
      mvarStatFg = "" & DS.Fields("STATFG").Value
      '
      'Collection Class Key Setting
      mvarKey = mvarEqpCd & mvarTransDt & mvarTransTm & mvarSpcYy & mvarSpcNo
      '
      mvarRecordCount = DS.RecordCount
      mvarLAB306.Add mvarTransDt, mvarEqpCd, mvarTransTm, mvarTransNo, _
         mvarSpcYy, mvarSpcNo, mvarWorkArea, mvarAccDt, mvarAccNo, mvarStatFg, mvarRecordCount, _
         mvarLAB306, mvarKey
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
   'With LAB306.Item(index)
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_INDEX = 0
         Else
            m_INDEX = Me.RecordCount + 1
         End If
         mvarTransDt = vbNullString
         mvarEqpCd = vbNullString
         mvarTransTm = vbNullString
         mvarTransNo = vbNullString
         mvarSpcYy = vbNullString
         mvarSpcNo = vbNullString
         mvarWorkArea = vbNullString
         mvarAccDt = vbNullString
         mvarAccNo = vbNullString
         mvarKey = vbNullString                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB306.Count
         m_INDEX = index
   Else
      With Me.LAB306.Item(index)
         mvarTransDt = .TransDt
         mvarEqpCd = .EqpCd
         mvarTransTm = .TransTm
         mvarTransNo = .TransNo
         mvarTransDt = .TransDt
         mvarSpcYy = .SpcYy
         mvarSpcNo = .SpcNo
         mvarWorkArea = .WorkArea
         mvarAccDt = .AccDt
         mvarAccNo = .AccNo
         mvarKey = .Key                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB306.Count
         m_INDEX = index
      End With
   End If
   '
   With Me.LAB306
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
   '
   If Me.LAB306.Count = 0 Then
      GetString = ""
      Exit Function
   End If
   ii = 0
   ReDim aryTmp(Me.LAB306.Count - 1)
   For ii = 1 To Me.LAB306.Count
      With Me.LAB306.Item(ii)
         Select Case UCase(FieldName)
            Case "EQPCD"
               aryTmp(ii - 1) = .EqpCd
            Case "TRANSTM"
               aryTmp(ii - 1) = .TransTm
            Case "TRANSDT"
               aryTmp(ii - 1) = .TransDt
            Case "TRANSNO"
               aryTmp(ii - 1) = .TransNo
            Case "WORKAREA"
               aryTmp(ii - 1) = .WorkArea
            Case "ACCDT"
               aryTmp(ii - 1) = .AccDt
            Case "ACCSEQ"
               aryTmp(ii - 1) = .AccNo
            Case "ACCNO"
               aryTmp(ii - 1) = .WorkArea & "-" & Mid(.AccDt, 3) & "-" & .AccNo
         End Select
      End With
   Next
   GetString = Join(aryTmp, "$")
   ReDim aryTmp(0)
End Function

Public Function IsEqpCd(ByVal EqpCd As String) As Boolean
   If Trim(EqpCd) = "" Then
      IsEqpCd = False
   Else
      IsEqpCd = OraDS.GetIsEqpCd(EqpCd)
   End If
End Function

Public Function GetStrEqpTrans() As String
Dim ii As Long
Dim aryTmp() As String
Dim strTransTm As String
   '
   If Me.LAB306.Count = 0 Then
      GetStrEqpTrans = ""
      Exit Function
   End If
   '
   ii = 0
   ReDim aryTmp(Me.LAB306.Count - 1)
   For ii = 1 To Me.LAB306.Count
      With Me.LAB306.Item(ii)
         aryTmp(ii - 1) = Format(.TransNo, "@@@@@@@") & _
            vbTab & .SpcYy & "-" & .SpcNo & vbTab & vbTab & .WorkArea & "-" & Mid(.AccDt, 3, 6) & "-" & .AccNo & vbTab & .StatFg
         'strTransTm = Mid(.TransTm, 1, 2) & ":" & Mid(.TransTm, 3, 2)
         'aryTmp(ii - 1) = Mid(.TransDt, 3, 8) & " " & strTransTm & _
            vbTab & .SpcYy & "-" & .SpcNo & vbTab & vbTab & .WorkArea & "-" & Mid(.AccDt, 3, 6) & "-" & .AccNo & vbTab & .StatFg
      End With
   Next
   GetStrEqpTrans = Join(aryTmp, vbNewLine)
   ReDim aryTmp(0)
   '
End Function


