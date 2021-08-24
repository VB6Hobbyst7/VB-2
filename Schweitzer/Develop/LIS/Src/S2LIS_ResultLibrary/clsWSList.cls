VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWSList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LAB031"
Attribute VB_Ext_KEY = "Member1" ,"LAB301"
Option Explicit

'/* Data Fields Property */
Private mvarWorkDt As String
Private mvarWorkCd As String
Private mvarWorkSeq As String
Private mvarWorkTm As String
Private mvarWorkId As String
Private mvarWorkArea As String
Private mvarAccDt As String
Private mvarAccSeq As String
Private mvarStatFg As String
'
'/* Class Default Property */
Private mvarRecordCount As Long
Private mvarLAB301 As LAB301
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
    vData = mvarWorkDt & mvarWorkCd & mvarWorkSeq
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
   CopyDataToProperty Me.LAB301.Count
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

Public Property Get LAB301() As LAB301
    If mvarLAB301 Is Nothing Then
        Set mvarLAB301 = New LAB301
    End If
    Set LAB301 = mvarLAB301
End Property

Public Property Set LAB301(vData As LAB301)
    Set mvarLAB301 = vData
End Property

Private Sub Class_Initialize()
   '
   m_INDEX = 0
End Sub

Private Sub Class_Terminate()
   Set mvarLAB301 = Nothing
End Sub

Public Property Let RecordCount(ByVal vData As Long)
   If Me.LAB301.Count > 0 Then
      mvarRecordCount = Me.LAB301.Count
   Else
      mvarRecordCount = vData
   End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = mvarRecordCount
End Property

'/* Data Fields Property */
Public Property Let WorkTm(ByVal vData As String)
    mvarWorkTm = vData
End Property

Public Property Get WorkTm() As String
    WorkTm = mvarWorkTm
End Property

Public Property Let WorkId(ByVal vData As String)
    mvarWorkId = vData
End Property

Public Property Get WorkId() As String
    WorkId = mvarWorkId
End Property

Public Property Let WorkArea(ByVal vData As String)
    mvarWorkArea = vData
End Property

Public Property Get WorkArea() As String
    WorkArea = mvarWorkArea
End Property

Public Property Let WorkDt(ByVal vData As String)
    mvarWorkDt = vData
End Property

Public Property Get WorkDt() As String
    WorkDt = mvarWorkDt
End Property

Public Property Let WorkSeq(ByVal vData As String)
    mvarWorkSeq = vData
End Property

Public Property Get WorkSeq() As String
    WorkSeq = mvarWorkSeq
End Property

Public Property Let WorkCd(ByVal vData As String)
    mvarWorkCd = vData
End Property

Public Property Get WorkCd() As String
    WorkCd = mvarWorkCd
End Property

Public Property Get AccNo() As String
'Read Only Property
    AccNo = mvarWorkArea & "-" & mvarAccDt & "-" & mvarAccSeq
End Property

Public Property Let AccDt(ByVal vData As String)
    mvarAccDt = vData
End Property

Public Property Get AccDt() As String
    AccDt = mvarAccDt
End Property

Public Property Let AccSeq(ByVal vData As String)
    mvarAccSeq = vData
End Property

Public Property Get AccSeq() As String
    AccSeq = mvarAccSeq
End Property

Public Property Let StatFg(ByVal vData As String)
    mvarStatFg = vData
End Property

Public Property Get StatFg() As String
    StatFg = mvarStatFg
End Property


'/* Data Method */
Public Function LoadTable(Optional pWorkCd As String = "Default", _
                            Optional pWorkDt As String = "Default", _
                            Optional pFrWorkSeq As String = "Default", _
                            Optional pToWorkSeq As String = "Default")

    Dim DS As Recordset
   '
'   MsgBox "LoadTable clsWSList ", vbCritical
    
    If pWorkDt = "Default" Then
       pWorkDt = vbNullString
    End If
    If pWorkCd = "Default" Then
       pWorkCd = vbNullString
    End If
   
    On Error GoTo ErrorLoadTable
    If pFrWorkSeq = "Default" Then
       OraDS.GetLAB301 DS, pWorkDt, pWorkCd
    Else
       OraDS.GetLAB301 DS, pWorkDt, pWorkCd, pFrWorkSeq, pToWorkSeq
    End If
    Set mvarLAB301 = New LAB301
   '
    If DS.RecordCount = 0 Then
       mvarBOF = False
       mvarEOF = False
       Exit Function
    End If
   '
    DS.MoveFirst
    Do Until DS.EOF
        mvarWorkDt = "" & DS.Fields("WORKDT").Value
        mvarWorkCd = "" & DS.Fields("WORKCD").Value
        mvarWorkSeq = "" & DS.Fields("WORKSEQ").Value
        mvarWorkTm = "" & DS.Fields("WORKTM").Value
        mvarWorkId = "" & DS.Fields("WORKID").Value
        mvarWorkArea = "" & DS.Fields("WORKAREA").Value
        mvarAccDt = "" & DS.Fields("ACCDT").Value
        mvarAccSeq = "" & DS.Fields("ACCSEQ").Value
        mvarStatFg = "" & DS.Fields("STATFG").Value
       '
        'Collection Class Key Setting
        mvarKey = mvarWorkDt & mvarWorkCd & mvarWorkSeq
        mvarRecordCount = DS.RecordCount
        mvarLAB301.Add mvarWorkDt, mvarWorkCd, mvarWorkSeq, _
           mvarWorkTm, mvarWorkId, mvarWorkArea, mvarAccDt, _
           mvarAccSeq, mvarStatFg, mvarRecordCount, mvarLAB301, mvarKey
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
   'With LAB301.Item(index)
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_INDEX = 0
         Else
            m_INDEX = Me.RecordCount + 1
         End If
         mvarWorkDt = 0
         mvarWorkCd = vbNullString
         mvarWorkSeq = vbNullString
         mvarWorkTm = vbNullString
         mvarWorkId = vbNullString
         mvarWorkArea = vbNullString
         mvarAccDt = vbNullString
         mvarAccSeq = vbNullString
         mvarKey = vbNullString                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB301.Count
         m_INDEX = index
   Else
      With Me.LAB301.Item(index)
         mvarWorkDt = .WorkDt
         mvarWorkCd = .WorkCd
         mvarWorkSeq = .WorkSeq
         mvarWorkDt = .WorkDt
         mvarWorkTm = .WorkTm
         mvarWorkId = .WorkId
         mvarWorkArea = .WorkArea
         mvarAccDt = .AccDt
         mvarAccSeq = .AccSeq
         mvarKey = .Key                          'Collection Class Key Setting
         mvarRecordCount = Me.LAB301.Count
         m_INDEX = index
      End With
   End If
   '
   With Me.LAB301
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
Dim objTmp As clsWSList
   '
   If Me.LAB301.Count = 0 Then
      GetString = ""
      Exit Function
   End If
   ii = 0
   ReDim aryTmp(Me.LAB301.Count)
   For Each objTmp In Me.LAB301
      With objTmp
         Select Case UCase(FieldName)
            Case "WORKCD"
               aryTmp(ii) = .WorkCd
            Case "WORKSEQ"
               aryTmp(ii) = .WorkSeq
            Case "WORKTM"
               aryTmp(ii) = .WorkTm
            Case "WORKID"
               aryTmp(ii) = .WorkId
            Case "WORKAREA"
               aryTmp(ii) = .WorkArea
            Case "ACCDT"
               aryTmp(ii) = .AccDt
            Case "ACCSEQ"
               aryTmp(ii) = .AccSeq
            Case "ACCNO"
               aryTmp(ii) = .AccNo
         End Select
      End With
      ii = ii + 1
   Next
   GetString = Join(aryTmp, "$")
   ReDim aryTmp(0)
End Function

Public Function IsWorkCd(ByVal WorkCd As String) As Boolean
   If Trim(WorkCd) = "" Then
      IsWorkCd = False
   Else
      IsWorkCd = OraDS.GetIsWorkCd(WorkCd)
   End If
End Function

Public Function GetStrWSList() As String
Dim ii As Long
Dim aryTmp() As String
Dim objTmp As clsWSBuild
   '
   If Me.LAB301.Count = 0 Then
      GetStrWSList = ""
      Exit Function
   End If
   '
   ii = 0
   ReDim aryTmp(Me.LAB301.Count - 1)
   For Each objTmp In Me.LAB301
      With objTmp
         aryTmp(ii) = .WorkSeq & vbTab & .WorkArea & "-" & Mid(.AccDt, 3, 8) & "-" & .AccSeq & vbTab & " "
      End With
      ii = ii + 1
   Next
   GetStrWSList = Join(aryTmp, vbNewLine)
   ReDim aryTmp(0)
   '
End Function

Public Sub LoadWorkSheetCode(ByVal pBuildingCd As String, ByRef lstWSCode As Object)

    Dim rsWSCode As Recordset
    Dim sSqlGetWSCode As String
    Dim I%
    Dim objSQL As New clsLisSqlResult

    Set rsWSCode = New Recordset
    rsWSCode.Open objSQL.GetWorksheetCode(pBuildingCd), dbconn

    If rsWSCode.EOF = True Then
        MsgBox " worksheet code가 존재하지 않습니다."
        GoTo NoData
    End If

    With lstWSCode
        rsWSCode.MoveFirst
'        For i = 0 To rsWSCode.RecordCount - 1
        While Not rsWSCode.EOF
            .AddItem "" & rsWSCode.Fields("WorkCd").Value & vbTab & _
                     "" & rsWSCode.Fields("WorkNm").Value & vbTab & _
                     "" & rsWSCode.Fields("WorkArea").Value & vbTab, I
                     I = I + 1
            rsWSCode.MoveNext
        Wend
'        Next i
    End With

NoData:
    Set rsWSCode = Nothing
    Set objSQL = Nothing
End Sub


