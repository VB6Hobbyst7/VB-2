VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWSBuild"
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
Private mvarAccNo As String
Private mvarAccDt As String
Private mvarAccSeq As String
Private mvarStatFg As String
Private mvarRcvDt As String
Private mvarRcvTm As String
Private mvarLastWorkSeq As Integer
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

Private mvarLoadAllData As Boolean


Private gLastDt As String
Private gLastTm As String

'WorkSheet Sorting 을 위해서 추가 하였음
Private mvarLastDt As String
Private mvarLastTm As String

Public Property Get GetLastDt() As String
    GetLastDt = mvarLastDt
End Property
Public Property Get GetLastTm() As String
    GetLastTm = mvarLastTm
End Property


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

Public Property Let LoadAllData(ByVal vData As Boolean)
    mvarLoadAllData = vData
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
    Set mvarLAB301 = New LAB301
    m_INDEX = 0
    gLastDt = ""
    gLastTm = ""

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
    mvarAccNo = mvarWorkArea & "-" & Mid(mvarAccDt, 3, 6) & "-" & mvarAccSeq
    AccNo = mvarAccNo
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

Public Property Let RcvDt(ByVal vData As String)
    mvarRcvDt = vData
End Property

Public Property Get RcvDt() As String
    RcvDt = mvarRcvDt
End Property

Public Property Let RcvTm(ByVal vData As String)
    mvarRcvTm = vData
End Property

Public Property Get RcvTm() As String
    RcvTm = mvarRcvTm
End Property


Public Function RemoveItem(ByVal sKey As String)
    If gLastDt > mvarLAB301(sKey).RcvDt Then
        gLastDt = mvarLAB301(sKey).RcvDt
        gLastTm = CStr(Val(mvarLAB301(sKey).RcvTm) - 1)
    End If
    If gLastDt = mvarLAB301(sKey).RcvDt Then
        If gLastTm > mvarLAB301(sKey).RcvTm Then
            gLastTm = CStr(Val(mvarLAB301(sKey).RcvTm) - 1)
        End If
    End If
    mvarLAB301.Remove sKey
End Function


'/* Data Method */
Public Function LoadTable(Optional pOraSession As Object, _
                          Optional pDatabase As Variant, _
                          Optional pWorkCd As String = "Default", _
                          Optional pRcvDt As String = "Default", _
                          Optional pRcvTm As String = "Default", _
                          Optional pWorkId As String = "Default", _
                          Optional pBuildCd As String = "", _
                          Optional pLastDtTm As String = "", _
                          Optional pWorkArea As String = "", _
                          Optional pNewFg As Boolean = True, _
                          Optional pFromSeq As String = "1", _
                          Optional pToSeq As String = "9999", _
                          Optional pStatFg As String = "0")

    Dim DS As Recordset
    Dim lngWorkSeq As Long
   
'   MsgBox "LoadTable clsWSBuild ", vbCritical
   
    On Error GoTo ErrorLoadTable
    If pNewFg Then  '신규 Worksheet
        OraDS.GetWSBuild DS, pWorkCd, pRcvDt, pRcvTm, pBuildCd, pLastDtTm, pWorkArea, mvarLoadAllData, pStatFg
    Else
        If Val(pFromSeq) = 0 Then pFromSeq = "1"
        If Val(pToSeq) = 0 Then pToSeq = "9999"
        OraDS.GetLAB301 DS, pRcvDt, pWorkCd, pFromSeq, pToSeq, pStatFg
    End If
    Set mvarLAB301 = New LAB301
   '
    If DS.RecordCount = 0 Then
        mvarBOF = False
        mvarEOF = False
        LoadTable = False
        Exit Function
    End If
   '
    gLastDt = ""
    gLastTm = ""
    If pNewFg Then    '신규
        mvarWorkDt = Format(GetSystemDate, CS_DateDbFormat)
        mvarWorkTm = ""
        mvarWorkCd = pWorkCd
        mvarWorkId = pWorkId
        lngWorkSeq = StartWorkSeq(mvarWorkCd, mvarWorkDt)
    Else    '수정
        mvarWorkDt = pRcvDt
        mvarWorkTm = ""
        mvarWorkCd = pWorkCd
        mvarWorkId = pWorkId
        lngWorkSeq = pFromSeq
    End If
   '
    DS.MoveFirst
    Do Until DS.EOF
        If pNewFg Then    '신규
            mvarWorkSeq = CStr(lngWorkSeq)
        Else
            mvarWorkSeq = "" & DS.Fields("WORKSEQ").Value
        End If
        mvarWorkArea = "" & DS.Fields("WORKAREA").Value
        mvarAccDt = "" & DS.Fields("ACCDT").Value
        mvarAccSeq = "" & DS.Fields("ACCSEQ").Value
      
        '** 가장 마지막 접수일시 ***********************
        If gLastDt < "" & DS.Fields("RCVDT").Value Then
            gLastDt = "" & DS.Fields("RCVDT").Value
            gLastTm = "" & DS.Fields("RCVTM").Value
            mvarLastDt = gLastDt: mvarLastTm = gLastTm
        End If
        If gLastDt = "" & DS.Fields("RCVDT").Value Then
            If gLastTm < "" & DS.Fields("RCVTM").Value Then
                gLastTm = "" & DS.Fields("RCVTM").Value
                mvarLastTm = gLastTm
            End If
        End If
      
        mvarRcvDt = "" & DS.Fields("RCVDT").Value
        mvarRcvTm = "" & DS.Fields("RCVTM").Value
      '
        'Collection Class Key Setting
        mvarKey = mvarWorkDt & mvarWorkCd & mvarWorkSeq
      '
        mvarRecordCount = DS.RecordCount
        mvarLAB301.Add mvarWorkDt, mvarWorkCd, mvarWorkSeq, _
                       mvarWorkTm, mvarWorkId, mvarWorkArea, mvarAccDt, _
                       mvarAccSeq, mvarStatFg, mvarRecordCount, mvarLAB301, mvarKey, _
                       mvarRcvDt, mvarRcvTm
        lngWorkSeq = lngWorkSeq + 1
        DS.MoveNext
    Loop
   '
    MoveFirst
    OraErr.Reset
   
    DS.Close
    Set DS = Nothing
   
    LoadTable = True
    Exit Function
   '
ErrorLoadTable:
    Set DS = Nothing
    LoadTable = False
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


Public Function LoadTable_NEW(Optional pWorkCd As String = "Default", _
                              Optional pRcvDt As String = "Default", _
                              Optional pRcvTm As String = "Default", _
                              Optional pWorkId As String = "Default", _
                              Optional pBuildCd As String = "", _
                              Optional pLastDtTm As String = "", _
                              Optional pWorkArea As String = "", _
                              Optional pNewFg As Boolean = True, _
                              Optional pFromSeq As String = "1", _
                              Optional pToSeq As String = "9999", _
                              Optional pStatFg As String = "0", _
                              Optional pChkFg As String = "1")

    Dim DS As Recordset
    Dim lngWorkSeq As Long
   
'   MsgBox "LoadTable clsWSBuild ", vbCritical
   
    On Error GoTo ErrorLoadTable
    If pNewFg Then  '신규 Worksheet
        OraDS.GetWSBuild_NEW DS, pWorkCd, pRcvDt, pRcvTm, pBuildCd, pLastDtTm, pWorkArea, mvarLoadAllData, pStatFg, pChkFg
    Else
        If Val(pFromSeq) = 0 Then pFromSeq = "1"
        If Val(pToSeq) = 0 Then pToSeq = "9999"
        OraDS.GetLAB301 DS, pRcvDt, pWorkCd, pFromSeq, pToSeq, pStatFg
    End If
    Set mvarLAB301 = New LAB301
   '
    If DS.RecordCount = 0 Then
        mvarBOF = False
        mvarEOF = False
        LoadTable = False
        Exit Function
    End If
   '
    gLastDt = ""
    gLastTm = ""
    If pNewFg Then    '신규
        mvarWorkDt = Format(GetSystemDate, CS_DateDbFormat)
        mvarWorkTm = ""
        mvarWorkCd = pWorkCd
        mvarWorkId = pWorkId
        lngWorkSeq = StartWorkSeq(mvarWorkCd, mvarWorkDt)
    Else    '수정
        mvarWorkDt = pRcvDt
        mvarWorkTm = ""
        mvarWorkCd = pWorkCd
        mvarWorkId = pWorkId
        lngWorkSeq = pFromSeq
    End If
   '
    DS.MoveFirst
    Do Until DS.EOF
        If pNewFg Then    '신규
            mvarWorkSeq = CStr(lngWorkSeq)
        Else
            mvarWorkSeq = "" & DS.Fields("WORKSEQ").Value
        End If
        mvarWorkArea = "" & DS.Fields("WORKAREA").Value
        mvarAccDt = "" & DS.Fields("ACCDT").Value
        mvarAccSeq = "" & DS.Fields("ACCSEQ").Value
      
        '** 가장 마지막 접수일시 ***********************
        If gLastDt < "" & DS.Fields("RCVDT").Value Then
            gLastDt = "" & DS.Fields("RCVDT").Value
            gLastTm = "" & DS.Fields("RCVTM").Value
            mvarLastDt = gLastDt: mvarLastTm = gLastTm
            
        End If
        If gLastDt = "" & DS.Fields("RCVDT").Value Then
            If gLastTm < "" & DS.Fields("RCVTM").Value Then
                gLastTm = "" & DS.Fields("RCVTM").Value
                mvarLastTm = gLastTm
            End If
        End If
      
        mvarRcvDt = "" & DS.Fields("RCVDT").Value
        mvarRcvTm = "" & DS.Fields("RCVTM").Value
      '
        'Collection Class Key Setting
        mvarKey = mvarWorkDt & mvarWorkCd & mvarWorkSeq
      '
        mvarRecordCount = DS.RecordCount
        mvarLAB301.Add mvarWorkDt, mvarWorkCd, mvarWorkSeq, _
                       mvarWorkTm, mvarWorkId, mvarWorkArea, mvarAccDt, _
                       mvarAccSeq, mvarStatFg, mvarRecordCount, mvarLAB301, mvarKey, _
                       mvarRcvDt, mvarRcvTm
        lngWorkSeq = lngWorkSeq + 1
        DS.MoveNext
    Loop
   '
    MoveFirst
    OraErr.Reset
   
    DS.Close
    Set DS = Nothing
   
    LoadTable_NEW = True
    Exit Function
   '
ErrorLoadTable:
    Set DS = Nothing
    LoadTable_NEW = False
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


Public Function LoadTable_NEW_2014(Optional pWorkCd As String = "Default", _
                              Optional pRcvDt As String = "Default", _
                              Optional pRcvTm As String = "Default", _
                              Optional pWorkId As String = "Default", _
                              Optional pPoct As String = "Default", _
                              Optional pBuildCd As String = "", _
                              Optional pLastDtTm As String = "", _
                              Optional pWorkArea As String = "", _
                              Optional pNewFg As Boolean = True, _
                              Optional pFromSeq As String = "1", _
                              Optional pToSeq As String = "9999", _
                              Optional pStatFg As String = "0", _
                              Optional pChkFg As String = "1")

    Dim DS As Recordset
    Dim lngWorkSeq As Long
   
   
    On Error GoTo ErrorLoadTable_2014
    If pNewFg Then  '신규 Worksheet
        OraDS.GetWSBuild_NEW_2014 DS, pPoct, pWorkCd, pRcvDt, pRcvTm, pBuildCd, pLastDtTm, pWorkArea, mvarLoadAllData, pStatFg, pChkFg
    Else
        If Val(pFromSeq) = 0 Then pFromSeq = "1"
        If Val(pToSeq) = 0 Then pToSeq = "9999"
        OraDS.GetLAB301 DS, pRcvDt, pWorkCd, pFromSeq, pToSeq, pStatFg
    End If
    Set mvarLAB301 = New LAB301
   '
    If DS.RecordCount = 0 Then
        mvarBOF = False
        mvarEOF = False
        LoadTable_NEW_2014 = False
        Exit Function
    End If
   '
    gLastDt = ""
    gLastTm = ""
    If pNewFg Then    '신규
        mvarWorkDt = Format(GetSystemDate, CS_DateDbFormat)
        mvarWorkTm = ""
        mvarWorkCd = pWorkCd
        mvarWorkId = pWorkId
        lngWorkSeq = StartWorkSeq(mvarWorkCd, mvarWorkDt)
    Else    '수정
        mvarWorkDt = pRcvDt
        mvarWorkTm = ""
        mvarWorkCd = pWorkCd
        mvarWorkId = pWorkId
        lngWorkSeq = pFromSeq
    End If
   '
    DS.MoveFirst
    Do Until DS.EOF
        If pNewFg Then    '신규
            mvarWorkSeq = CStr(lngWorkSeq)
        Else
            mvarWorkSeq = "" & DS.Fields("WORKSEQ").Value
        End If
        mvarWorkArea = "" & DS.Fields("WORKAREA").Value
        mvarAccDt = "" & DS.Fields("ACCDT").Value
        mvarAccSeq = "" & DS.Fields("ACCSEQ").Value
      
        '** 가장 마지막 접수일시 ***********************
        If gLastDt < "" & DS.Fields("RCVDT").Value Then
            gLastDt = "" & DS.Fields("RCVDT").Value
            gLastTm = "" & DS.Fields("RCVTM").Value
            mvarLastDt = gLastDt: mvarLastTm = gLastTm
            
        End If
        If gLastDt = "" & DS.Fields("RCVDT").Value Then
            If gLastTm < "" & DS.Fields("RCVTM").Value Then
                gLastTm = "" & DS.Fields("RCVTM").Value
                mvarLastTm = gLastTm
            End If
        End If
      
        mvarRcvDt = "" & DS.Fields("RCVDT").Value
        mvarRcvTm = "" & DS.Fields("RCVTM").Value
      '
        'Collection Class Key Setting
        mvarKey = mvarWorkDt & mvarWorkCd & mvarWorkSeq
      '
        mvarRecordCount = DS.RecordCount
        mvarLAB301.Add mvarWorkDt, mvarWorkCd, mvarWorkSeq, _
                       mvarWorkTm, mvarWorkId, mvarWorkArea, mvarAccDt, _
                       mvarAccSeq, mvarStatFg, mvarRecordCount, mvarLAB301, mvarKey, _
                       mvarRcvDt, mvarRcvTm
        lngWorkSeq = lngWorkSeq + 1
        DS.MoveNext
    Loop
   '
    MoveFirst
    OraErr.Reset
   
    DS.Close
    Set DS = Nothing
   
    LoadTable_NEW_2014 = True
    Exit Function
   '
ErrorLoadTable_2014:
    Set DS = Nothing
    LoadTable_NEW_2014 = False
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
        mvarWorkDt = vbNullString
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
                Case "WorkCd"
                   aryTmp(ii) = .WorkCd
                Case "WorkSeq"
                   aryTmp(ii) = .WorkSeq
                Case "WorkTm"
                   aryTmp(ii) = .WorkTm
                Case "WorkId"
                   aryTmp(ii) = .WorkId
                Case "WorkArea"
                   aryTmp(ii) = .WorkArea
                Case "AccDt"
                   aryTmp(ii) = .AccDt
                Case "AccSeq"
                   aryTmp(ii) = .AccSeq
            End Select
        End With
       ii = ii + 1
    Next
    GetString = Join(aryTmp, "$")
   ReDim aryTmp(0)
End Function

Public Function Insert(ByVal WorkTm As String, _
                        ByVal AddAccNo As String, _
                        Optional ByVal pWorkCd As String = "Default", _
                        Optional ByVal pWorkDt As String = "Default", _
                        Optional ByVal pWorkId As String = "Default", _
                        Optional ByRef pPrgBar As Object, _
                        Optional ByVal pNewFg As Boolean = True) As Boolean
'Add LAB031 to Collection Classs
    Dim blnFirst        As Boolean
    Dim blnSqlErr       As Boolean
    
    Dim strSQL          As String
    Dim strWorkDt       As String
    Dim strWorkId       As String
    Dim strWorkCd       As String
    Dim strBuildList    As String
    Dim aryAddAccNo()   As String
    Dim aryAddAccList() As String
    
    Dim ii              As Integer
    Dim lngStartWSSeq   As Long
   '/Insert
   
    Insert = True
    strBuildList = ""
    
    If pNewFg Then
        With Me
            blnFirst = True
            .MoveFirst
            Do Until .EOF
                If blnFirst = True Then
                    '접수번호 & Work No & 추가여부
                    strBuildList = DBAccNo(.AccNo) & COL_DIV & "" & COL_DIV & ""
                    lngStartWSSeq = Val(.WorkSeq)
                    strWorkDt = .WorkDt
                    strWorkId = .WorkId
                    strWorkCd = .WorkCd
                    blnFirst = False
                Else
                    strBuildList = strBuildList & vbTab & DBAccNo(.AccNo) & COL_DIV & "" & COL_DIV & ""
                End If
                .MoveNext
            Loop
        End With
    End If
   
    '저장할 웍슅 내역이 없는경우 Return... 2000.6.16 김미경
    If Trim(strBuildList) = "" And Trim(AddAccNo) = "" Then Exit Function
   '
    If AddAccNo <> "" Then
        aryAddAccList = Split(AddAccNo, vbTab)
        For ii = 0 To UBound(aryAddAccList)
            aryAddAccNo = Split(aryAddAccList(ii), COL_DIV)
            aryAddAccNo(0) = DBAccNo(aryAddAccNo(0))
            aryAddAccList(ii) = Join(aryAddAccNo, COL_DIV)
        Next ii
        AddAccNo = Join(aryAddAccList, vbTab)
        If strBuildList = "" Then
            strBuildList = AddAccNo
            lngStartWSSeq = StartWorkSeq(pWorkCd, pWorkDt)
            strWorkDt = pWorkDt
            strWorkId = pWorkId
            strWorkCd = pWorkCd
        Else
            strBuildList = strBuildList & vbTab & AddAccNo
        End If
    End If
   
    OraDS.BuildWorkSheet strWorkDt, strWorkCd, WorkTm, _
                        lngStartWSSeq, strWorkId, strBuildList, _
                        gLastDt, gLastTm, pPrgBar
   '
    With OraDS
        If .LastServerErr <> 0 Then
            glngErrorNo = .LastServerErr
            gstrErrorMsg = .LastServerErrText
            Insert = False
            Exit Function
        End If
    End With
   '
    OraErr.Reset
    Insert = True
    
End Function

Public Function GetStrBuildList() As String
    
    Dim ii As Long
    Dim aryTmp() As String
    Dim objTmp As clsWSBuild
   '
    If Me.LAB301.Count = 0 Then
       GetStrBuildList = ""
       Exit Function
    End If
    '
    ii = 0
    GetStrBuildList = ""
    ReDim aryTmp(Me.LAB301.Count - 1)
    For Each objTmp In Me.LAB301
        With objTmp
            aryTmp(ii) = .WorkSeq & vbTab & .WorkArea & "-" & _
                        Mid(.AccDt, 3, 6) & "-" & .AccSeq
        End With
        ii = ii + 1
    Next
    GetStrBuildList = Join(aryTmp, vbNewLine)
    ReDim aryTmp(0)
   '
End Function

Public Function IsWorkCd(ByVal WorkCd As String) As Boolean
    If Trim(WorkCd) = "" Then
       IsWorkCd = False
    Else
       IsWorkCd = OraDS.GetIsWorkCd(WorkCd)
    End If
End Function

Public Function GetAddInfo(ByRef pOraSession As Object, _
                            ByRef pDatabase As Variant, ByVal AccNo As String, _
                            ByVal pWorkCd As String, ByVal pRcvDt As String, _
                            ByVal pRcvTm As String) As String
    Dim objLAB032   As clsComcode032
    Dim strTmp      As String
    Dim strSpecimen As String
    Dim strColDt    As String
    Dim strTest     As String
    Dim aryTmp()    As String
    Dim strWorkInfo As String
    
    Dim ii          As Integer
   '
'    MsgBox "GetAddInfo clsWSBuild", vbCritical
    
    GetAddInfo = OraDS.GetAddInfo(AccNo, pWorkCd, pRcvDt, pRcvTm)
    If GetAddInfo = "" Then Exit Function
    '
    Set objLAB032 = New clsComcode032
    strTmp = medGetP(GetAddInfo, 1, "$")
    strWorkInfo = medGetP(GetAddInfo, 3, "$")
    strSpecimen = medGetP(strTmp, 1, ",")
    strColDt = medGetP(strTmp, 2, ",")
    
    With objLAB032
        .LoadTable LC3_Specimen, , strSpecimen
        If .RecordCount > 0 Then
            strSpecimen = .Field3
        End If
    End With
    '
    strTest = medGetP(GetAddInfo, 2, "$")
    aryTmp = Split(strTest, vbTab)
    For ii = 0 To UBound(aryTmp)
        aryTmp(ii) = OraDS.GetTestNm(aryTmp(ii), strColDt)
    Next ii
    strTest = Join(aryTmp, ",")
    
    GetAddInfo = strSpecimen & vbTab & strTest
   '
   
   GetAddInfo = strSpecimen & vbTab & strTest & vbTab & strWorkInfo
End Function

Public Function StartWorkSeq(ByVal WorkCd As String, _
                             ByVal WorkDt As String) As Long
   '
    StartWorkSeq = Val(OraDS.GetWSLastSeq(WorkCd, WorkDt))
   '
End Function

Public Function Delete(ByVal DelWorkSeq As String, _
                       ByVal pWorkCd As String, _
                       ByVal pWorkDt As String) As Boolean
    
    Dim blnSqlErr As Boolean
   
    Delete = True
   
    If Trim(DelWorkSeq) = "" Then Exit Function
   '
On Error GoTo Err_Trap

    Delete = OraDS.DeleteWorkSheet(pWorkDt, pWorkCd, DelWorkSeq)
   '
    With OraDS
        If .LastServerErr <> 0 Then
            glngErrorNo = .LastServerErr
            gstrErrorMsg = .LastServerErrText
            Delete = False
            Exit Function
        End If
    End With
   '
    OraErr.Reset
    Exit Function
    
Err_Trap:
    Delete = False
End Function

Public Function WorkHistoryDelete(ByVal WorkCd As String, ByVal WorkDt As String) As Boolean
'기존 WorkSheet내역 삭제
'2001/10/17울산동강병원
    
    Dim sSQL As String
    
On Error GoTo DELETE_ERROR
    sSQL = " delete from " & T_LAB301 & _
           " where " & DBW("workcd=", WorkCd) & _
           " and " & DBW("workdt<=", WorkDt)
    
    DBConn.BeginTrans   '2001-11-07 추가
    DBConn.Execute sSQL
    WorkHistoryDelete = True
    DBConn.CommitTrans
    Exit Function
    
DELETE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
    
End Function
