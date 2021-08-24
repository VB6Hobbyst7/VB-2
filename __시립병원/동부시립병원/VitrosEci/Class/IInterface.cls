VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim iphase As Integer               ' 통신상태
Dim rcvbuf(1 To 50) As String       ' 결과받는 버퍼
' addRcvbuf 함수와 관련 1..50
Dim ibufcnt As Integer              ' 결과받는 버퍼 Index
Dim rcvbufs As String               ' 버퍼 통합
'Dim lEqpInfo As tEquipInfo          ' 통신구성정보
'Dim lWorklist As tWorklist          ' Worklist정보
'Dim lSendItem As tSendItem          ' Send Item
'Dim lServerRst As tServerRst        ' 서버결과값
Dim iSnd_Phase As Integer           ' 송신 통신상태
Dim iState As String                ' 통신 모드
Dim iFrameN As Integer              ' ASTM Protocol에서 송신시 Frame Number
Dim iBlockCnt As Integer            ' Block Count
Dim iack_etx As Integer             ' ASTM중 ACK/ETX를 확인자로 사용하는 IF

Public Property Get phase() As Variant
    phase = iphase
End Property

Public Property Let phase(ByVal vNewValue As Variant)
    iphase = vNewValue
End Property

' 장비통신구성 얻기
'Public Sub getEquipInfo(sEqpCd As String)
'
'    Dim cEqp As New clsILab006
'    With lEqpInfo
'        .bEqpInfo = False
'        .sEqpCd = ""
'        .sEqpNm = ""
'        .sPort = ""
'        .sBaudrate = ""
'        .sDataBit = ""
'        .sStopBit = ""
'        .sParityBit = ""
'        .sStoreDate = ""
'    End With
'
'    If gsServerDBNm = "SybaseSom" Then
'        Dim rss As DrSqlOcx.DrRecordSet
'        Set rss = cEqp.getEqpInfo_s(sEqpCd)
'
'        Do Until rss.EOF
'            With lEqpInfo
'                .bEqpInfo = True
'                .sEqpCd = rss.Fields("eqpcd").Value & ""
'                .sEqpNm = rss.Fields("eqpnm").Value & ""
'                .sPort = rss.Fields("Port").Value & ""
'                .sBaudrate = rss.Fields("baudrate").Value & ""
'                .sDataBit = rss.Fields("databit").Value & ""
'                .sStopBit = rss.Fields("stopbit").Value & ""
'                .sParityBit = rss.Fields("paritybit").Value & ""
'                .sStoreDate = rss.Fields("storedate").Value & ""
'            End With
'            Exit Do
'        Loop
'
'        Set rss = Nothing
'    Else
'        Dim rs As ADODB.Recordset
'        Set rs = cEqp.getEqpInfo(sEqpCd)
'
'        Do Until rs.EOF
'            With lEqpInfo
'                .bEqpInfo = True
'                .sEqpCd = rs!EqpCD
'                .sEqpNm = rs!eqpnm
'                .sPort = rs!Port & ""
'                .sBaudrate = rs!baudrate & ""
'                .sDataBit = rs!Databit & ""
'                .sStopBit = rs!Stopbit & ""
'                .sParityBit = rs!Paritybit & ""
'                .sStoreDate = rs!storedate & ""
'            End With
'            Exit Do
'        Loop
'
'        Set rs = Nothing
'    End If
'
'
'    ' 과거데이타삭제 처리
'    Call processDeleteFile(lEqpInfo.sStoreDate)
'
''    gsDbLoc = GetSetting("Interface", "DB", "위치", "")
'
''    MsgBox lEqpInfo.sStoreDate & "이전 데이타를 삭제하시겠습니까?", vbYesNo, "과거데이타삭제"
'
'End Sub

'Public Function getlEqpInfosPort() As String
'    getlEqpInfosPort = lEqpInfo.sPort
'End Function

'Public Function getlEqpInfosBaudrate() As String
'    getlEqpInfosBaudrate = lEqpInfo.sBaudrate
'End Function

'Public Function getlEqpInfosParityBit() As String
'    getlEqpInfosParityBit = lEqpInfo.sParityBit
'End Function
'
'Public Function getlEqpInfosDataBit() As String
'    getlEqpInfosDataBit = lEqpInfo.sDataBit
'End Function
'
'Public Function getlEqpInfosStopBit() As String
'    getlEqpInfosStopBit = lEqpInfo.sStopBit
'End Function

Public Property Get bufcnt() As Variant
    bufcnt = ibufcnt
End Property

Public Property Let bufcnt(ByVal vNewValue As Variant)
    ibufcnt = vNewValue
End Property

Public Sub clearRcvbuf()
    Dim i%
    For i = 1 To 50
        rcvbuf(i) = ""
    Next
End Sub

Public Sub addRcvbuf(sBufChar As String)
    If ibufcnt > 0 And ibufcnt <= 50 Then
        rcvbuf(ibufcnt) = rcvbuf(ibufcnt) & sBufChar
    End If
End Sub

Public Sub addBufs()
    Dim i%
    rcvbufs = ""
    
    For i = 1 To ibufcnt
        rcvbufs = rcvbufs & rcvbuf(i)
    Next
End Sub

Public Function getrcvbufs() As String
    getrcvbufs = rcvbufs
End Function

'Public Property Get lWorklistbWk() As Variant
'    lWorklistbWk = lWorklist.bWk
'End Property
'
'Public Property Let lWorklistbWk(ByVal vNewValue As Variant)
'    lWorklist.bWk = vNewValue
'End Property

'Public Sub setlWorklistsSpcYy(sSpcYy As String)
'    lWorklist.sSpcYy = sSpcYy
'End Sub
'
'Public Sub setlWorklistsSpcNo(sSpcNo As String)
'    lWorklist.sSpcNo = sSpcNo
'End Sub

'Public Sub getWkInfo()
'    Dim sPtSNo$
'    If gsServerDBNm = "SybaseSom" Then
'        Dim cLab201s As New clsILab201
'        Dim rss As DrSqlOcx.DrRecordSet
'        Set rss = cLab201s.getWkInfo_s(lWorklist.sSpcYy, lWorklist.sSpcNo)
'
'        lWorklist.bWk = False
'
'        Do Until rss.EOF
'            With lWorklist
'                .bWk = True
'                .sWorkArea = rss.Fields("workarea").Value & ""
'                .sWorkAreaNm = rss.Fields("field1").Value & ""
'                .sAccDt = rss.Fields("accdt").Value & ""
'                .sAccSeq = rss.Fields("accseq").Value & ""
'
'                .sPtId = rss.Fields("ptid").Value & ""
'
'                .sPtNm = rss.Fields("ptnt_nm").Value & ""
'                sPtSNo = rss.Fields("ptnt_prsn_no").Value & ""
'                .sPtSNo = sPtSNo
'                Select Case Mid(sPtSNo, 9, 1)
'                    Case "2"
'                        .sSex = "F"
'                    Case "1"
'                        .sSex = "M"
'                End Select
''                Call debuggetWkInfo
'            End With
'            Exit Do
'        Loop
'
'        If Not rss Is Nothing Then
'            Set rss = Nothing
'        End If
'    Else
'        Dim cLab201 As New clsILab201
'        Dim rs As ADODB.Recordset
'        Set rs = cLab201.getWkInfo(lWorklist.sSpcYy, lWorklist.sSpcNo)
'        lWorklist.bWk = False
'
'        Do Until rs.EOF
'            With lWorklist
'                .bWk = True
'                .sWorkArea = rs!WorkArea & ""
'                .sWorkAreaNm = rs!field1 & ""
'                .sAccDt = rs!AccDt & ""
'                .sAccSeq = rs!AccSeq & ""
'
'                .sPtId = rs!PtId & ""
'                Select Case gsServerDBNm
'                    Case "Sybase"
'                        .sPtNm = rs!ptnt_nm & ""
'                        sPtSNo = rs!ptnt_prsn_no & ""
'                        Select Case Mid(sPtSNo, 9, 1)
'                            Case "2"
'                                .sSex = "F"
'                            Case "1"
'                                .sSex = "M"
'                        End Select
'
'                    Case "Access"
'                        .sPtNm = rs!ptnm & ""
'                        .sSex = rs!SEX & ""
'                End Select
'            End With
'            Exit Do
'        Loop
'        If Not rs Is Nothing Then
'            Set rs = Nothing
'        End If
'
'    End If
'End Sub

'Public Sub getTestItem()
'    '_NotVerify_s
'    Dim cLab302 As New clsILab302
'    Dim saLab302PK(1 To 3) As String
'    Dim i%
'    i = 0
'    With lWorklist
'        saLab302PK(1) = .sWorkArea
'        saLab302PK(2) = .sAccDt
'        saLab302PK(3) = .sAccSeq
'        .sSpcCd = ""
'    End With
'    Dim bSpcCd As Boolean
'    bSpcCd = False
'    If gsServerDBNm = "SybaseSom" Then
'        Dim rss As DrSqlOcx.DrRecordSet
'        Set rss = cLab302.getServerRstNm_s(saLab302PK)
'        lServerRst.lCnt = rss.RecordCount
'
'        If lServerRst.lCnt < 1 Then
'            If Not rss Is Nothing Then
'                Set rss = Nothing
'            End If
'            Exit Sub
'        End If
'
'        With lServerRst
'            ReDim .sTestCd(1 To .lCnt) As String
'            ReDim .sTestNm(1 To .lCnt) As String
'        End With
'
'        Do Until rss.EOF
'            i = i + 1
'            With lServerRst
'                .sTestCd(i) = Trim(rss.Fields("testcd").Value & "")
'                .sTestNm(i) = Trim(rss.Fields("testnm").Value & "")
'                If bSpcCd = False Then
'                    If Trim(rss.Fields("spccd").Value & "") <> "" Then
'                        lWorklist.sSpcCd = Trim(rss.Fields("spccd").Value & "")
'                        bSpcCd = True
'                    End If
'                End If
'            End With
'            rss.MoveNext
'        Loop
'        If Not rss Is Nothing Then
'            Set rss = Nothing
'        End If
'        lServerRst.lCnt = i
'        Call debuggetWkInfo
'    Else
'        Dim rs As ADODB.Recordset
'        Set rs = cLab302.getServerRstNm(saLab302PK)
'
'        lServerRst.lCnt = rs.RecordCount
'
'        If lServerRst.lCnt < 1 Then
'            If Not rs Is Nothing Then
'                Set rs = Nothing
'            End If
'            Exit Sub
'        End If
'
'        With lServerRst
'            ReDim .sTestCd(1 To .lCnt) As String
'            ReDim .sTestNm(1 To .lCnt) As String
'        End With
'
'        Do Until rs.EOF
'            i = i + 1
'            With lServerRst
'                .sTestCd(i) = Trim(rs!Testcd & "")
'                .sTestNm(i) = Trim(rs!TestNm & "")
'                If bSpcCd = False Then
'                    If Trim(rs!SpcCd & "") <> "" Then
'                        lWorklist.sSpcCd = Trim(rs!SpcCd & "")
'                        bSpcCd = True
'                    End If
'                End If
'            End With
'            rs.MoveNext
'        Loop
'        If Not rs Is Nothing Then
'            Set rs = Nothing
'        End If
'        lServerRst.lCnt = i
'    End If
'End Sub
'
'Public Sub getTestItemNotVerify()
'    '_NotVerify_s
'    Dim cLab302 As New clsILab302
'    Dim saLab302PK(1 To 3) As String
'    Dim i%
'    i = 0
'    With lWorklist
'        saLab302PK(1) = .sWorkArea
'        saLab302PK(2) = .sAccDt
'        saLab302PK(3) = .sAccSeq
'        .sSpcCd = ""
'    End With
'    Dim bSpcCd As Boolean
'    bSpcCd = False
'    If gsServerDBNm = "SybaseSom" Then
'        Dim rss As DrSqlOcx.DrRecordSet
'        Set rss = cLab302.getServerRstNm_NotVerify_s(saLab302PK)
'        lServerRst.lCnt = rss.RecordCount
'
'        If lServerRst.lCnt < 1 Then
'            If Not rss Is Nothing Then
'                Set rss = Nothing
'            End If
'            Exit Sub
'        End If
'
'        With lServerRst
'            ReDim .sTestCd(1 To .lCnt) As String
'            ReDim .sTestNm(1 To .lCnt) As String
'        End With
'
'        Do Until rss.EOF
'            i = i + 1
'            With lServerRst
'                .sTestCd(i) = Trim(rss.Fields("testcd").Value & "")
'                .sTestNm(i) = Trim(rss.Fields("testnm").Value & "")
'                If bSpcCd = False Then
'                    If Trim(rss.Fields("spccd").Value & "") <> "" Then
'                        lWorklist.sSpcCd = Trim(rss.Fields("spccd").Value & "")
'                        bSpcCd = True
'                    End If
'                End If
'            End With
'            rss.MoveNext
'        Loop
'        If Not rss Is Nothing Then
'            Set rss = Nothing
'        End If
'        lServerRst.lCnt = i
'        Call debuggetWkInfo
'    Else
'        Dim rs As ADODB.Recordset
'        Set rs = cLab302.getServerRstNm(saLab302PK)
'
'        lServerRst.lCnt = rs.RecordCount
'
'        If lServerRst.lCnt < 1 Then
'            If Not rs Is Nothing Then
'                Set rs = Nothing
'            End If
'            Exit Sub
'        End If
'
'        With lServerRst
'            ReDim .sTestCd(1 To .lCnt) As String
'            ReDim .sTestNm(1 To .lCnt) As String
'        End With
'
'        Do Until rs.EOF
'            i = i + 1
'            With lServerRst
'                .sTestCd(i) = Trim(rs!Testcd & "")
'                .sTestNm(i) = Trim(rs!TestNm & "")
'                If bSpcCd = False Then
'                    If Trim(rs!SpcCd & "") <> "" Then
'                        lWorklist.sSpcCd = Trim(rs!SpcCd & "")
'                        bSpcCd = True
'                    End If
'                End If
'            End With
'            rs.MoveNext
'        Loop
'        If Not rs Is Nothing Then
'            Set rs = Nothing
'        End If
'        lServerRst.lCnt = i
'    End If
'End Sub
'
'Public Sub getTestItemCnt()
'    Dim cLab302 As New clsILab302
'    Dim rs As ADODB.Recordset
'    'Call setTestCd
'    Dim saLab302PK(1 To 3) As String
'    With lWorklist
'        saLab302PK(1) = .sWorkArea
'        saLab302PK(2) = .sAccDt
'        saLab302PK(3) = .sAccSeq
'    End With
'
'    Set rs = cLab302.getServerRstCnt(saLab302PK)
'
'    lServerRst.lCnt = 0
'    Do Until rs.EOF
'        lServerRst.lCnt = rs.Fields(0)
'        Exit Do
'    Loop
'    If Not rs Is Nothing Then
'        Set rs = Nothing
'    End If
'End Sub

'Public Function getlServerRstsTestcd(i As Integer) As String
'    getlServerRstsTestcd = lServerRst.sTestCd(i)
'End Function
'
'Public Function getlServerRstsTestNm(i As Integer) As String
'    getlServerRstsTestNm = lServerRst.sTestNm(i)
'End Function

'Public Function getlServerRstlCnt() As Long
'    getlServerRstlCnt = lServerRst.lCnt
'End Function

'Public Sub setTestItem(lInstTestlRegCnt As Long, lInstTestsTestcd() As String)
'    Dim i%, j%, k%
'    ReDim lSendItem.sItemPos(1 To lInstTestlRegCnt) As String
'    k = 0
'    With lServerRst
'        For i = 1 To .lCnt
'            For j = 1 To lInstTestlRegCnt
'                If lInstTestsTestcd(j - 1) <> "" Then
'                    If .sTestCd(i) = lInstTestsTestcd(j - 1) Then
'                        k = k + 1
'                        lSendItem.sItemPos(k) = j
'                        Exit For
'                    End If
'                End If
'            Next
'        Next
'    End With
'    lSendItem.lCnt = k
'End Sub
'
'Public Function getlSendItemlCnt() As Long
'    getlSendItemlCnt = lSendItem.lCnt
'End Function
'
'Public Function getlWorklistsSpcYy() As String
'    getlWorklistsSpcYy = lWorklist.sSpcYy
'End Function
'
'Public Function getlWorklistsSpcNo() As String
'    getlWorklistsSpcNo = lWorklist.sSpcNo
'End Function
'
'Public Function getlWorklistsWorkAreaNm() As String
'    getlWorklistsWorkAreaNm = lWorklist.sWorkAreaNm
'End Function
'
'Public Function getlWorklistsAccDt() As String
'    getlWorklistsAccDt = lWorklist.sAccDt
'End Function
'
'Public Function getlWorklistsAccSeq() As String
'    getlWorklistsAccSeq = lWorklist.sAccSeq
'End Function
'
'Public Function getlWorklistsPtId() As String
'    getlWorklistsPtId = lWorklist.sPtId
'End Function
'
'Public Function getlWorklistsPtNm() As String
'    getlWorklistsPtNm = lWorklist.sPtNm
'End Function
'
'Public Function getlWorklistsSpcCd() As String
'    getlWorklistsSpcCd = lWorklist.sSpcCd
'End Function
'
'Public Function getlWorklistsSex() As String
'    getlWorklistsSex = lWorklist.sSex
'End Function
'
'Public Function getlSendItemsItemPos(i As Integer) As String
'    getlSendItemsItemPos = lSendItem.sItemPos(i)
'End Function


Public Property Get Snd_Phase() As Variant
    Snd_Phase = iSnd_Phase
End Property

Public Property Let Snd_Phase(ByVal vNewValue As Variant)
    iSnd_Phase = vNewValue
End Property

Public Property Get State() As Variant
    State = iState
End Property

Public Property Let State(ByVal vNewValue As Variant)
    iState = vNewValue
End Property

Public Property Get FrameN() As Variant
    FrameN = iFrameN
End Property

Public Property Let FrameN(ByVal vNewValue As Variant)
    iFrameN = vNewValue
End Property

Public Sub addFrameN()
    iFrameN = iFrameN + 1
    If iFrameN = 8 Then
        iFrameN = 0
    End If
End Sub

Public Property Get BlockCnt() As Variant
    BlockCnt = iBlockCnt
End Property

Public Property Let BlockCnt(ByVal vNewValue As Variant)
    iBlockCnt = vNewValue
End Property

Public Sub rmRcvbufChkSum()
    rcvbuf(ibufcnt) = left$(rcvbuf(ibufcnt), Len(rcvbuf(ibufcnt)) - 4)
End Sub

Public Property Get ack_etx() As Variant
    ack_etx = iack_etx
End Property

Public Property Let ack_etx(ByVal vNewValue As Variant)
    iack_etx = vNewValue
End Property

'
'Private Sub processDeleteFile(sStoreDate As String)
'
'    Dim sDeleteDate$, iYesNo%
'
'    sDeleteDate = GetSetting("Interface", "DeleteDate", gEqpCd, "")
'
'    Dim sDate$
'
'    sDate = Format(gSysDate, "YYYYMMDD")
'
'    If sDeleteDate <> sDate Then
'        iYesNo = MsgBox(lEqpInfo.sStoreDate & "일 이전 데이타를 삭제하시겠습니까?", vbYesNo, "과거데이타삭제")
'        If iYesNo = vbYes Then
'            Call SaveSetting("Interface", "DeleteDate", gEqpCd, sDate)
'            Call deleteFile(sStoreDate)
'        End If
'
'    End If
'
'End Sub

'Private Sub deleteFile(sStoreDate As String)
'    Dim MyPath$, MyName$, tDate As Date, iDate As Integer, dBaseDate As Date
'    MyPath = gsDbLoc   ' 경로를 설정합니다.
'
'    dBaseDate = FileDateTime(MyPath & "\standard\client.mdb")
'
'    MyName = Dir(MyPath, vbDirectory)   ' 첫번째 항목을 검색합니다.
'    Do While MyName <> ""   ' 루프(loop)를 시작합니다.
'       ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
'       If MyName <> "." And MyName <> ".." Then
'          ' MyName이 디렉토리인지 확인하기 위해서 비트별(bitwise) 비교를 사용합니다.
'            If left(MyName, Len(gEqpCd)) = gEqpCd Then
'                If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory Then
'                    tDate = FileDateTime(MyPath & MyName)
'                    If dBaseDate <> tDate Then
'                        iDate = DateDiff("d", tDate, Now)
'                        If Val(sStoreDate) < iDate Then
''                            Debug.Print MyName & "     ";    ' 항목만 표시합니다
''                            Debug.Print tDate & "  day : ";
''                            Debug.Print iDate
'                            Kill MyPath & MyName
'                        End If
'                    End If
'                End If
'            End If   ' 그것은 디렉토리를 표시합니다.
'
'       End If
'       MyName = Dir   ' 다음 항목을 읽어들입니다.
'    Loop
'End Sub
'
'Public Sub debuggetWkInfo()
'    If gbDebug = True Then
'        Dim j%
'        Open "c:\Debug_" & gEqpCd & ".log" For Append As #2
'        Dim sPrt$, sTmp
'        sPrt = "@@@@@@@@@"
'
'        With lWorklist
'            sTmp = ""
'            sTmp = sTmp & Format(.sSpcCd, "@@@@")
'            sTmp = sTmp & Format(.sWorkArea, "@@@")
'
'            sTmp = sTmp & Format(.sWorkAreaNm, "@@@@@@@@@@@@@")
'            sTmp = sTmp & Format(.sAccDt, "@@@@@@@@@@")
'            sTmp = sTmp & Format(.sAccSeq, "@@@@@@@@@@")
'            sTmp = sTmp & Format(.sPtId, "@@@@@@@@@@@@")
'            sTmp = sTmp & Format(.sPtNm, "@@@@@@@@@@")
'            sTmp = sTmp & Format(.sPtSNo, "@@@@@@@@@@@@@@@@")
'            sTmp = sTmp & Format(.sSex, "@@")
'        End With
'        Print #2, sTmp
'        Close #2
'    End If
'End Sub
'
'Public Sub debuglServerRst()
'    Dim i%
'    With lServerRst
'        For i = 1 To .lCnt
'
''            Debug.Print .sArrPos(i) & "    ";
''            Debug.Print .sResult(i) & "    ";
''            Debug.Print .sPrtord(i) & "    "
'        Next
'    End With
'End Sub
