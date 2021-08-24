VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCAccInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISCAccInfo.cls
'   작성자  : 이상대
'   내  용  : ClientDb의 접수내역 클래스
'   작성일  : 2004-04-29
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mItemSeq    As Long     '전송순서
Private mSpcNo      As String   '검체번호(바코드번호)
Private mPtId       As String   '환자ID
Private mName       As String   '이름
Private mSex        As String   '성별
Private mAgeDay     As Long     '일령
Private mDoctNm     As String   '처방의명
Private mDeptNm     As String   '진료과명
Private mWardNm     As String   '병동명
Private mSpcCd      As String   '검체코드
Private mSpcNm      As String   '검체명
Private mStatFg     As String   '응급여부 (0:비응급, 1:응급)
Private mQcFg       As String   'QC유무 (0:무, 1:유)
Private mCtrlCd     As String   'Control Code
Private mLevelCd    As String   'Level Code
Private mLotNo      As String   'Lot No
Private mSpcPos     As String   '장비에서 검체위치
Private mEqpCd      As String   '장비코드
Private mTransDt    As String   '전송일자
Private mTransTm    As String   '전송일시

Public Property Get ItemSeq() As Long
    ItemSeq = mItemSeq
End Property

Public Property Let ItemSeq(ByVal vData As Long)
    mItemSeq = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mSpcNo
End Property

Public Property Let SpcNo(ByVal vData As String)
    mSpcNo = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vData As String)
    mName = vData
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Let Sex(ByVal vData As String)
    mSex = vData
End Property

Public Property Get AgeDay() As Long
    AgeDay = mAgeDay
End Property

Public Property Let AgeDay(ByVal vData As Long)
    mAgeDay = vData
End Property

Public Property Get DoctNm() As String
    DoctNm = mDoctNm
End Property

Public Property Let DoctNm(ByVal vData As String)
    mDoctNm = vData
End Property

Public Property Get DeptNm() As String
    DeptNm = mDeptNm
End Property

Public Property Let DeptNm(ByVal vData As String)
    mDeptNm = vData
End Property

Public Property Get WardNm() As String
    WardNm = mWardNm
End Property

Public Property Let WardNm(ByVal vData As String)
    mWardNm = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mSpcCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mSpcCd = vData
End Property

Public Property Get SpcNm() As String
    SpcNm = mSpcNm
End Property

Public Property Let SpcNm(ByVal vData As String)
    mSpcNm = vData
End Property

Public Property Get StatFg() As String
    StatFg = mStatFg
End Property

Public Property Let StatFg(ByVal vData As String)
    mStatFg = vData
End Property

Public Property Get QcFg() As String
    QcFg = mQcFg
End Property

Public Property Let QcFg(ByVal vData As String)
    mQcFg = vData
End Property

Public Property Get SpcPos() As String
    SpcPos = mSpcPos
End Property

Public Property Let SpcPos(ByVal vData As String)
    mSpcPos = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get TransDt() As String
    TransDt = mTransDt
End Property

Public Property Let TransDt(ByVal vData As String)
    mTransDt = vData
End Property

Public Property Get TransTm() As String
    TransTm = mTransTm
End Property

Public Property Let TransTm(ByVal vData As String)
    mTransTm = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 접수내역 객체를 이용해 ClientDb에 결과저장
'   인수 :
'       - pAccInfo : 접수내역 객체
'       - pEqpCd   : 장비코드
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function SaveClientDb(ByVal pAccInfo As clsIISAccInfo, ByVal pEqpCd) As Boolean
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim objCResult  As clsIISCResult    'ClientDb 결과내역 클래스
    Dim arySQL()    As String           '결과입력에 사용되는 SQL 배열
    Dim i           As Long
    
    With pAccInfo
        mItemSeq = GetMaxItemSeq
        mSpcNo = .GetBarNo:     mPtId = .PtId
        mName = .Name:          mSex = .Sex
        mAgeDay = .AgeDay:      mDoctNm = .OrdDoctNm
        mDeptNm = .DeptNm:      mWardNm = .WardNm
        mSpcCd = .SpcCd:        mSpcNm = Replace$(.SpcNm, "'", "''")
        mStatFg = .StatFg:      mQcFg = IIf(.QcFg = "1", "1", "0")
        mCtrlCd = .CtrlCd:      mLevelCd = .LevelCd
        mLotNo = .LotNo:        mSpcPos = .SpcPos
        mEqpCd = pEqpCd:        mTransDt = Format$(Now, "YYYYMMDD")
        mTransTm = Format$(Time, "HHMMSS")

        Set objCResult = New clsIISCResult
        If mQcFg = "0" Then         '## 일반검체
            '## 1.ClientDb에 일반검사 접수내역을 입력하는 쿼리생성
            ReDim arySQL(1)
            arySQL(0) = GetCliRegisterSql
            
            '## 2.ClientDb에 일반검사 결과내역을 입력하는 쿼리생성
            i = 1
            ReDim Preserve arySQL(.Results.Count + 1)
            For Each objResult In .Results
                If objResult.Result <> "" Then
                    arySQL(i) = objCResult.GetCliRegisterSql(mItemSeq, objResult.IntNm.IntNm, _
                        objResult.TestCd, objResult.Result, objResult.HLDiv, objResult.DPDiv)
                    i = i + 1
                End If
            Next
        ElseIf mQcFg = "1" Then     '## QC검체
            '## 1.ClientDb에 QC검사 접수내역을 입력하는 쿼리생성
            ReDim arySQL(1)
            arySQL(0) = GetCliQCRegisterSql
            
            '## 2.ClientDb에 QC검사 결과내역을 입력하는 쿼리생성
            i = 1
            ReDim Preserve arySQL(.QCResults.Count + 1)
            For Each objQCResult In .QCResults
                If objQCResult.Result <> "" Then
                    arySQL(i) = objCResult.GetCliQCRegisterSql(mItemSeq, objQCResult.IntNm.IntNm, _
                        objQCResult.TestCd, objQCResult.Result, objQCResult.RADiv, objQCResult.Text)
                    i = i + 1
                End If
            Next
        End If
        
        Set objCResult = Nothing
        Set objQCResult = Nothing
        Set objResult = Nothing
    End With
    
On Error GoTo Errors
    CliCon.BeginTrans
    
    '## 3.ClientDb에 적용
    For i = 0 To UBound(arySQL)
        If arySQL(i) = "" Then Exit For
        CliCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    CliCon.CommitTrans
    SaveClientDb = True
    Exit Function

Errors:
    CliCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISCAccInfo", "SaveClientDb", Err.Description, Now
    Error.SetLog App.EXEName, "clsIISCAccInfo", "SaveClientDb", "SQL: " & arySQL(i), Now
End Function

'-----------------------------------------------------------------------------'
'   기능 : MAX(전송순서)+1 을 조회
'   반환 : MAX(전송순서)+1
'-----------------------------------------------------------------------------'
Public Function GetMaxItemSeq() As Long
    Dim Rs As ADODB.Recordset
    
    Set Rs = CliCon.Execute(mSQL.SelectMaxSeq, , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        If IsNull(Rs.Fields("MAXSEQ").Value) Then
            GetMaxItemSeq = 1
        Else
            GetMaxItemSeq = CLng(Rs.Fields("MAXSEQ").Value)
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISCAccInfo", "GetMaxItemSeq", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb에 일반검사 접수내역을 입력하는 쿼리조회
'-----------------------------------------------------------------------------'
Public Function GetCliRegisterSql() As String
    Dim SQL As String
    
    SQL = mSQL.InsertClientAccInfo(mItemSeq, mSpcNo, mPtId, mName, mSex, mAgeDay, mDoctNm, mDeptNm, _
            mWardNm, mSpcCd, mSpcNm, mStatFg, mQcFg, mSpcPos, mEqpCd, mTransDt, mTransTm)
            
    GetCliRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb에 QC검사 접수내역을 입력하는 쿼리조회
'-----------------------------------------------------------------------------'
Public Function GetCliQCRegisterSql() As String
    Dim SQL As String
    
    SQL = mSQL.InsertClientQCAccInfo(mItemSeq, mSpcNo, mPtId, mQcFg, mCtrlCd, mLevelCd, mLotNo, _
            mSpcPos, mEqpCd, mTransDt, mTransTm)
    
    GetCliQCRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : ClientDb 보관일이 지난 데이터 삭제
'   인수 :
'       - pEqpCd   : 장비코드
'       - pStoreDt : 보관일
'-----------------------------------------------------------------------------'
Public Function DelHisClientDb(ByVal pEqpCd As String, ByVal pStoreDt As Long) As Boolean
    Dim strBaseDt As String     '삭제기준일
    Dim SQL       As String
    
    strBaseDt = Format$(DateAdd("d", -pStoreDt, Now), "YYYYMMDD")
    
On Error GoTo Errors
    '## ACC204 테이블의 데이터는 Access에서 Cascade형식으로 삭제됨
    CliCon.BeginTrans
    CliCon.Execute mSQL.DeleteHisData(pEqpCd, strBaseDt), , adCmdText + adExecuteNoRecords
    CliCon.CommitTrans
    DelHisClientDb = True
    Exit Function

Errors:
    CliCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISCAccInfo", "DelHisClientDb", Err.Description, Now
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub
