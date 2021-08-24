VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWBatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsResultPt"
Option Explicit
''결과등록과 관련된 기본 환자정보 및 환자별 장비결과 전송내역을 다룬다.
''
''/* Class Default Property */
'Private mvarRecordCount As Long
'Private mvarTestCount As Long
'Private mvarEOF As Boolean
'Private mvarBOF As Boolean
'Private mvarKey As String
'Private mvarErrNo As Long
'Private mvarErrText As String
''Private mvarResult As clsResultCol
'Private m_INDEX As Long
''
''/* Data Fields Property */
'Private mvarPtId As String
'Private mvarPtNm As String
'Private mvarSex As String
'Private mvarAge As String
'Private mvarAgeDay As String
'Private mvarSpcNm As String
'
'Private mvarDOB As String
'Private mvarOrdDoctNm As String
'Private mvarDeptCd As String
'Private mvarDeptNm As String
'Private mvarWardId As String
'Private mvarOrdDoct As String
'Private mvarSpcYy As String
'Private mvarSpcNo As String
'Private mvarWorkArea As String
'Private mvarAccSeq As String
'Private mvarStsCd As String
'Private mvarBedInDt As String
'Private mvarMajDoct As String
'Private mvarColDt As String
'Private mvarAccDt As String
'Private mvarRmkCd As String
'Private mvarRmkNm As String
'Private mvarItemCnt As Long
'Private mvarReqTotCnt As Long
'Private mvarReqInputCnt As Long
'Private mvarAltTotCnt As Long
'Private mvarAltInputCnt As Long
'Private mvarFootNoteFg As String
'Private mvarFootNote As String
'Private mvarMFootNote As String
'Private mvarQcFg As String
''
''/* Information Fields Property */
'Private mvarAccNo As String
'Private mvarSpcNum As String
'Private mvarSexAge As String
'Private mvarLocation As String
''
''/*  Patient's Test Info.(Result Info.) *'
'Private m_WorkArea As String
'Private m_AccDt As String
'Private m_AccSeq As String
'Private m_SpcYy As String
'Private m_SpcNo As String
'Private m_OrdDt As String
'Private m_OrdNo As String
'Private m_OrdSeq As String
'Private m_TestCd As String
'Private m_RstVal As String
'Private m_RstCd As String
'Private m_RstUnit As String
'Private m_HLDiv As String
'Private m_DPDiv As String
'Private m_SpcCd As String
'Private m_PtInfo As String
'Private m_LastRst As String
'Private m_LastVfyDt As String
'Private m_LastVfyTm As String
'Private m_LastVfyId As String
'Private m_VfyDt As String
'Private m_VfyTm As String
'Private m_VfyId As String
'Private m_MfyFg As String
'Private m_GrpFg As String
'Private m_TxtFg As String
'Private m_ExcFg As String
'Private m_AutoFg As String
'Private m_RstType As String
'Private m_RstDiv As String
'Private m_TextRst As String
'Private m_SuppText As String
'Private m_TxtType As String
'Private m_TestNm As String
'Private m_ColDt As String
'Private m_AvalVal As String
'Private m_PanicFg As String
'Private m_PanicFrVal As Double
'Private m_PanicToVal  As Double
'Private m_DeltaFg As String
'Private m_DeltaVal1 As Long
'Private m_DeltaVal2 As Long
'Private m_RefValFrom As Double
'Private m_RefValTo As Double
'Private m_RefCd As String
'Private m_EmpId As String
'Private m_DetailFg As String
'Private m_DataFg As String
'Private m_EqpCd As String
'Private m_OTmpCd As String
'Private m_EqpNm As String
'Private m_TransDtTm As String
'Private m_MRstVal As String
'Private m_MRstCd As String
'Private m_MHLDiv As String
'Private m_MDPDiv As String
'Private m_MfyRsn As String
'Private m_MRstType As String
'Private m_FootNote As String
'Private m_ReqAltCnt As String
'Private m_WorkSeq As String
''
''/* Condition Fields Property */
'' +mvarPtType-1.일반정보,2.일반정보+결과정보,3.일반정보+장비정보,
''                    4.일반정보+장비정보+결과정보
'Private mvarPtType As Integer
'Private m_PrgSts As Boolean
'Private mvarPrgBar As Object
''
''/* Help List Popup 위치
'Private mvarSsTop As Long
'Private mvarSsLeft As Long
'
''/* SPREAD Information */
''Private ssRst As Object
'Private SS_COL_TESTNM As Integer
'Private SS_COL_RESULT As Integer
'Private SS_COL_EC As Integer
'Private SS_COL_LASTRST As Integer
'Private SS_COL_RSTUNIT As Integer
'Private SS_COL_HLDIV As Integer
'Private SS_COL_DPDIV As Integer
'Private SS_COL_TXT As Integer
'Private SS_COL_GRPFG As Integer
'Private SS_COL_EQPNM As Integer
'Private SS_COL_REFVAL As Integer
'Private SS_COL_SCREEN_MAX As Integer
'Private SS_COL_SEQ As Integer
'Private SS_COL_ACCNO As Integer
'Private SS_COL_PTID As Integer
'Private SS_COL_PTNM As Integer
'Private SS_COL_SEXAGE As Integer
''
'
'Private mvarTextTipDelay As Long
'
'Private Const P_FreeRstLength = 30
'
'
'
''/* Event Information & ETC */
'Dim WithEvents objCodeList As clsCodeList
'Dim WithEvents mnuPopup As Menu
'Dim WithEvents mnuUtil1 As Menu
'Dim WithEvents mnuUtil2 As Menu
'
'Private ssRst           As Object
'Private objdic          As New clsDictionary
'Private ResultCol       As Integer
'
'Private Enum tblCol
'    tcNO = 1
'    tcLABNO
'    tcPTID
'    tcSpc
'    tcStat
'    tcHold
'End Enum
'
'
'
'
'
'
'
''Public Property Get Result() As clsResultCol
''    If mvarResult Is Nothing Then
''        Set mvarResult = New clsResultCol
''    End If
''    Set Result = mvarResult
''End Property
'
''Public Property Set Result(vData As clsResultCol)
''    Set mvarResult = vData
''End Property
'
'Public Property Let SexAge(ByVal vData As String)
'    mvarSexAge = vData
'End Property
'
'Public Property Get SexAge() As String
'    SexAge = mvarSexAge
'End Property
'
'Public Property Let RmkCd(ByVal vData As String)
'    mvarRmkCd = vData
'End Property
'
'Public Property Get RmkCd() As String
'    RmkCd = mvarRmkCd
'End Property
'
'Public Property Let RmkNm(ByVal vData As String)
'    mvarRmkNm = vData
'End Property
'
'Public Property Get RmkNm() As String
'    RmkNm = mvarRmkNm
'End Property
'
'Public Property Let OrdDoctNm(ByVal vData As String)
'    mvarOrdDoctNm = vData
'End Property
'
'Public Property Get OrdDoctNm() As String
'    OrdDoctNm = mvarOrdDoctNm
'End Property
'
'Public Property Let MajDoct(ByVal vData As String)
'    mvarMajDoct = vData
'End Property
'
'Public Property Get MajDoct() As String
'    MajDoct = mvarMajDoct
'End Property
'
'Public Property Let BedIndt(ByVal vData As String)
'    mvarBedInDt = vData
'End Property
'
'Public Property Get BedIndt() As String
'    BedIndt = mvarBedInDt
'End Property
'
'Public Property Let OrdDoct(ByVal vData As String)
'    mvarOrdDoct = vData
'End Property
'
'Public Property Get OrdDoct() As String
'    OrdDoct = mvarOrdDoct
'End Property
'
'Public Property Let coldt(ByVal vData As String)
'    mvarColDt = vData
'End Property
'
'Public Property Get coldt() As String
'    coldt = mvarColDt
'End Property
'
'Public Property Let dob(ByVal vData As String)
'    mvarDOB = vData
'End Property
'
'Public Property Get dob() As String
'    dob = mvarDOB
'End Property
'
'Public Property Let Age(ByVal vData As String)
'    mvarAge = vData
'End Property
'
'Public Property Get Age() As String
'    Age = mvarAge
'End Property
'
'Public Property Let AgeDay(ByVal vData As String)
'    mvarAgeDay = vData
'End Property
'
'Public Property Get AgeDay() As String
'    AgeDay = mvarAgeDay
'End Property
'
'Public Property Let SpcNm(ByVal vData As String)
'    mvarSpcNm = vData
'End Property
'
'Public Property Get SpcNm() As String
'    SpcNm = mvarSpcNm
'End Property
'
'Public Property Let AltTotCnt(ByVal vData As Long)
'    mvarAltTotCnt = vData
'End Property
'
'Public Property Get AltTotCnt() As Long
'    AltTotCnt = mvarAltTotCnt
'End Property
'
'Public Property Let AltInputCnt(ByVal vData As Long)
'    mvarAltInputCnt = vData
'End Property
'
'Public Property Get AltInputCnt() As Long
'    AltInputCnt = mvarAltInputCnt
'End Property
'
'Public Property Let ReqInputCnt(ByVal vData As Long)
'    mvarReqInputCnt = vData
'End Property
'
'Public Property Get ReqInputCnt() As Long
'    ReqInputCnt = mvarReqInputCnt
'End Property
'
'Public Property Get ItemTotCnt() As Long
'    ItemTotCnt = mvarReqTotCnt + mvarAltTotCnt
'End Property
'
'Public Property Let ReqTotCnt(ByVal vData As Long)
'    mvarReqTotCnt = vData
'End Property
'
'Public Property Get ReqTotCnt() As Long
'    ReqTotCnt = mvarReqTotCnt
'End Property
'
'Public Property Let sex(ByVal vData As String)
'    mvarSex = vData
'End Property
'
'Public Property Get sex() As String
'    sex = mvarSex
'End Property
'
'Public Property Let PtNm(ByVal vData As String)
'    mvarPtNm = vData
'End Property
'
'Public Property Get PtNm() As String
'    PtNm = mvarPtNm
'End Property
'
'Public Property Let wardid(ByVal vData As String)
'    mvarWardId = vData
'End Property
'
'Public Property Get wardid() As String
'    wardid = mvarWardId
'End Property
'
'Public Property Let DeptCd(ByVal vData As String)
'    mvarDeptCd = vData
'End Property
'
'Public Property Get DeptCd() As String
'    DeptCd = mvarDeptCd
'End Property
'
'Public Property Let DeptNm(ByVal vData As String)
'    mvarDeptNm = vData
'End Property
'
'Public Property Get DeptNm() As String
'    DeptNm = mvarDeptNm
'End Property
'
'Public Property Let PtType(ByVal vData As Integer)
'    mvarPtType = vData
'End Property
'
'Public Property Get PtType() As Integer
'    PtType = mvarPtType
'End Property
'
'Public Property Let ptid(ByVal vData As String)
'    mvarPtId = vData
'End Property
'
'Public Property Get ptid() As String
'    ptid = mvarPtId
'End Property
'
'Public Property Let stscd(ByVal vData As String)
'    mvarStsCd = vData
'End Property
'
'Public Property Get stscd() As String
'    stscd = mvarStsCd
'End Property
'
'Public Property Let spcyy(ByVal vData As String)
'    mvarSpcYy = vData
'End Property
'
'Public Property Get spcyy() As String
'    spcyy = mvarSpcYy
'End Property
'
'Public Property Let spcno(ByVal vData As String)
'    mvarSpcNo = vData
'End Property
'
'Public Property Get spcno() As String
'    spcno = mvarSpcNo
'End Property
'
'Public Property Let accdt(ByVal vData As String)
'    mvarAccDt = vData
'End Property
'
'Public Property Get accdt() As String
'    accdt = mvarAccDt
'End Property
'
'Public Property Let accseq(ByVal vData As String)
'    mvarAccSeq = vData
'End Property
'
'Public Property Get accseq() As String
'    accseq = mvarAccSeq
'End Property
'
'Public Property Let WorkArea(ByVal vData As String)
'    mvarWorkArea = vData
'End Property
'
'Public Property Get WorkArea() As String
'    WorkArea = mvarWorkArea
'End Property
'
'Public Property Let FootNoteFg(ByVal vData As String)
'    mvarFootNoteFg = vData
'End Property
'
'Public Property Get FootNoteFg() As String
'    FootNoteFg = mvarFootNoteFg
'End Property
'
'Public Property Let FootNote(ByVal vData As String)
'    mvarFootNote = vData
'End Property
'
'Public Property Get FootNote() As String
'    FootNote = mvarFootNote
'End Property
'
'Public Property Let QcFg(ByVal vData As String)
'    mvarQcFg = vData
'End Property
'
'Public Property Get QcFg() As String
'    QcFg = mvarQcFg
'End Property
'
'Public Property Let MFootNote(ByVal vData As String)
'    mvarMFootNote = vData
'End Property
'
'Public Property Get MFootNote() As String
'    MFootNote = mvarMFootNote
'End Property
'
'Public Property Get Version() As String
'    Version = "LisResultEntry.dll" & vbTab & "v" & App.Major & "." & App.Minor & "." & App.Revision & vbTab & "1999.10.25"
'End Property
'
'Public Property Let SsTop(ByVal vData As Long)
'    mvarSsTop = vData
'End Property
'
'Public Property Get SsTop() As Long
'    SsTop = mvarSsTop
'End Property
'
'Public Property Let SsLeft(ByVal vData As Long)
'    mvarSsLeft = vData
'End Property
'
'Public Property Get SsLeft() As Long
'    SsLeft = mvarSsLeft
'End Property
'
'
'Public Property Let AccNo(ByVal vData As String)
''    mvarWorkArea = medGetP(vData, 1, "-")
''    mvarAccDt = FormatAccDt(medGetP(vData, 2, "-"))
''    mvarAccSeq = Trim(medGetP(vData, 3, "-"))
''    mvarAccNo = mvarWorkArea & "-" & Trim(medGetP(vData, 2, "-")) & "-" & mvarAccSeq
'End Property
'
'Public Property Get AccNo() As String
'    mvarAccNo = mvarWorkArea & "-" & Mid(mvarAccDt, 3, 6) & "-" & mvarAccSeq
'    AccNo = mvarAccNo
'End Property
'
'Public Property Let SpcNum(ByVal vData As String)
'    mvarSpcYy = medGetP(vData, 1, "-")
'    mvarSpcNo = Format(medGetP(vData, 2, "-"), "########")
'    mvarSpcNum = vData
'End Property
'
'Public Property Get SpcNum() As String
'    mvarSpcNum = mvarSpcYy & "-" & mvarSpcNo
'    SpcNum = mvarSpcNum
'End Property
'
'Public Property Let Location(ByVal vData As String)
'    mvarLocation = vData
'End Property
'
'Public Property Get Location() As String
'    If mvarBedInDt = "" Then
'        mvarLocation = mvarDeptNm
'    Else
'        If mvarPtId = "" Then
'            mvarLocation = ""
'        Else
''            mvarLocation = OraDS.GetWardLocation(mvarPtId, mvarBedInDt)
'        End If
'    End If
'    Location = mvarLocation
'End Property
'
'Public Property Set PrgBar(ByRef vData As Object)
'    Set mvarPrgBar = vData
'    m_PrgSts = True
'End Property
'
'Public Property Get PrgBar() As Object
'    Set PrgBar = mvarPrgBar
'End Property
'
''/* Class Default Property */
'Public Property Get ErrNo() As Long
''    mvarErrNo = OraErr.Err
'    ErrNo = mvarErrNo
'End Property
'
'Public Property Get ErrText() As String
''    mvarErrText = OraErr.ErrText
'    ErrText = mvarErrText
'End Property
'
'Public Property Let Key(ByVal vData As String)
'    vData = m_TestCd
'    mvarKey = vData
'End Property
'
'Public Property Get Key() As String
'    Key = m_TestCd
'End Property
'
'Public Property Get BOF() As Boolean
'    BOF = mvarBOF
'End Property
'
'Public Property Get EOF() As Boolean
'    EOF = mvarEOF
'End Property
'
'Public Sub MoveFirst()
''Copy First Data Collection to Class Property
'    CopyDataToProperty 1
'    '
'End Sub
'
'Public Sub MoveLast()
''Copy Last Data Collection to Class Propertys
''    CopyDataToProperty Me.Result.Count
'    '
'End Sub
'
'Public Sub MoveNext()
''Copy Next Data Collection to Class Propertys
'    With Me
'        CopyDataToProperty (m_INDEX + 1)
'    End With
'    '
'End Sub
'
'Public Sub MovePrevious()
''Copy Previous Data Collection to Class Propertys
'    With Me
'        CopyDataToProperty (m_INDEX - 1)
'    End With
'    '
'End Sub
'
'Private Sub Class_Terminate()
''    Set mvarResult = Nothing
'    Set objCodeList = Nothing
'    Set mnuPopup = Nothing
'    Set mnuUtil1 = Nothing
'End Sub
'
'Public Property Let RecordCount(ByVal vData As Long)
'    If Me.RecordCount > 0 Then
'        mvarRecordCount = Me.RecordCount
'    Else
'        mvarRecordCount = vData
'    End If
'End Property
'
'Public Property Get RecordCount() As Long
'    RecordCount = mvarRecordCount
'End Property
'
'Public Property Let TestCount(ByVal vData As Long)
''    If Me.Result.Count > 0 Then
''        mvarTestCount = Me.Result.Count
''    Else
''        mvarTestCount = vData
''    End If
'End Property
'
'Public Property Get TestCount() As Long
''    mvarTestCount = Me.Result.Count
'    TestCount = mvarTestCount
'End Property
'
'Public Function WSDataExists(Optional pOraSession As Object, _
'                             Optional pDatabase As DrDatabase, _
'                             Optional pSetCode As String = "Default") As Boolean
'
'    Dim DS As RECORDSET
'
''    On Error GoTo ErrorLoadTable
''
'''    OraDS.CheckWSData DS, mvarWorkArea, mvarAccDt, mvarAccSeq, pSetCode
''    If DS.RecordCount > 0 Then
''        WSDataExists = True
''    Else
''        WSDataExists = False
''    End If
''    DS.RsClose
''    Set DS = Nothing
''    Exit Function
''
''ErrorLoadTable:
'''    OraErr.Reset
''    If OraDS.LastServerErr <> 0 Then
''        glngErrorNo = OraDS.LastServerErr
''        gstrErrorMsg = OraDS.LastServerErrText
''    Else
''        If IsObject(Err) Then
''            glngErrorNo = Err.Number
''            gstrErrorMsg = Err.Description
''        End If
''    End If
'    '
'End Function
'
'Public Function GetAccSeq(ByVal pWorkArea As String, ByVal pAccDt As String, _
'                          ByVal pAccseq As String, ByVal pOption As Integer, Optional pOraSession As Object, _
'                          Optional pDatabase As DrDatabase) As String
'
''    Dim DS As RECORDSET
''
''    On Error GoTo ErrorLoadTable
''
''    OraDS.GetAccSeq DS, pWorkArea, pAccDt, pAccseq, pOption
''    GetAccSeq = "" & DS.Fields("NextAccSeq").Value
''    DS.RsClose
''    Set DS = Nothing
''    Exit Function
''
''ErrorLoadTable:
''    OraErr.Reset
''    If OraDS.LastServerErr <> 0 Then
''        glngErrorNo = OraDS.LastServerErr
''        gstrErrorMsg = OraDS.LastServerErrText
''    Else
''        If IsObject(Err) Then
''            glngErrorNo = Err.Number
''            gstrErrorMsg = Err.Description
''        End If
''    End If
'    '
'End Function
'
'Public Function GetCalType(ByVal pTestCd As String, ByVal pSpcCd As String) As String
'
''    Dim DS As RECORDSET
''
''    On Error GoTo ErrorLoadTable
''
''    OraDS.GetCalType DS, pTestCd, pSpcCd
''    If Not DS.EOF Then
''        GetCalType = "" & DS.Fields("caltype").Value
''    Else
''        GetCalType = ""
''    End If
''    DS.RsClose
''    Set DS = Nothing
''    Exit Function
''
''ErrorLoadTable:
''    OraErr.Reset
''    If OraDS.LastServerErr <> 0 Then
''        glngErrorNo = OraDS.LastServerErr
''        gstrErrorMsg = OraDS.LastServerErrText
''    Else
''        If IsObject(Err) Then
''            glngErrorNo = Err.Number
''            gstrErrorMsg = Err.Description
''        End If
''    End If
'    '
'End Function
'
'
'
''/* Data Method */
'Public Function LoadTable(Optional pOraSession As Object, _
'                          Optional pDatabase As DrDatabase, _
'                          Optional pSetCode As String = "Default", _
'                          Optional pEmpId As String = "SYSTEM", _
'                          Optional pItemMode As String = "WorkSheet", _
'                          Optional pOcsPtCheck As String = "OCS", _
'                          Optional pBuildCd As String = "")
'
''    Dim DS As RECORDSET
''    Dim ES As RECORDSET
''    Dim RS As RECORDSET
''    Dim TS As RECORDSET
''    Dim strPtInfo As String
''   '
''    PrgSet 20
''
''    On Error GoTo ErrorLoadTable
''    If mvarPtType < 1 Or mvarPtType > 6 Then
''        mvarPtType = 1
''    End If
''    '
''    If mvarPtType <> RESULT_BY_ITEM Then
''        OraDS.GetResultPtAccNo mvarPtType, DS, mvarWorkArea, mvarAccDt, mvarAccSeq
''        PrgSet 40
''    End If
''    '
''    If mvarResult Is Nothing Then
''        Set mvarResult = New clsResultCol
''    Else
''        Set mvarResult = Nothing
''        Set mvarResult = New clsResultCol
''    End If
''    '
''    If mvarPtType <> RESULT_BY_ITEM Then
''        If DS.RecordCount = 0 Then
''            mvarBOF = False
''            mvarEOF = False
''            mvarRecordCount = 0
''            PrgSet 250
''            Exit Function
''        Else
''            mvarRecordCount = 1
''        End If
''    Else
''        mvarRecordCount = 1
''    End If
''    '
''    mvarTextTipDelay = 500
''    m_EmpId = pEmpId
''
''    If mvarPtType <> RESULT_BY_ITEM Then
''        DS.MoveFirst
''        mvarWorkArea = "" & DS.Fields("WORKAREA").Value
''        mvarAccSeq = "" & DS.Fields("ACCSEQ").Value
''        mvarAccDt = "" & DS.Fields("ACCDT").Value
''        mvarAccNo = mvarWorkArea & "-" & Mid(mvarAccDt, 3, 8) & "-" & mvarAccSeq
''        mvarStsCd = "" & DS.Fields("STSCD").Value
''        mvarPtId = "" & DS.Fields("PTID").Value
''        mvarAgeDay = "" & DS.Fields("AGEDAY").Value
''        mvarDeptCd = "" & DS.Fields("DEPTCD").Value
''        mvarWardId = "" & DS.Fields("WARDID").Value
''        mvarOrdDoct = "" & DS.Fields("ORDDOCT").Value
''        mvarMajDoct = "" & DS.Fields("MAJDOCT").Value
''        mvarBedInDt = "" & DS.Fields("BEDINDT").Value
''        mvarRmkCd = "" & DS.Fields("RMKCD").Value
''        mvarColDt = "" & DS.Fields("COLDT").Value
''        mvarSpcNm = "" & DS.Fields("SPCNM").Value
''        mvarQcFg = "" & DS.Fields("QCFG").Value
''        '
''        If pOcsPtCheck = "OCS" Then
''            '/* 외부DB에서 기본정보를 가져오는 경우(병원OCS) :길병원
''            '/* GetOcsPtInfo:길병원용 Function : 수정요!! */
''            strPtInfo = OraDS.GetOcsPtInfo(mvarPtId, mvarDeptCd, mvarOrdDoct)
''            mvarPtNm = medGetP(strPtInfo, 1, vbTab)
''            mvarDOB = medGetP(strPtInfo, 2, vbTab)
''            mvarSex = medGetP(strPtInfo, 3, vbTab)
''            mvarDeptNm = medGetP(strPtInfo, 4, vbTab)
''            mvarOrdDoctNm = medGetP(strPtInfo, 5, vbTab)
''            mvarAge = medFindAge(mvarDOB, "Y", mvarAccDt)
''        Else
''            '/현재의 DB에서 기본정보를 가져오는 경우(검사실)
''            mvarPtNm = "" & DS.Fields("PTNM").Value
''            mvarSex = "" & DS.Fields("SEX").Value
''            mvarDOB = "" & DS.Fields("DOB").Value
''            mvarDeptNm = "" & DS.Fields("DEPTNM").Value
''            mvarOrdDoctNm = "" & DS.Fields("DOCTNM").Value
''            mvarAge = medFindAge("" & DS.Fields("DOB"), "Y", mvarAccDt)
''        End If
''        mvarSexAge = mvarSex & "/" & mvarAge
''        '
''        mvarMFootNote = ""
''        '
''        If mvarBedInDt <> "" Then
''            mvarLocation = "" & DS.Fields("WARDID").Value & "-" & _
''                           "" & DS.Fields("ROOMID").Value & "-" & _
''                           "" & DS.Fields("BEDID").Value
''        Else
''            mvarLocation = DeptNm
''        End If
''        If mvarRmkCd <> "" Then
''            mvarRmkNm = OraDS.GetTempleteText1(LC4_Remark, mvarRmkCd)
''        Else
''            mvarRmkNm = ""
''        End If
''        mvarReqTotCnt = "" & DS.Fields("REQTOTCNT").Value
''        mvarReqInputCnt = "" & DS.Fields("REQINPUTCNT").Value
''        mvarFootNoteFg = "" & DS.Fields("FOOTNOTEFG").Value
''        '
''        PrgSet 60
''        '
''        SELECT Case mvarPtType
''            Case RESULT_BY_DEFAULT:                 '/* 환자별 일반정보 */
''                mvarTestCount = 0
''                SetResultInit
''            Case RESULT_BY_ACCESSION:               '/* 접수번호별 일반정보+결과정보 */
''                OraDS.GetResult RESULT_BY_ACCESSION, RS, _
''                                mvarWorkArea, mvarAccDt, mvarAccSeq
''                PrgSet 80
''                SetResultCol RESULT_BY_ACCESSION, RS, TS, pOcsPtCheck
''                PrgSet 100
''            Case RESULT_BY_WORKSHEET:               '/* 워크쉬트별 일반정보+결과정보 */
''                OraDS.GetResult RESULT_BY_WORKSHEET, RS, _
''                                mvarWorkArea, mvarAccDt, mvarAccSeq, pSetCode
''                PrgSet 80
''                SetResultCol RESULT_BY_WORKSHEET, RS, TS, pOcsPtCheck
''                PrgSet 100
''            Case RESULT_BY_EQUIPMENT:               '/* 장비별 일반정보+결과정보 */
''                OraDS.GetResult RESULT_BY_EQUIPMENT, RS, _
''                                mvarWorkArea, mvarAccDt, mvarAccSeq, pSetCode
''                PrgSet 80
''                SetResultCol RESULT_BY_EQUIPMENT, RS, TS, pOcsPtCheck
''                PrgSet 100
''            Case RESULT_BY_MODIFY:                  '/* 결과수정 일반정보+결과정보 */
''                OraDS.GetResult RESULT_BY_MODIFY, RS, _
''                                mvarWorkArea, mvarAccDt, mvarAccSeq
''                PrgSet 80
''                SetResultCol RESULT_BY_MODIFY, RS, TS, pOcsPtCheck
''                PrgSet 100
''            Case RESULT_BY_DIFFCOUNT:               '/* 접수번호별 일반정보+결과정보 */
''                OraDS.GetResult RESULT_BY_DIFFCOUNT, RS, _
''                                mvarWorkArea, mvarAccDt, mvarAccSeq
''                PrgSet 80
''                SetResultCol RESULT_BY_DIFFCOUNT, RS, TS, pOcsPtCheck
''                PrgSet 100
''        End SELECT
''        '
''        mvarRecordCount = DS.RecordCount
''        MoveFirst
''    Else
''        'Case :RESULT_BY_ITEM( 아이템별 일반정보+결과정보 )
''        If pItemMode = "WorkSheet" Then
''            '워크쉬트별 아이템 결과등록
''            OraDS.GetItemResult "WorkSheet", RS, medGetP(pSetCode, 1, vbTab), _
''                                medGetP(pSetCode, 2, vbTab) & vbTab & medGetP(pSetCode, 3, vbTab) & vbTab & _
''                                medGetP(pSetCode, 4, vbTab) & vbTab & medGetP(pSetCode, 5, vbTab), pBuildCd
''            PrgSet 40
''        Else
''            '접수번호별 아이템 결과등록
''            OraDS.GetItemResult "AccNo", RS, medGetP(pSetCode, 1, vbTab), _
''                                DBAccNo(medGetP(pSetCode, 2, vbTab)) & vbTab & _
''                                DBAccNo(medGetP(pSetCode, 3, vbTab)), pBuildCd
''            PrgSet 40
''        End If
''        SetResultCol RESULT_BY_ITEM, RS, TS, pOcsPtCheck
''        PrgSet 100
''    End If
''    OraErr.Reset
''    Exit Function
''    '
''ErrorLoadTable:
''    OraErr.Reset
''    'Resume Next
''    If OraDS.LastServerErr <> 0 Then
''        glngErrorNo = OraDS.LastServerErr
''        gstrErrorMsg = OraDS.LastServerErrText
''    Else
''        If IsObject(Err) Then
''            glngErrorNo = Err.Number
''            gstrErrorMsg = Err.Description
''        End If
''    End If
''    PrgSet 250
'    '
'End Function
'
'Private Sub CopyDataToProperty(ByVal index As Long)
'
'    'With Result.Item(index)
''    'Date형식과 Enum은 초기치(vbnullstring or 0)를 주면 않된다.
''    If (index < 1) Or (index > Me.TestCount) Then
''        If index < 1 Then
''            m_INDEX = 0
''        Else
''            m_INDEX = Me.TestCount + 1
''        End If
''        '
''        SetResultInit
''        mvarKey = vbNullString                          'Collection Class Key Setting
''        mvarTestCount = Me.Result.Count
''        m_INDEX = index
''    Else
''        With Me.Result.Item(index)
''            m_DataFg = .DataFg
''            m_WorkArea = .WorkArea
''            m_AccDt = .AccDt
''            m_AccSeq = .AccSeq
''            m_TestCd = .TestCd
''            m_OrdDt = .orddt
''            m_OrdNo = .OrdNo
''            m_OrdSeq = .OrdSeq
''            m_RstVal = .rstval
''            m_RstCd = .RstCd
''            m_RstUnit = .RstUnit
''            m_HLDiv = .HLDiv
''            m_DPDiv = .DPDiv
''            m_SpcCd = .SpcCd
''            m_PtInfo = .PtInfo
''            m_LastRst = .LastRst
''            m_LastVfyDt = .LastVfyDt
''            m_LastVfyTm = .LastVfyTm
''            m_LastVfyId = .LastVfyId
''            m_VfyDt = .VfyDt
''            m_VfyTm = .VfyTm
''            m_VfyId = .VfyId
''            m_MfyFg = .MfyFg
''            m_GrpFg = .GrpFg
''            m_AutoFg = .autofg
''            m_TxtFg = .TxtFg
''            m_ExcFg = .ExcFg
''            m_RstType = .rsttype
''            m_RstDiv = .RstDiv
''            m_TextRst = .TextRst
''            m_SuppText = .SuppText
''            m_ColDt = .ColDt
''            m_TxtType = .TxtType
''            m_TestNm = .TestNm
''            m_AvalVal = .AvalVal
''            m_PanicFg = .PanicFg
''            m_PanicFrVal = .PanicFrVal
''            m_PanicToVal = .PanicToVal
''            m_DeltaFg = .DeltaFg
''            m_DeltaVal1 = .DeltaVal1
''            m_DeltaVal2 = .DeltaVal2
''            m_RefValFrom = .RefValFrom
''            m_RefValTo = .RefValTo
''            m_RefCd = .RefCd
''            m_DetailFg = .DetailFg
''            m_EqpCd = .EqpCd
''            m_OTmpCd = .OTmpCd
''            m_EqpNm = .EqpNm
''            m_TransDtTm = .TransDtTm
''            m_SpcYy = .SpcYy
''            m_SpcNo = .SpcNo
''            m_MRstVal = .MRstVal
''            m_MRstCd = .MRstCd
''            m_MHLDiv = .MHLDiv
''            m_MDPDiv = .MDPDiv
''            m_MfyRsn = .MfyRsn
''            m_MRstType = .MRstType
''            m_FootNote = .FootNote
''            m_ReqAltCnt = .ReqAltCnt
''            mvarKey = .Key                                   'Collection Class Key Setting
''            mvarTestCount = Me.Result.Count
''            m_INDEX = index
''        End With
''    End If
''    '
''    With Me.Result
''        If .Count > 0 Then
''            SELECT Case m_INDEX
''                Case Is < 1
''                    mvarBOF = True
''                    mvarEOF = False
''                Case Is > .Count
''                    mvarEOF = True
''                    mvarBOF = False
''                Case Else
''                    mvarBOF = False
''                    mvarEOF = False
''            End SELECT
''        Else
''            mvarEOF = True
''            mvarBOF = True
''        End If
''    End With
''    '
'End Sub
'
'Private Sub SetResultInit()
'      m_DataFg = "0": m_WorkArea = ""
'      m_AccDt = "": m_AccSeq = ""
'      m_TestCd = "": m_OrdDt = ""
'      m_OrdNo = "": m_OrdSeq = ""
'      m_RstVal = "": m_RstCd = ""
'      m_RstUnit = "": m_HLDiv = ""
'      m_DPDiv = "": m_SpcCd = ""
'      m_PtInfo = "": m_LastRst = ""
'      m_LastVfyDt = "": m_LastVfyTm = ""
'      m_LastVfyId = "": m_VfyDt = ""
'      m_SpcYy = "": m_SpcNo = ""
'      m_VfyTm = "": m_VfyId = ""
'      m_MfyFg = "": m_GrpFg = ""
'      m_AutoFg = "0": m_TxtFg = "0"
'      m_ExcFg = "0": m_RstType = ""
'      m_RstDiv = "": m_TextRst = ""
'      m_SuppText = "": m_ColDt = ""
'      m_TxtType = "0": m_TestNm = ""
'      m_AvalVal = "": m_PanicFg = ""
'      m_PanicFrVal = 0: m_PanicToVal = 0
'      m_DeltaFg = "": m_DetailFg = ""
'      m_DeltaVal1 = 0: m_DeltaVal2 = 0
'      m_RefValFrom = 0: m_RefValTo = 0
'      m_RefCd = ""
'      m_EqpCd = "": m_OTmpCd = ""
'      m_EqpNm = "": m_TransDtTm = ""
'      m_MRstVal = "": m_MRstCd = ""
'      m_MHLDiv = "": m_MDPDiv = ""
'      m_MfyRsn = "": m_MRstType = ""
'      m_FootNote = "": m_ReqAltCnt = ""
'      m_WorkSeq = ""
'End Sub
'
'Private Sub SetResultCol(ByVal QueryType As Integer, _
'                         ByRef RS As RECORDSET, ByRef TS As RECORDSET, _
'                         ByVal OcsChk As String)
''    Dim strSex As String
''    Dim blnFootNote As Boolean
''    Dim strTest As String
''    Dim strEqp As String
''    Dim m_Key As String
''    Dim strPtInfo As String
''    Dim strAge As String
''    '
''    mvarTestCount = RS.RecordCount
''    If RS.RecordCount = 0 Then
''       SetResultInit
''       Exit Sub
''    End If
''    '
''    '/* 아이템별 결과등록이 아닌 로직에서 한번만 FOOTNOTE내역을 읽기위해 */
''    blnFootNote = False
''    '
''    Do Until RS.EOF
''        m_DataFg = "0"
''        '탁과장님---------------------------------------
''        '장비별결과등록에서 장비전송내역을 읽어오기위함.
''        m_SpcYy = "" & RS.Fields("SPCYY").Value
''        m_SpcNo = "" & RS.Fields("SPCNO").Value
''        '-----------------------------------------------
''        m_WorkArea = "" & RS.Fields("WORKAREA").Value
''        m_AccDt = "" & RS.Fields("ACCDT").Value
''        m_AccSeq = "" & RS.Fields("ACCSEQ").Value
''        m_TestCd = "" & RS.Fields("TESTCD").Value
''        m_OrdDt = "" & RS.Fields("ORDDT").Value
''        m_OrdNo = "" & RS.Fields("ORDNO").Value
''        m_OrdSeq = "" & RS.Fields("ORDSEQ").Value
''        m_RstVal = "" & RS.Fields("RSTVAL").Value
''        m_RstCd = "" & RS.Fields("RSTCD").Value
''        m_RstUnit = "" & RS.Fields("RSTUNIT").Value
''        m_RstDiv = "" & RS.Fields("RSTDIV").Value
''        m_HLDiv = "" & RS.Fields("HLDIV").Value
''        m_DPDiv = "" & RS.Fields("DPDIV").Value
''        m_SpcCd = "" & RS.Fields("SPCCD").Value
''        m_LastRst = "" & RS.Fields("LASTRST").Value
''        m_LastVfyDt = "" & RS.Fields("LASTVFYDT").Value
''        m_LastVfyTm = "" & RS.Fields("LASTVFYTM").Value
''        m_LastVfyId = "" & RS.Fields("LASTVFYID").Value
''        m_VfyDt = "" & RS.Fields("VFYDT").Value
''        m_VfyTm = "" & RS.Fields("VFYTM").Value
''        m_VfyId = "" & RS.Fields("VFYID").Value
''        m_MfyFg = "" & RS.Fields("MFYFG").Value
''        m_GrpFg = "" & RS.Fields("GRPFG").Value
''        m_AutoFg = "" & RS.Fields("AUTOFG").Value
''        m_TxtFg = "" & RS.Fields("TXTFG").Value
''        m_DetailFg = "" & RS.Fields("DETAILFG").Value
''        m_ColDt = "" & RS.Fields("COLDT").Value
''        m_RstType = "" & RS.Fields("RSTTYPE").Value
''        m_EqpCd = "" & RS.Fields("EQPCD").Value
''
''        '
''        If mvarPtType = RESULT_BY_MODIFY Then
''            '결과수정시 원래 장비코드를 클래스에 보관.
''            m_OTmpCd = m_EqpCd
''        ElseIf mvarPtType = RESULT_BY_ITEM Then
''            '아이템별 결과등록시 Remark코드르 클래스에 보관.
''            m_OTmpCd = "" & RS.Fields("RMKCD").Value
''            '***** WorkSeq : Item별 결과등록에 보여주기 위해서...
''            m_WorkSeq = "" & RS.Fields("WORKSEQ").Value
''
''        Else
''            m_OTmpCd = ""
''        End If
''        '
''        If mvarPtType = RESULT_BY_ITEM Then
''            If OcsChk = "OCS" Then
''                '병원쪽에서 기본정보를 가져오는 경우
''                strPtInfo = OraDS.GetOcsPtInfo("" & RS.Fields("PTID").Value, "", "")
''                strPtInfo = strPtInfo
''                m_PtInfo = m_WorkArea & "-" & Mid(m_AccDt, 3, 6) & "-" & m_AccSeq & vbTab & _
''                            "" & RS.Fields("PTID").Value & vbTab & _
''                            medGetP(strPtInfo, 1, vbTab) & vbTab & _
''                            medGetP(strPtInfo, 3, vbTab) & " / " & medFindAge(medGetP(strPtInfo, 2, vbTab), "Y", m_AccDt) & vbTab & _
''                            "" & RS.Fields("AGEDAY").Value & vbTab & _
''                            medGetP(strPtInfo, 3, vbTab)
''            Else
''                '자체 SYSTEM에서 기본정보를 가져오는 경우
''                m_PtInfo = m_WorkArea & "-" & Mid(m_AccDt, 3, 6) & "-" & m_AccSeq & vbTab & _
''                            "" & RS.Fields("PTID").Value & vbTab & _
''                            "" & RS.Fields("PTNM").Value & vbTab & _
''                            "" & RS.Fields("SEX").Value & " / " & medFindAge("" & RS.Fields("DOB").Value, "Y", m_AccDt) & vbTab & _
''                            "" & RS.Fields("AGEDAY").Value & vbTab & _
''                            "" & RS.Fields("SEX").Value
''            End If
''            m_ReqAltCnt = "" & RS.Fields("REQTOTCNT").Value & vbTab & _
''                            "" & RS.Fields("REQINPUTCNT").Value & vbTab
''               '"" & rs.Fields( "ALTTOTCNT").value & vbTab & _
''               '"" & rs.Fields( "ALTINPUTCNT").value
''        Else
''            m_ReqAltCnt = ""
''            m_PtInfo = ""
''            m_FootNote = ""
''        End If
''        '
''        If mvarPtType <> RESULT_BY_MODIFY Then
''            m_MRstVal = ""
''            m_MRstCd = ""
''            m_MHLDiv = ""
''            m_MDPDiv = ""
''            m_MRstType = ""
''        Else
''            m_MRstVal = "" & RS.Fields("RSTVAL").Value
''            m_MRstCd = "" & RS.Fields("RSTCD").Value
''            m_MHLDiv = "" & RS.Fields("HLDIV").Value
''            m_MDPDiv = "" & RS.Fields("DPDIV").Value
''            m_MRstType = "" & RS.Fields("RSTTYPE").Value
''        End If
''        '
''        '/* Foot Note내역 가져오기. */
''        If mvarPtType <> RESULT_BY_ITEM Then
''            If blnFootNote = False AND Me.FootNoteFg = "1" Then
''                mvarFootNote = OraDS.GetRstTxt(m_WorkArea, m_AccDt, m_AccSeq)
''                blnFootNote = True
''            End If
''        Else
''            m_FootNote = OraDS.GetRstTxt(m_WorkArea, m_AccDt, m_AccSeq)
''        End If
''        '
''        '/* Text결과 및 Supplemental Report 가져오기. */
''        If m_TxtFg = "1" Then
''            m_TextRst = OraDS.GetRstTxt(m_WorkArea, m_AccDt, m_AccSeq, m_TestCd)
''            m_SuppText = ""
''        ElseIf m_TxtFg = "2" Then
''            m_TextRst = OraDS.GetRstTxt(m_WorkArea, m_AccDt, m_AccSeq, m_TestCd)
''            If mvarPtType = RESULT_BY_MODIFY Then
''                m_TextRst = m_TextRst & _
''                            OraDS.GetSuppRstTxt(m_WorkArea, m_AccDt, m_AccSeq, m_TestCd)
''                m_SuppText = ""
''            Else
''                m_SuppText = ""
''            End If
''        Else
''            m_TextRst = ""
''            m_SuppText = ""
''        End If
''        '
''        '/*아이템 마스터에서 아이템명,결과유형,텍스트결과유형을 가져와
''        '  결과내역의 기본 데이터를 설정한다. */
''        If mvarPtType <> RESULT_BY_ITEM Then
'''            strTest = OraDS.GetTestNm(m_TestCd, m_ColDt, "TxtType")
''            m_TestNm = "" & RS.Fields("TESTNM").Value    'medGetP(strTest, 1, vbTab)
''            m_TxtType = "" & RS.Fields("TXTTYPE").Value   'medGetP(strTest, 2, vbTab)
''            If m_TxtType = "" Then
''                m_TxtType = "0"
''            End If
''            If mvarPtType <> RESULT_BY_MODIFY Then
''                If (m_RstVal <> "0.0000" AND m_RstVal <> "") Or m_RstCd <> "" Or m_TextRst <> "" Then
''                    '기존에 결과를 넣고 보류한 경우.
''                    m_ExcFg = "1"         '기존에 이미 결과를 넣고 보류된 상태
''                Else
''                    '결과가 있지 않은 경우(보류가 않된 상태)
''                    m_ExcFg = "0"         '기존에 보류가 않된 상태
''                End If
''            End If
''        Else
''            '아이템별결과 등록에서는 테스트명을 셋팅하지 않는다.
''            m_TestNm = ""
''            m_TxtType = "" & RS.Fields("TXTTYPE").Value
''            If m_TxtType = "" Then
''                m_TxtType = "0"
''            End If
''        End If
''        '
''        '*/아이템별 참고치,유효숫자,패닉델타정보 등을 가져온다. */
'''        OraDS.GetPanicDelta TS, m_TestCd, m_SpcCd, m_ColDt
'''        If TS.RecordCount = 0 Then
'''            '/* 유효숫자 및 패닉델타정보를 읽지 못한 경우 유효숫자가 NULL이 된다. */
'''            '/* 유효숫자가 "9": 유효숫자 관리를 하지않고 숫자형의 모든 데이터를 허용한다. */
'''            m_AvalVal = "9"
'''            m_PanicFg = ""
'''            m_PanicFrVal = 0
'''            m_PanicToVal = 0
'''            m_DeltaFg = ""
'''            m_DeltaVal1 = 0
'''            m_DeltaVal2 = 0
'''        Else
''            ' *유효숫자 NULL-모든숫자 허용,'0'-정수형,>'0'-유효숫자 관리
''            m_AvalVal = "" & RS.Fields("AVALVAL").Value
''            '
''            m_PanicFg = "" & RS.Fields("PANICFG").Value
''            If m_PanicFg <> "1" Then
''                m_PanicFg = ""
''                m_PanicFrVal = 0
''                m_PanicToVal = 0
''            Else
''                m_PanicFrVal = "" & RS.Fields("PANICFRVAL").Value
''                m_PanicToVal = "" & RS.Fields("PANICTOVAL").Value
''            End If
''            m_DeltaFg = "" & RS.Fields("DELTAFG").Value
''            If m_DeltaFg <> "1" Then
''                m_DeltaFg = ""
''                m_DeltaVal1 = 0
''                m_DeltaVal2 = 0
''            Else
''                m_DeltaVal1 = "" & RS.Fields("DELTAVAL").Value
''                m_DeltaVal2 = "" & RS.Fields("DELTAVAL2").Value
''            End If
'''        End If
''
''        'medFindAge(mvarDOB, "D", m_ColDt)
''        '원래 성('M','F','A','-')으로 참조치를 구한다.
''        If mvarPtType = RESULT_BY_ITEM Then
''            'm_PtInfo = ACCNO + PTID + PTNM + SEX/AGE + AGEDAY  + SEX : VbTab
''            OraDS.GetRefVal TS, m_TestCd, m_SpcCd, m_ColDt, "B", medGetP(m_PtInfo, 5, vbTab)
''        Else
''            OraDS.GetRefVal TS, m_TestCd, m_SpcCd, m_ColDt, "B", mvarAgeDay
''        End If
''        If TS.RecordCount = 0 Then
''            '원래 성('M','F','A','-')으로 조회해서 없으면 'B'(BOTH)형의 성의 가장
''            '참조치를 가져온다.
''            If mvarPtType = RESULT_BY_ITEM Then
''                OraDS.GetRefVal TS, m_TestCd, m_SpcCd, m_ColDt, medGetP(m_PtInfo, 6, vbTab), medGetP(m_PtInfo, 5, vbTab)
''            Else
''                OraDS.GetRefVal TS, m_TestCd, m_SpcCd, m_ColDt, mvarSex, mvarAgeDay
''            End If
''            If TS.RecordCount = 0 Then
''                '/* 주의 : 참조치가 없으면 RefValFrom,RefValTo가 모두 0이 된다. */
''                m_RefValFrom = 0
''                m_RefValTo = 0
''                m_RefCd = ""
''            End If
''        End If
''        '
''        If TS.RecordCount > 0 Then
''            strSex = "" & TS.Fields("APPLYSEX").Value
''            m_RefValFrom = "" & TS.Fields("REFVALFROM").Value
''            m_RefValTo = "" & TS.Fields("REFVALTO").Value
''            m_RefCd = "" & TS.Fields("REFCD").Value
''        End If
''        '
''        '/* 장비정보 가져오기 */
''        If m_EqpCd <> "" Then
''            If mvarPtType = RESULT_BY_ITEM Then
'''                strEqp = OraDS.GetEqpTrans(m_EqpCd, m_SpcYy, m_SpcNo, m_AutoFg)
''                m_EqpNm = medShift(strEqp, vbTab)
'''                m_TransDtTm = medShift(strEqp, vbTab)
'''                m_TransDtTm = m_TransDtTm & " " & Format(strEqp, "AMPM hh:mm:ss")
''            Else
''                m_EqpNm = "" & RS.Fields("EQPNM").Value         'medShift(strEqp, vbTab)
'''                m_TransDtTm = "" & rs.Fields("TRANSDT").Value   'medShift(strEqp, vbTab)
'''                strEqp = Mid("" & rs.Fields("TRANSTM").Value, 1, 2) & ":" & _
'''                         Mid("" & rs.Fields("TRANSTM").Value, 3, 2) & ":" & _
'''                         Mid("" & rs.Fields("TRANSTM").Value, 5, 2)
'''                m_TransDtTm = m_TransDtTm & " " & Format(strEqp, "AMPM hh:mm:ss")
''            End If
''        Else
''            m_EqpNm = ""
''            m_TransDtTm = ""
''        End If
''        '*
''        RS.MoveNext
''        '
''        If mvarPtType <> RESULT_BY_ITEM Then
''            'Collection Class Key Setting
''            mvarKey = m_TestCd
''            m_Key = m_TestCd
''        Else
''            'Collection Class Key Setting
''            m_Key = m_WorkSeq & medGetP(m_PtInfo, 1, vbTab) & m_TestCd
''        End If
''        mvarResult.Add mvarResult, m_Key, m_WorkArea, m_AccDt, m_AccSeq, m_TestCd, _
''                        m_RstVal, m_RstCd, m_RstUnit, m_HLDiv, m_DPDiv, m_SpcCd, _
''                        m_PtInfo, m_LastRst, m_LastVfyDt, m_LastVfyTm, m_LastVfyId, _
''                        m_VfyDt, m_VfyTm, m_VfyId, m_MfyFg, m_GrpFg, m_AutoFg, m_TxtFg, m_ExcFg, _
''                        m_RstType, m_RstDiv, m_TextRst, m_SuppText, _
''                        m_ColDt, m_TxtType, m_TestNm, m_AvalVal, m_PanicFg, m_PanicFrVal, _
''                        m_PanicToVal, m_DeltaFg, m_DeltaVal1, m_DeltaVal2, m_RefValFrom, m_RefValTo, _
''                        m_RefCd, m_OrdDt, m_OrdNo, m_OrdSeq, m_DetailFg, m_DataFg, _
''                        m_EqpCd, m_OTmpCd, m_EqpNm, m_TransDtTm, m_MRstVal, m_MRstCd, _
''                        m_MHLDiv, m_MDPDiv, m_MfyRsn, m_MRstType, m_FootNote, m_ReqAltCnt, mvarTestCount, m_WorkSeq
'
''    Loop
'    '
'    'Exit Sub
'End Sub
'
'
''/* Spread AND Progresss Bar Method
'Public Sub GetResultSpread(ByRef objResult As clsDictionary, _
'                           ByRef objRST As clsDictionary)
''    Dim objRST As clsResult
''    Dim strTmp As String
''    Dim aryFld() As String
''    Dim aryPtInfo() As String
''    Dim aryRec() As String
''    Dim ii As Long
''    Dim jj As Long
''
''    Set ssRst = objSS
''    ssRst.MaxRows = 0
''    If Me.TestCount = 0 Then
''       ssRst.Visible = True
''       PrgSet 250
''       Exit Sub
''    End If
''    PrgSet 120
''    '
''
''
''    '"labno,testcd", "spccd,rstval,rstcd,rsttype,hldiv,dpdiv,hold"
''
''    'SS_COL_SCREEN_MAX=18로 통일
''
''    ReDim aryRec(Me.TestCount - 1)
''    SELECT Case intSpType
''        Case RESULT_BY_DEFAULT:
''            SS_COL_TESTNM = 1:         SS_COL_RESULT = 2
''            SS_COL_EC = 3:             SS_COL_LASTRST = 4
''            SS_COL_RSTUNIT = 5:        SS_COL_HLDIV = 6
''            SS_COL_DPDIV = 7:          SS_COL_TXT = 8
''            SS_COL_EQPNM = 9:          SS_COL_GRPFG = 10
''            SS_COL_REFVAL = 11
''            SS_COL_SCREEN_MAX = 18
'''            SS_COL_SCREEN_MAX = SS_COL_GRPFG + 1
''        Case RESULT_BY_ACCESSION:
''            SS_COL_TESTNM = 1:         SS_COL_RESULT = 2
''            SS_COL_EC = 4:             SS_COL_LASTRST = 5
''            SS_COL_RSTUNIT = 6:        SS_COL_HLDIV = 7
''            SS_COL_DPDIV = 8:          SS_COL_REFVAL = 9
''            SS_COL_TXT = 10:           SS_COL_EQPNM = 11
''            SS_COL_GRPFG = 12
''            SS_COL_SCREEN_MAX = 18
'''            SS_COL_SCREEN_MAX = SS_COL_GRPFG
''        Case RESULT_BY_WORKSHEET:
''            SS_COL_SEQ = 1
''            SS_COL_ACCNO = 2:          SS_COL_PTID = 3
''            SS_COL_PTNM = 4:           SS_COL_SEXAGE = 5
''            SS_COL_TESTNM = 6:         SS_COL_RESULT = 7
''            SS_COL_EC = 8:             SS_COL_LASTRST = 9
''            SS_COL_RSTUNIT = 10:       SS_COL_HLDIV = 11
''            SS_COL_DPDIV = 12:         SS_COL_REFVAL = 13
''            SS_COL_TXT = 14:           SS_COL_EQPNM = 15
''            SS_COL_GRPFG = 16
''            SS_COL_SCREEN_MAX = 18
'''            SS_COL_SCREEN_MAX = SS_COL_GRPFG
''    End SELECT
''    '
''    If (intSpType = RESULT_BY_DEFAULT) Or _
''       (intSpType = RESULT_BY_ACCESSION) Or _
''       (intSpType = RESULT_BY_WORKSHEET) Then
''
''        For ii = 1 To mvarTestCount
''            ReDim aryFld(SS_COL_SCREEN_MAX - 1)
''            With Me.Result.Item(ii)
''                If intSpType <> 3 Then
''                    aryFld(SS_COL_TESTNM - 1) = .TestNm    '& "-" & .TestCd
''                End If
''                If .RstDiv <> "*" Then
''                    If .rsttype = "N" Then
''                        aryFld(SS_COL_RESULT - 1) = .RstCd
''                    Else
''                        If (.RstCd Like "*:*") AND .rsttype <> "F" Then
''                            aryFld(SS_COL_RESULT - 1) = .rstval
''                        Else
''                            'aryFld(SS_COL_RESULT - 1) = .RstCd
''
''                            aryFld(SS_COL_RESULT - 1) = GetRstcd(.TestCd, .RstCd)
''                        End If
''                    End If
''                    '18번째 Col에 결과 코드값 (실제 저장할 값을 가지고 있는다)
''                    '추가부분(결과 입력란을 코드가 아닌 코드값으로 저장하기 위해서 수정함
''                    '.RSTCD를 가지고 결과코드를 무조건 한번은 뒤져야 한다.
''                    '2002/10/09추가 by 김정규
''
''                    aryFld(SS_COL_SCREEN_MAX - 1) = .RstCd
''
''
''                    aryFld(SS_COL_EC - 1) = .ExcFg
''                    If .LastRst <> "" Then
''                        If IsNumeric(.LastRst) Then
''                            aryFld(SS_COL_LASTRST - 1) = AvalNum(.LastRst, .AvalVal)
''                        Else
''                            '결과값가져오기
''                            'aryFld(SS_COL_LASTRST - 1) = .LastRst
''
''                            aryFld(SS_COL_LASTRST - 1) = GetRstcd(.TestCd, .LastRst)
''                        End If
''                    Else
''                        aryFld(SS_COL_LASTRST - 1) = ""
''                    End If
''                    aryFld(SS_COL_RSTUNIT - 1) = .RstUnit
''                    aryFld(SS_COL_HLDIV - 1) = .HLDiv
''                    aryFld(SS_COL_DPDIV - 1) = .DPDiv
''                    aryFld(SS_COL_TXT - 1) = ""
''                    aryFld(SS_COL_GRPFG - 1) = ""
''                    aryFld(SS_COL_EQPNM - 1) = .EqpNm
''                    aryRec(ii - 1) = Join(aryFld, vbTab)
''                    'If intSpType = 2 Or intSpType = 3 Then
''                        If .RefValFrom = 0 AND .RefValTo = 0 Then
''                            aryFld(SS_COL_REFVAL - 1) = " - "
''                        Else
''                            'aryFld(SS_COL_REFVAL - 1) = CStr(.RefValFrom) & " ~ " & CStr(.RefValTo)
''                            aryFld(SS_COL_REFVAL - 1) = AvalNum(CStr(.RefValFrom), .AvalVal) & _
''                                                        " ~ " & AvalNum(CStr(.RefValTo), .AvalVal)
''                        End If
''                    'End If
''                    If intSpType = 3 Then
''                         '아이템별 결과등록의 Spread 셋팅.
''                         'm_PtInfo = ACCNO + PTID + PTNM + SEX/AGE + AGEDAY : VbTab
''                         aryPtInfo = Split(.PtInfo, vbTab)
''                         aryFld(SS_COL_SEQ - 1) = .WorkSeq
''                         aryFld(SS_COL_ACCNO - 1) = aryPtInfo(0)
''                         aryFld(SS_COL_PTID - 1) = aryPtInfo(1)
''                         aryFld(SS_COL_PTNM - 1) = aryPtInfo(2)
''                         aryFld(SS_COL_SEXAGE - 1) = aryPtInfo(3)
''                    End If
''                    aryRec(ii - 1) = Join(aryFld, vbTab)
''                Else
''                    If PROJECT_HOSCD = "04" Then aryFld(SS_COL_EC - 1) = .ExcFg
''
''                    aryRec(ii - 1) = Join(aryFld, vbTab)
''                End If
''            End With
''        Next
''        strTmp = Join(aryRec, vbNewLine)
''    ElseIf intSpType = 3 Then
''    End If
''    '
''    PrgSet 140
''    With ssRst
''        .MaxRows = Me.TestCount
''        .Col = 1:  .Col2 = UBound(aryFld) + 1
''        .Row = 1:  .Row2 = .MaxRows
''        .Clip = strTmp
''        PrgSet 190
''        SpreadLock 1, (SS_COL_RESULT - 1)
''        SpreadLock SS_COL_LASTRST, SS_COL_SCREEN_MAX
''    End With
''    '
''    For ii = 1 To ssRst.MaxRows
''        ssRst.Row = ii
''        With Me.Result.Item(ii)
'''            If intSpType = 2 Then
''               ssRst.Col = SS_COL_REFVAL
''               ssRst.ForeColor = &H136604
'''            End If
''            '
''            ssRst.Col = SS_COL_RESULT
''            'Interface결과가 오류 데이타('Err')이면 빨간색으로...
''            If Trim(ssRst.Value) = CS_EqpError Then ssRst.ForeColor = vbRed
''
''            If Me.Result.Item(ii).RstDiv <> "*" Then
''                If .TxtType = "0" Then
''                    .TextRst = ""
''                    ssRst.Col = SS_COL_TXT
''                    ssRst.Value = ""
''                    ssRst.CellType = 5 'CellTypeStaticText
''                Else
''                    If .TxtType = "1" Then
''                        ssRst.Col = SS_COL_RESULT
''                        ssRst.Value = ""
''                        ssRst.CellType = 5 'CellTypeStaticText
''                    End If
''                    ssRst.Col = SS_COL_TXT
''                    ssRst.CellType = 9 ' CellTypePicture
''                    LoadPic "T", ii
''                End If
''                '----------------------------------------------------------------------
''                '  유효숫자대로 Cell Format을 설정 -- 입력 오류 가능성이 큼 --> 보류
''                '----------------------------------------------------------------------
''                '    If .AvalVal <> 9 AND .AvalVal <> "" Then
''                '        ssRst.Col = SS_COL_RESULT
''                '        ssRst.CellType = CellTypeFloat
''                '        ssRst.TypeFloatDecimalChar = Asc(".")
''                '        ssRst.TypeFloatDecimalPlaces = .AvalVal
''                '    End If
''                '----------------------------------------------------------------------
''                '----------------------------------------------------------------------
''                '  1999.07.18  수정  by 김미경
''                '  Detail일 경우 대표항목이 아니면 보류버튼 없음 .
''                '----------------------------------------------------------------------
''                If .DetailFg <> "" Then
''                    ssRst.Col = SS_COL_EC
''                    ssRst.Value = ""
''                    ssRst.CellType = 5 'CellTypeStaticText
''                End If
''                '----------------------------------------------------------------------
''
''                ssRst.Col = SS_COL_GRPFG
''                If .GrpFg = "0" Then
''                    ssRst.Value = ""
''                    ssRst.CellType = 5 'CellTypeStaticText
''                Else
''                    LoadPic "G", ii
''                End If
''                If .HLDiv <> "" Then
''                    LoadPic .HLDiv, ii
''                End If
''                If .DPDiv <> "" Then
''                    LoadPic .DPDiv, ii
''                End If
''                ssRst.Col = SS_COL_TESTNM
''                ssRst.ForeColor = DCM_MidBlue   'vbBlack
''                If .DetailFg <> "" Then
''                    ssRst.Value = "       " & ssRst.Value
''                End If
''                ssRst.Col = SS_COL_LASTRST
''                ssRst.ForeColor = &H8000&
''                ssRst.Col = SS_COL_RSTUNIT
''                ssRst.ForeColor = &H40&
''                If .EqpCd <> "" Then
''                    ssRst.Col = SS_COL_EQPNM
''                    If .autofg = "1" Then
''                        ssRst.ForeColor = &H808080
''                    Else
''                        ssRst.ForeColor = vbBlack
''                    End If
''                End If
''            Else
''                ssRst.Col = SS_COL_TESTNM
''                ssRst.ForeColor = vbBlack
''                For jj = 1 To SS_COL_SCREEN_MAX
''                    ssRst.Col = jj
''                    '----------------------------------------------------------------------
''                    '  1999.07.18  수정  by 김미경
''                    '  Detail일 경우 대표항목에만 보류버튼을 제공한다.
''                    '----------------------------------------------------------------------
''                    'ssRst.CellType = CellTypeStaticText
''                    If jj = 1 Then
''                        ssRst.BackColor = vbGrayText
''                        ssRst.ForeColor = vbWhite
''                    Else
''                        ssRst.BackColor = vbWhite
''                    End If
''                Next
''                ssRst.Row = ii: ssRst.Row2 = ii
''                ssRst.Col = SS_COL_RESULT: ssRst.Col2 = SS_COL_RESULT
''                ssRst.BlockMode = True
''                ssRst.Lock = True
''                ssRst.Protect = True
''                ssRst.BlockMode = False
''            End If
''        End With
''    Next ii
''    '
''    PrgSet 250
''    ssRst.Visible = True
'    '
'End Sub
'
'Public Sub GetDiffCntSpread(ByRef objSS As Object, _
'                            ByVal intSpType As Integer)
''    Dim objRST As clsResult
''    Dim strTmp As String
''    Dim aryFld() As String
''    Dim aryPtInfo() As String
''    Dim aryRec() As String
''    Dim ii As Long
''    Dim jj As Long
''
''    Set ssRst = objSS
''    ssRst.MaxRows = 0
''    If Me.TestCount = 0 Then
''       ssRst.Visible = True
''       PrgSet 250
''       Exit Sub
''    End If
''    PrgSet 120
''    '
''    ReDim aryRec(Me.TestCount - 1)
''    SELECT Case intSpType
''        Case RESULT_BY_DEFAULT:
''            SS_COL_TESTNM = 1:         SS_COL_RESULT = 2
''            SS_COL_EC = 3:             SS_COL_LASTRST = 4
''            SS_COL_RSTUNIT = 5:        SS_COL_HLDIV = 6
''            SS_COL_DPDIV = 7:          SS_COL_TXT = 8
''            SS_COL_EQPNM = 9:          SS_COL_GRPFG = 10
''            SS_COL_REFVAL = 11
''            SS_COL_SCREEN_MAX = SS_COL_GRPFG + 1
''        Case RESULT_BY_ACCESSION:
''            SS_COL_TESTNM = 1:         SS_COL_RESULT = 2
''            SS_COL_EC = 4:             SS_COL_LASTRST = 5
''            SS_COL_RSTUNIT = 6:        SS_COL_HLDIV = 7
''            SS_COL_DPDIV = 8:          SS_COL_REFVAL = 9
''            SS_COL_TXT = 10:           SS_COL_EQPNM = 11
''            SS_COL_GRPFG = 12
''            SS_COL_SCREEN_MAX = SS_COL_GRPFG
''        Case RESULT_BY_WORKSHEET:
''            SS_COL_SEQ = 1
''            SS_COL_ACCNO = 2:          SS_COL_PTID = 3
''            SS_COL_PTNM = 4:           SS_COL_SEXAGE = 5
''            SS_COL_TESTNM = 6:         SS_COL_RESULT = 7
''            SS_COL_EC = 8:             SS_COL_LASTRST = 9
''            SS_COL_RSTUNIT = 10:       SS_COL_HLDIV = 11
''            SS_COL_DPDIV = 12:         SS_COL_REFVAL = 13
''            SS_COL_TXT = 14:           SS_COL_EQPNM = 15
''            SS_COL_GRPFG = 16
''            SS_COL_SCREEN_MAX = SS_COL_GRPFG
''    End SELECT
''    '
''    If (intSpType = RESULT_BY_DEFAULT) Or _
''       (intSpType = RESULT_BY_ACCESSION) Or _
''       (intSpType = RESULT_BY_WORKSHEET) Then
''
''        For ii = 1 To mvarTestCount
''            ReDim aryFld(SS_COL_SCREEN_MAX - 1)
''            With Me.Result.Item(ii)
''                If intSpType <> 3 Then
''                    aryFld(SS_COL_TESTNM - 1) = .TestNm    '& "-" & .TestCd
''                End If
''                If .RstDiv <> "*" Then
''                    If .rsttype = "N" Then
''                        aryFld(SS_COL_RESULT - 1) = .RstCd
''                    Else
''                        If (.RstCd Like "*:*") AND .rsttype <> "F" Then
''                            aryFld(SS_COL_RESULT - 1) = .rstval
''                        Else
''                            aryFld(SS_COL_RESULT - 1) = .RstCd
''                        End If
''                    End If
''                    aryFld(SS_COL_EC - 1) = .ExcFg
''                    If .LastRst <> "" Then
''                        If IsNumeric(.LastRst) Then
''                            aryFld(SS_COL_LASTRST - 1) = AvalNum(.LastRst, .AvalVal)
''                        Else
''                            aryFld(SS_COL_LASTRST - 1) = .LastRst
''                        End If
''                    Else
''                        aryFld(SS_COL_LASTRST - 1) = ""
''                    End If
''                    aryFld(SS_COL_RSTUNIT - 1) = .RstUnit
''                    aryFld(SS_COL_HLDIV - 1) = .HLDiv
''                    aryFld(SS_COL_DPDIV - 1) = .DPDiv
''                    aryFld(SS_COL_TXT - 1) = ""
''                    aryFld(SS_COL_GRPFG - 1) = ""
''                    aryFld(SS_COL_EQPNM - 1) = .EqpNm
''                    aryRec(ii - 1) = Join(aryFld, vbTab)
''                    'If intSpType = 2 Or intSpType = 3 Then
''                        If .RefValFrom = 0 AND .RefValTo = 0 Then
''                            aryFld(SS_COL_REFVAL - 1) = " - "
''                        Else
''                            'aryFld(SS_COL_REFVAL - 1) = CStr(.RefValFrom) & " ~ " & CStr(.RefValTo)
''                            aryFld(SS_COL_REFVAL - 1) = AvalNum(CStr(.RefValFrom), .AvalVal) & _
''                                                        " ~ " & AvalNum(CStr(.RefValTo), .AvalVal)
''                        End If
''                    'End If
''                    If intSpType = 3 Then
''                         '아이템별 결과등록의 Spread 셋팅.
''                         'm_PtInfo = ACCNO + PTID + PTNM + SEX/AGE + AGEDAY : VbTab
''                         aryPtInfo = Split(.PtInfo, vbTab)
''                         aryFld(SS_COL_SEQ - 1) = .WorkSeq
''                         aryFld(SS_COL_ACCNO - 1) = aryPtInfo(0)
''                         aryFld(SS_COL_PTID - 1) = aryPtInfo(1)
''                         aryFld(SS_COL_PTNM - 1) = aryPtInfo(2)
''                         aryFld(SS_COL_SEXAGE - 1) = aryPtInfo(3)
''                    End If
''                    aryRec(ii - 1) = Join(aryFld, vbTab)
''                Else
''                    aryRec(ii - 1) = Join(aryFld, vbTab)
''                End If
''            End With
''        Next
''        strTmp = Join(aryRec, vbNewLine)
''    ElseIf intSpType = 3 Then
''    End If
''    '
''    PrgSet 140
''    With ssRst
''        .MaxRows = Me.TestCount \ 2 + 1
''        .Col = 1:  .Col2 = UBound(aryFld) + 1
''        .Row = 1:  .Row2 = .MaxRows
''        .Clip = strTmp
''        PrgSet 190
''        SpreadLock 1, (SS_COL_RESULT - 1)
''        SpreadLock SS_COL_LASTRST, SS_COL_SCREEN_MAX
''    End With
''    '
''    For ii = 1 To ssRst.MaxRows
''        ssRst.Row = ii
''        With Me.Result.Item(ii)
'''            If intSpType = 2 Then
''               ssRst.Col = SS_COL_REFVAL
''               ssRst.ForeColor = &H136604
'''            End If
''            '
''            ssRst.Col = SS_COL_RESULT
''            'Interface결과가 오류 데이타('Err')이면 빨간색으로...
''            If Trim(ssRst.Value) = CS_EqpError Then ssRst.ForeColor = vbRed
''
''            If Me.Result.Item(ii).RstDiv <> "*" Then
''                If .TxtType = "0" Then
''                    .TextRst = ""
''                    ssRst.Col = SS_COL_TXT
''                    ssRst.Value = ""
''                    ssRst.CellType = 5 'CellTypeStaticText
''                Else
''                    If .TxtType = "1" Then
''                        ssRst.Col = SS_COL_RESULT
''                        ssRst.Value = ""
''                        ssRst.CellType = 5 'CellTypeStaticText
''                    End If
''                    ssRst.Col = SS_COL_TXT
''                    ssRst.CellType = 9 ' CellTypePicture
''                    LoadPic "T", ii
''                End If
''                '----------------------------------------------------------------------
''                '  유효숫자대로 Cell Format을 설정 -- 입력 오류 가능성이 큼 --> 보류
''                '----------------------------------------------------------------------
''                '    If .AvalVal <> 9 AND .AvalVal <> "" Then
''                '        ssRst.Col = SS_COL_RESULT
''                '        ssRst.CellType = CellTypeFloat
''                '        ssRst.TypeFloatDecimalChar = Asc(".")
''                '        ssRst.TypeFloatDecimalPlaces = .AvalVal
''                '    End If
''                '----------------------------------------------------------------------
''                '----------------------------------------------------------------------
''                '  1999.07.18  수정  by 김미경
''                '  Detail일 경우 대표항목이 아니면 보류버튼 없음 .
''                '----------------------------------------------------------------------
''                If .DetailFg <> "" Then
''                    ssRst.Col = SS_COL_EC
''                    ssRst.Value = ""
''                    ssRst.CellType = 5 'CellTypeStaticText
''                End If
''                '----------------------------------------------------------------------
''
''                ssRst.Col = SS_COL_GRPFG
''                If .GrpFg = "0" Then
''                    ssRst.Value = ""
''                    ssRst.CellType = 9 'CellTypeStaticText
''                Else
''                    LoadPic "G", ii
''                End If
''                If .HLDiv <> "" Then
''                    LoadPic .HLDiv, ii
''                End If
''                If .DPDiv <> "" Then
''                    LoadPic .DPDiv, ii
''                End If
''                ssRst.Col = SS_COL_TESTNM
''                ssRst.ForeColor = DCM_MidBlue   'vbBlack
''                If .DetailFg <> "" Then
''                    ssRst.Value = "       " & ssRst.Value
''                End If
''                ssRst.Col = SS_COL_LASTRST
''                ssRst.ForeColor = &H8000&
''                ssRst.Col = SS_COL_RSTUNIT
''                ssRst.ForeColor = &H40&
''                If .EqpCd <> "" Then
''                    ssRst.Col = SS_COL_EQPNM
''                    If .autofg = "1" Then
''                        ssRst.ForeColor = &H808080
''                    Else
''                        ssRst.ForeColor = vbBlack
''                    End If
''                End If
''            Else
''                ssRst.Col = SS_COL_TESTNM
''                ssRst.ForeColor = vbBlack
''                For jj = 1 To SS_COL_SCREEN_MAX
''                    ssRst.Col = jj
''                    '----------------------------------------------------------------------
''                    '  1999.07.18  수정  by 김미경
''                    '  Detail일 경우 대표항목에만 보류버튼을 제공한다.
''                    '----------------------------------------------------------------------
''                    'ssRst.CellType = CellTypeStaticText
''                    If jj = 1 Then
''                        ssRst.BackColor = vbGrayText
''                        ssRst.ForeColor = vbWhite
''                    Else
''                        ssRst.BackColor = vbWhite
''                    End If
''                Next
''                ssRst.Row = ii: ssRst.Row2 = ii
''                ssRst.Col = SS_COL_RESULT: ssRst.Col2 = SS_COL_RESULT
''                ssRst.BlockMode = True
''                ssRst.Lock = True
''                ssRst.Protect = True
''                ssRst.BlockMode = False
''            End If
''        End With
''    Next ii
''    '
''    PrgSet 250
''    ssRst.Visible = True
''    '
'End Sub
'
'
'Private Function AvalNum(ByVal Val As String, _
'                         ByVal AvalVal As String) As String
'    Dim strFormat As String
'    AvalNum = "0"
'    strFormat = "#0"
'    If Val <> "" AND IsNumeric(Val) Then
'        If AvalVal = "9" Then
'            AvalVal = "0"
'        End If
'        If CLng(AvalVal) > 0 Then
'            strFormat = strFormat & "." & String(AvalVal, "0")
'            AvalNum = Format(Val, strFormat)
'        Else
'            AvalNum = Val
'        End If
'    End If
'
'End Function
'
'Private Sub SpreadLock(ByVal Col As Long, ByVal Col2 As Long)
'    With ssRst
'        .Row = 1: .Row2 = .MaxRows
'        .Col = Col: .Col2 = Col2
'        .BlockMode = True
'        .Lock = True
'        .Protect = True
'        .BlockMode = False
'    End With
'End Sub
'
'Public Sub LoadPic(ByVal strKey As String, ByVal iCol As Long, ByVal strType As String, ByVal Row As Long)
'    Dim strValue    As String
'    Dim lngIndex    As Long
'    Dim Col         As Long
'
'    If Row < 1 Then Exit Sub
'
'    objdic.KeyChange strKey
'
'    objdic.Fields("dpdiv") = ""
'    objdic.Fields("hldiv") = ""
'
'    SELECT Case UCase(strType)
'            Case "*":
'                lngIndex = 111: ' Col = SS_COL_HLDIV
'                strValue = "*"
'                objdic.Fields("hldiv") = strValue
'
'            Case "H":
'                lngIndex = 101: ' Col = SS_COL_HLDIV
'                strValue = "H"
'                objdic.Fields("hldiv") = strValue
'
'            Case "L":
'                lngIndex = 103: ' Col = SS_COL_HLDIV
'                strValue = "L"
'                objdic.Fields("hldiv") = strValue
'            Case "D":
'                lngIndex = 107: 'Col = SS_COL_DPDIV
'                strValue = "D"
'                objdic.Fields("dpdiv") = strValue
'            Case "P":
'                lngIndex = 105: 'Col = SS_COL_DPDIV
'                strValue = "P"
'                objdic.Fields("dpdiv") = strValue
'            Case "DP":
'                lngIndex = 109: 'Col = SS_COL_DPDIV
'                strValue = "DP"
'                objdic.Fields("dpdiv") = strValue
'            Case "C":
'                'Col = SS_COL_HLDIV
'                strValue = ""
'                objdic.Fields("hldiv") = strValue
'            Case "T":
'                lngIndex = 113: ' Col = SS_COL_TXT
'            Case "G":
'                lngIndex = 115: ' Col = SS_COL_GRPFG
'            Case "X":
'                lngIndex = 117: ' Col = SS_COL_RESULT
'    End SELECT
'    '
'    If (Row Mod 2) <> 0 Then
'        lngIndex = lngIndex + 1
'    End If
'    '
'    With ssRst
'        .Row = Row
'        .Col = iCol
'        .CellType = 9                                   'CellTypePicture
'        If strType = "C" Then
'            .CellType = 1                               'CellTypeEdit
''            .Value = objDic.Fields("rstcd")
'            objdic.Fields("rstcd") = Trim(.Value)
'            .ForeColor = vbBlack
'            .TypeHAlign = 2 'TypeHAlignCenter
'            .TypeVAlign = 2 'TypeVAlignCenter
'
'        ElseIf strType = "*" Then
'            .CellType = 1                               ' CellTypeEdit
'            .TypeHAlign = 2                             'TypeHAlignCenter
'            .TypeVAlign = 2                             'TypeVAlignCenter
'            objdic.Fields("rstcd") = Trim(.Value)
'            .Value = objdic.Fields("rstcd") & strType
'            .ForeColor = vbRed   '&H7477EF              '약간 붉은색
'            Exit Sub
'        ElseIf strType = "H" Then
'            .CellType = 1 'CellTypeEdit
'            .TypeHAlign = 2 'TypeHAlignCenter
'            .TypeVAlign = 2 'TypeVAlignCenter
'            objdic.Fields("rstcd") = Trim(.Value)
'            .Value = objdic.Fields("rstcd") & "▲"
'            .ForeColor = &H7477EF   '&H7477EF  '약간 붉은색
'            Exit Sub
'        ElseIf strType = "L" Then
'            .CellType = 1 ' CellTypeEdit
'            .TypeHAlign = 2 'TypeHAlignCenter
'            .TypeVAlign = 2 'TypeVAlignCenter
'            objdic.Fields("rstcd") = Trim(.Value)
'            .Value = objdic.Fields("rstcd") & "▼"
'            .ForeColor = &HE48372  '&HDF6A3E       '   &HE48372     '약간 파란색
'            Exit Sub
'        Else
'            .CellType = 1 'CellTypeEdit
'            .TypeHAlign = 2 'TypeHAlignCenter
'            .TypeVAlign = 2 'TypeVAlignCenter
'            .ForeColor = vbRed
'            .FontSize = 11
'            .FontBold = True
'            .RowHeight(.Row) = 13.3
'            objdic.Fields("rstcd") = Trim(.Value)
'            .Value = objdic.Fields("rstcd") & strType
'        End If
'    End With
'End Sub
'
'Private Sub PrgSet(ByVal intValue As Integer)
'    Dim intRest As Integer
'    If m_PrgSts = False Then Exit Sub
'    intRest = mvarPrgBar.Max - mvarPrgBar.Value
'    If intValue > intRest Then
'        intValue = intRest
'    End If
'    mvarPrgBar.Value = mvarPrgBar.Value + intValue
'End Sub
'
'Public Sub PrgBarInit()
'    If m_PrgSts = False Then Exit Sub
'    With mvarPrgBar
'        .Visible = True
'        .Min = 1
'        .Max = 250
'        .Value = .Min
'    End With
'End Sub
'
'Public Sub PrgBarClear()
'    If m_PrgSts = False Then Exit Sub
'    If Not mvarPrgBar Is Nothing Then mvarPrgBar.Visible = False
'    Set mvarPrgBar = Nothing
'End Sub
''
''/* Business Method */
'Public Function GetStringPtInfo() As String
'    GetStringPtInfo = mvarPtId & vbTab & mvarPtNm & vbTab & mvarSexAge & vbTab & _
'        mvarDOB & vbTab & mvarLocation & vbTab & mvarOrdDoctNm & vbTab & mvarSpcNm
'    If mvarPtType = RESULT_BY_EQUIPMENT Or mvarPtType = RESULT_BY_WORKSHEET Then
'        GetStringPtInfo = mvarAccNo & vbTab & GetStringPtInfo
'    End If
'End Function
'Public Function GetEQPStringPtInfo() As String
'    GetEQPStringPtInfo = mvarPtId & vbTab & mvarPtNm & vbTab & mvarSexAge & vbTab & _
'                        mvarDOB & vbTab & mvarLocation & vbTab & mvarOrdDoctNm & vbTab & mvarSpcNm
'
'    GetEQPStringPtInfo = mvarAccNo & vbTab & GetEQPStringPtInfo
'
'End Function
'Public Sub ResultCheck(ByVal iCol As Integer, Optional ByVal Row As Long = -1)
'    Dim Value   As String
'    Dim sLabNo  As String
'    Dim sTestCD As String
'
'    If Row = -1 Then
'        Row = ssRst.ActiveRow
'    End If
'
'    ssRst.Row = Row
'    ssRst.Col = tblCol.tcLABNO:  sLabNo = ssRst.Value
'    ssRst.Col = iCol - 1:        sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'    ssRst.Col = iCol:            Value = Trim(ssRst.Value)
'
''    MsgBox sLabNo & "       " & sTestCD
'
'    objdic.KeyChange sLabNo & COL_DIV & sTestCD
'
'    With objdic
'        '결과가 널이면 초기화.
'        If Value = "" Then
'            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'            If .Fields("rsttype") = "N" Or .Fields("rsttype") = "A" Then
'                .Fields("rsttype") = ""
'            ElseIf .Fields("rsttype") = "F" Or .Fields("rsttype") = "R" Then
'                .Fields("rsttype") = .Fields("rsttype")
'            End If
'            .Fields("rstval") = ""
'            .Fields("rstcd") = ""
'            .Fields("dpdiv") = ""
'            .Fields("hldiv") = ""
'            Exit Sub
'        End If
'        '결과가 들어왔고 결과유형이 널이 아니면 결과유형에 NUMERIC 혹은 ALPHA 셋팅.
'        If .Fields("rsttype") = "N" Or .Fields("rsttype") = "A" Or .Fields("rsttype") = "" Then
'            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'            If Not (Value Like "*+" Or Value Like "*-") AND IsNumeric(Value) Then
'                .Fields("rsttype") = "N"
'            Else
'                .Fields("rsttype") = "A"
'            End If
'        End If
'    End With
'
'    ssRst.Row = ssRst.ActiveRow
'    ssRst.Col = iCol
'End Sub
'
'Public Sub MResultCheck()
'
''    Dim Value As String
''    Dim Row As Long
''    Row = ssRst.ActiveRow
''    ssRst.Row = Row
''    ssRst.Col = SS_COL_RESULT
''    Value = Trim(ssRst.Value)
''    With Me.Result.Item(Row)
''        '결과가 널이면 초기화.
''        If Value = "" Then
''            LoadPic "C", Row
''            If .MRstType = "N" Or .MRstType = "A" Then
''                .MRstType = ""
''            ElseIf .MRstType = "F" Or .MRstType = "R" Then
''                .MRstType = .MRstType
''            End If
''            'ElseIf .MRstType = "" Then
''            '   If .MRstType = "" AND (.MRstCd <> "" Or .MRstVal <> "") Then
''                  .MRstVal = "": .MRstCd = "": .MDPDiv = "": .MHLDiv = ""
''            '   End If
''            'End If
''            Exit Sub
''        End If
''        '결과가 들어왔고 결과유형이 널이 아니면 결과유형에 NUMERIC 혹은 ALPHA 셋팅.
''        If .MRstType = "N" Or .MRstType = "A" Or .MRstType = "" Then
''            LoadPic "C", Row
''            If Not (Value Like "*+" Or Value Like "*-") AND IsNumeric(Value) Then
''                .MRstType = "N"
''            Else
''                .MRstType = "A"
''            End If
''        End If
''       'ElseIf .MRstType = "" Then
''       '   If Not (Value Like "*+" Or Value Like "*-") AND IsNumeric(Value) Then
''       '      .MRstType = "N"
''       '   Else
''       '      .MRstType = "A"
''       '   End If
''       '
''    End With
''    ssRst.Row = ssRst.ActiveRow
''    ssRst.Col = SS_COL_RESULT
'End Sub
'
'Public Sub NumValCheck(ByVal iCol As Long, Optional ByVal Row As Long = -1)
'    Dim blnPanic    As Boolean
'    Dim blnDelta    As Boolean
'
'    Dim Value       As String
'    Dim sLabNo      As String
'    Dim sTestCD     As String
'
'    Dim dblDeltaNum As Double
'
'    If Row = -1 Then
'       Row = ssRst.ActiveRow
'    End If
'
'    ssRst.Row = Row
'    ssRst.Col = iCol
'    Value = Trim(ssRst.Value)
'
'    ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'    ssRst.Col = tblCol.tcLABNO: sLabNo = Trim(ssRst.Value)
'
'    With objdic
'        objdic.KeyChange sTestCD & COL_DIV & sLabNo
'
'        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'
'        If Trim(Value) = "" Then Exit Sub
'        If Not IsNumeric(Value) Then Exit Sub
'
'        '/* HIGH,LOW CHECK */
'        If Val(.Fields("refvalfrom")) <> 0 Or Val(.Fields("refvalto")) <> 0 Then
'            If Val(Value) < .Fields("refvalfrom") Then
'                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "L", Row
'            End If
'            If Val(Value) > Val(.Fields("refvalto")) Then
'                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "H", Row
'            End If
'        End If
'        '/* PANIC CHECK */
'        If .Fields("panicfg") = "1" Then
'            If Val(Value) < Val(.Fields("panicfrval")) Or Val(Value) > Val(.Fields("panictoval")) Then
'                blnPanic = True
'            End If
'        End If
'        '/* DELTA CHECK */
'        If .Fields("deltafg") = "1" AND .Fields("lastrst") <> "" AND IsNumeric(.Fields("lastrst")) Then
'            If Val(.Fields("lastrst")) = 0 Then
'                dblDeltaNum = 0
'            Else
'                dblDeltaNum = ((Val(.Fields("lastrst")) - Val(Value)) / Val(.Fields("lastrst"))) * 100
'            End If
'            If dblDeltaNum < CDbl(.Fields("deltaval1")) * (-1) Or dblDeltaNum > CDbl(.Fields("deltaval2")) Then
'                blnDelta = True
'            End If
'        End If
'        '/* PANIC,DELTA PICTURE LOAD */
'        If blnPanic = True AND blnDelta = True Then
'            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "DP", Row
'        Else
'            If blnPanic = True Then
'                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "P", Row
'            End If
'            If blnDelta = True Then
'                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "D", Row
'            End If
'        End If
'    End With
'    ssRst.Row = ssRst.ActiveRow
'    ssRst.Col = iCol
'End Sub
'
'Public Function IsAvalVal(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
'    Dim Value   As String
'    Dim sTestCD As String
'    Dim sLabNo  As String
'
'    IsAvalVal = False
'    If Row = -1 Then
'        Row = ssRst.ActiveRow
'    End If
'
'    ssRst.Row = Row
'    ssRst.Col = iCol
'    Value = Trim(ssRst.Value)
'
'    If Value = "" Then
'        IsAvalVal = True
'        Exit Function
'    End If
'
'    With ssRst
'        .Row = .ActiveRow
'        .Col = tblCol.tcLABNO: sLabNo = .Value
'        .Col = iCol - 1: sTestCD = medGetP(.Value, 1, COL_DIV)
'        If objdic.Exists(sLabNo & COL_DIV & sTestCD) Then
'            objdic.KeyChange sLabNo & COL_DIV & sTestCD
'        End If
'        If IsNumeric(Value) = False Then
'            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'            Exit Function
'        Else
'            If objdic.Fields("avalval") = "9" Then
'                objdic.Fields("rstcd") = Value
'                objdic.Fields("rstval") = Value
'            End If
'        End If
'        If objdic.Fields("avalval") = "9" Then
'            '유효숫자 관리를 않함. 숫자형이면 다 허용 -- by 김미경
'            IsAvalVal = True
'            Exit Function
'        ElseIf objdic.Fields("avalval") = "0" Then
'        '정수형만 관리
'            Value = CStr(Val(Value))
'            If Value Like "*." Or Value Like "*.*" Then
'                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'                Exit Function
'            End If
'            objdic.Fields("rstcd") = Value
'            objdic.Fields("rstval") = Value
'        Else
'            '유효숫자 관리
'            If Len(medGetP(Value, 2, ".")) <> objdic.Fields("avalval") Then
'                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'                Exit Function
'            End If
'        End If
'    End With
'    '
'    IsAvalVal = True
'    '
'End Function
'
'Public Function IsAlphaCd(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
'    Dim RS      As RECORDSET
'    Dim Value   As String
'    Dim sLabNo  As String
'    Dim sTestCD As String
'    Dim Col     As Long
'
'    ' LC2_ItemResult = "C110"    Item - 결과코드
'    ' LC2_WAResult = "C111"     WorkArea - 결과코드
'    IsAlphaCd = False
'    If Row = -1 Then Row = ssRst.ActiveRow
'
'    ssRst.Row = Row
'    ssRst.Col = tblCol.tcLABNO: sLabNo = ssRst.Value
'    ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'
'    objdic.KeyChange sLabNo & COL_DIV & sTestCD
'
'    If ssRst.Value = "" Then
'        ssRst.Row = Row:  ssRst.Col = iCol
'    End If
'    ssRst.Col = iCol
'
'    Value = objdic.Fields("rstcd")
'
'
'
'    '/* 아이템과 결과코드 연결 체크 */
'    Set RS = GetRstCdSQL(LC2_ItemResult, sTestCD, Value)
'
'    If RS.EOF Then
'        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'        objdic.Fields("rsttype") = ""
'        Set RS = Nothing
'        Exit Function
'    End If
'
'
'    '/* 이상결과코드 체크 */
'    If Trim("" & RS.Fields("Field2").Value = "") Then
'        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'    Else
'        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "*", Row
'    End If
'
'    IsAlphaCd = True
'
'    RS.RsClose
'    Set RS = Nothing
'    '
'End Function
'
'Public Function IsRateCd(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
'    Dim sTestCD As String
'    Dim sLabNo  As String
'
'    Dim Value   As String
'    Dim Col     As Long
'
'    IsRateCd = False
'    If Row = -1 Then Row = ssRst.ActiveRow
'
'    ssRst.Row = Row:
'    ssRst.Col = tblCol.tcLABNO: sLabNo = ssRst.Value
'    ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'
'    objdic.KeyChange sLabNo & COL_DIV & sTestCD
'
'    ssRst.Col = iCol
'    Value = UCase(Trim(ssRst.Value))
'    '
'    '/* 비율결과 체크 : 반드시 숫자형으로 들어감, DB는 1:VALUE 형태*/
'    LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'    '
'    If CStr(Value) = "" Then
'        '결과가 들어가지 않으면 Validation체크 하지 않음.
'        IsRateCd = True
'        Exit Function
'    End If
'    '
'    If IsNumeric(Value) = True Then
'        '숫자형이 들어올 경우 정수타입인지 체크
'        If Value Like "*.*" Then
'            Exit Function
'        Else
'            IsRateCd = True
'        End If
'    Else
'        '문자가 들어올 경우 아이템과 비율결과코드 연결 체크
'        If UpdatableLAB031(LC2_ItemResult, sTestCD, Value) = False Then
'            Exit Function
'        Else
'            IsRateCd = True
'        End If
'    End If
'End Function
'
'Public Function IsFreeResult(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
'    Dim sLabNo  As String
'    Dim sTestCD As String
'    Dim Value   As String
'    Dim Col     As Long
'
'    IsFreeResult = False
'    If Row = -1 Then Row = ssRst.ActiveRow
'
'    ssRst.Row = Row
'    ssRst.Col = tblCol.tcLABNO: sLabNo = ssRst.Value
'    ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'    objdic.KeyChange sLabNo & COL_DIV & sTestCD
'
'    ssRst.Col = iCol
'    Value = Trim(ssRst.Value)
'    '
'    '/* Free형태의 결과는 (P_FreeRstLength)자리 이내 */
'    LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
'    '
'    If CStr(Value) = "" Then
'        '결과가 들어가지 않으면 Validation체크 하지 않음.
'        IsFreeResult = True
'        Exit Function
'    End If
'    '
'    If Len(Value) <= P_FreeRstLength Then
'        IsFreeResult = True
'    End If
'    '
'End Function
'
''탁과장님--------------------------------------------------------------------------------
''결과등록화면의 툴팁인데, 너비조정만 했어요.
'Public Sub SpToolTip(ByRef Row As Long, ByRef Col As Long, _
'                     ByRef MultiLine As Integer, ByRef ShowTip As Boolean, _
'                     ByRef TipText As String, ByRef TipWidth As Long)
'
''    Dim strMsg As String
''    With Me.Result.Item(Row)
''        ssRst.TextTip = 3
''        ssRst.TextTipDelay = mvarTextTipDelay
''        If Col = SS_COL_HLDIV Then
''            MultiLine = 1   '--->수정
''            If .RefValFrom = 0 AND .RefValTo = 0 Then
''                TipText = vbCrLf & " Reference Range : Undefined " & vbCrLf & " "
''            Else
''                TipText = vbCrLf & " Reference Range : " & .RefValFrom & " - " & .RefValTo & vbCrLf & " "
''            End If
''            TipWidth = 3500   '--->수정
''            ssRst.TextTipDelay = mvarTextTipDelay
''            Call ssRst.SetTextTipAppearance("굴림", 9, True, False, &HEEFDF2, &H996666)
''        ElseIf Col = SS_COL_LASTRST Then
''            MultiLine = 1   '--->수정
''            If .LastRst = "" Then Exit Sub
''            TipText = vbCrLf & " 최근 결과일시 : " & Format(.LastVfyDt, "0###-##-##") & " " & _
''                                                     Format(Mid(.LastVfyTm, 1, 4), "0#:##") & vbCrLf
''            TipWidth = 4000   '--->수정
''            ssRst.TextTipDelay = mvarTextTipDelay
''            Call ssRst.SetTextTipAppearance("굴림", 9, True, False, &HEEFDF2, &H996666)
''        ElseIf Col = SS_COL_EQPNM Then
''            If .autofg = "1" Then
''                MultiLine = 1   '--->수정
''                If .TransDtTm = "" Then Exit Sub
''                TipText = vbCrLf & "  - 자동화장비 결과 전송내역 -" & vbCrLf & vbCrLf & _
''                                   " 테스트명 : " & Trim(.TestNm) & vbCrLf & _
''                                   " 전송장비 : " & Trim(.EqpNm) & vbCrLf & _
''                                   " 전송일시 : " & Trim(.TransDtTm) & vbCrLf
''                'TipText = vbCrLf & strMsg & vbCrLf
''                TipWidth = 4000   '--->수정
''                ssRst.TextTipDelay = mvarTextTipDelay
''                Call ssRst.SetTextTipAppearance("굴림", 9, True, False, &HEEFDF2, &H996666)
''            End If
''        Else
''            ShowTip = False
''            Exit Sub
''        End If
''        ShowTip = True
''    End With
'End Sub
'
'Public Function DataEntry() As Boolean
'
''    Dim ii As Integer
''    Dim jj As Integer
''    Dim blnSqlErr As Boolean
''    Dim strSQL As String
''    Dim objRST As clsPatientInfo
''    Dim strResult As String
''    Dim strVfyDt As String
''    Dim strVfyTm As String
''    Dim strVfyId As String
''    Dim intScrReqInputCnt As Integer
''    Dim intScrReqTotCnt As Integer
''    Dim intDF As String
''    Dim aryDetail() As String
''    Dim aryCnt() As String
''    Dim blnAryDetailFirst As Boolean
''    Dim blnException As Boolean
''    Dim blnAllVfy As Boolean
''   '
''    DataEntry = True
''    '/* 결과 Item이 없는 경우 빠짐 */
''    If ssRst.MaxRows < 1 Then Exit Function
''    strVfyDt = Format(Date, "yyyymmdd")
''    strVfyTm = Format(Time, "hhmmss")
''    strVfyId = m_EmpId
''   '
''    '/* 현재 화면의 결과 등록상의 Required,Alternative건수 체크
''    ' intScrReqTotCnt : 화면상의 총R건수(단,상세항목R필드 제외)
''    ' intScrReqInputCnt :화면상에 결과입력된 R결과건수(단,상세항목R필드 제외)
''    '
''    intScrReqInputCnt = 0: intScrReqTotCnt = 0
''    intDF = 0:
''    blnAryDetailFirst = False
''    blnException = False
''    '/* 결과확인체크 Array구조 :
''    For ii = 1 To ssRst.MaxRows
''        With Me.Result.Item(ii)
''            '결과등록 화면상의 보류여부 Setting.
''            ssRst.Row = ii
''            ssRst.Col = SS_COL_EC
''            If Val(ssRst.Value) = 1 Then
''                .ExcFg = "1"
''                '화면상에 단 1건의 아이템이라도 보류된 아이템이 있으면 전체 Verify를 보류한다.
''                 '-> blnException = True
''                 blnException = True
''            Else
''                .ExcFg = "0"
''            End If
''
''            '======================================
''            '검사코드를 가지고 오기위해서...
''            'col=18번 값의 결과를 저장함
''            '예전값은(ssRst.Col = SS_COL_RESULT)
''            ''======================================
''
''            'ssRst.Col = SS_COL_RESULT
''            ssRst.Col = 18
''            'If ssRst.Value <> "" Then
''               '결과가 입력 되었으면 결과유형대로 저장.
''               ResultSet Trim(ssRst.Value), ii
''            'End If
''            '
''            If .DetailFg <> "" Then
''            '/ * 상세항목인경우 체크Array에 정보 저장. */
''                ' aryDetail : 상세항목중 하나라도 보류가 되었는지 여부 Check "$"
''                '                Required 총수 "$"
''                '                Required 입력카운트 "$"
''                '                Alternative 총수 "$"
''                '                Alternative 입력카운트
''                ' 주의 : 상세항목중 어떤 한 아이템이라도 보류버튼이 눌려지면 모두 Verify하지 않는다.
''                '#################################################
''                '#  1999.07.18  수정  by 김미경                                          #
''                '#  상세항목의 대표코드가 보류되면 전체보류체크한다.     #
''                '#################################################
''
''                If blnAryDetailFirst = False Then
''                    ReDim aryDetail(Val(.DetailFg))
''                    blnAryDetailFirst = True
''                    aryDetail(Val(.DetailFg)) = "0$0$0$0$0"
''                Else
''                    If Val(.DetailFg) > UBound(aryDetail) Then
''                        ReDim Preserve aryDetail(Val(.DetailFg))
''                    End If
''                    If aryDetail(Val(.DetailFg)) = "" Then
''                        '상세체크 Array가 초기화되지 않았으면 초기화.
''                        aryDetail(Val(.DetailFg)) = "0$0$0$0$0"
''                    End If
''                End If
''                intDF = Val(.DetailFg)
''                aryCnt = Split(aryDetail(intDF), "$")
''
''                '#############################
''                '대표코드인 경우 전체항목수 Add
''                '#############################
''                If .RstDiv = "*" Then
''                    intScrReqTotCnt = intScrReqTotCnt + 1
''                    If .ExcFg = "1" Then
''                        '상세항목중 보류된 행이 나타나면 상세항목 전체보류체크한다.
''                        aryCnt(0) = .ExcFg
''                    End If
''                Else
''                    If .ExcFg = "1" Then
''                        '상세항목중 보류된 행이 나타나면 상세항목 전체보류체크한다.
''                        aryCnt(0) = .ExcFg
''                    End If
''                End If
''
''                '###############################
''                'RstDiv="*"인 경우(대표코드)는 제외
''                '###############################
''                If .RstDiv = "R" Then
''                    '상세항목중 Required 총수 Add
''                    aryCnt(1) = CStr(Val(aryCnt(1)) + 1)
''
''                    '탁과장님---------------------------------------------------------
''                    '2001.6.22 수정 : 결과가 입력안되었을때 .RstVal에 "null"이 들어감
''                    'If .RstCd <> "" Or .RstVal <> "" Then
''                    If Trim(ssRst.Value) <> "" Then
''                    '-----------------------------------------------------------------
''                        '상세항목중 Required입력수 Add
''                        aryCnt(2) = CStr(Val(aryCnt(2)) + 1)
''                    End If
''                ElseIf .RstDiv = "A" Then
''                    '상세항목중 Alternative 총수 Add
''                    aryCnt(3) = CStr(Val(aryCnt(3)) + 1)
''                    '탁과장님---------------------------------------------------------
''                    '2001.6.22 수정 : 결과가 입력안되었을때 .RstVal에 "null"이 들어감
''                    'If .RstCd <> "" Or .RstVal <> "" Then
''                    If Trim(ssRst.Value) <> "" Then
''                    '-----------------------------------------------------------------
''                        '상세항목중 Alternative입력수 Add
''                        aryCnt(4) = CStr(Val(aryCnt(4)) + 1)
''                    End If
''                End If
''                aryDetail(intDF) = Join(aryCnt, "$")
''            Else
''                '/* 그룹 혹은 단일 아이템인 경우 전체Verify여부를 판단하기 위해 카운트.
''                '그룹 혹은 단일 아이템수 ADD
''                intScrReqTotCnt = intScrReqTotCnt + 1
''                If Trim(ssRst.Value) <> "" Then
''                    If .ExcFg = "0" Then
''                        ' 보류가 되지않고 입력된 수를 카운트 한다. */
''                        intScrReqInputCnt = intScrReqInputCnt + 1
''                    End If
''                End If
''
''            End If
''        End With
''    Next
''    '
''    '/* Partial Verify처리 및 전체 Verify여부 체크 */
''    For ii = 1 To ssRst.MaxRows
''        With Me.Result.Item(ii)
''            If .DetailFg = "" Then
''                '/* 그룹 혹은 단일 아이템인 경우 VfyDt셋팅 처리 */
''                '1)보류되지 않았고,
''                '2)결과가 들어가 있으면 VfyDt셋팅.
''                If .ExcFg = "0" Then
''                    '1)번 조건 성립
''                    If (.rstval <> "" AND UCase(.rstval) <> "NULL") Or .RstCd <> "" Then
''                        '2)번 조건 성립
''                        .VfyDt = strVfyDt
''                        .VfyTm = strVfyTm
''                        .VfyId = strVfyId
''                    End If
''                End If
''            Else
''                '/* 상세항목인 경우 VfyDt셋팅 처리 */
''                '1)일단 상세항목의 모든 아이템이 보류가 되지 않았고,
''                '2)모든 Required필드가 입력되어 있고,
''                '3)하나이상의 Alternative아이템에'결과가 입력되어 있으면에 VfyDt를 셋팅한다.
''                '참고:1),2),3)조건이 모두 만족되면 "*":대표코드에도 VfyDt를 셋팅한다.
''                intDF = Val(.DetailFg)
''                aryCnt = Split(aryDetail(intDF), "$")
''                If aryCnt(0) = "0" Then
''                    '1)번조건 성립
''                    If Val(aryCnt(1)) = Val(aryCnt(2)) Then
''                        '2)번조건 성립
''                        If (Val(aryCnt(3)) = 0) Or (Val(aryCnt(3)) > 0 AND Val(aryCnt(4)) > 0) Then
''                            '3)번조건 성립
''                            .VfyDt = strVfyDt
''                            .VfyTm = strVfyTm
''                            .VfyId = strVfyId
''                            If .RstDiv = "*" Then
''                                '상세항목의 모든 아이템에 대한 결과처리.
''                                ' 상세항목별로 매번 하지 않고 대표코드시 한번만 처리한다.
''                                aryCnt(4) = aryCnt(3)
''                            End If
''                        End If
''                    End If
''                End If
''                aryDetail(intDF) = Join(aryCnt, "$")
''            End If
''            '/* 텍스트소견 결과 체크 */
''            If .TxtFg = "0" Then
''                If .TextRst <> "" Then
''                    .TxtFg = "1"
''                End If
''            End If
''        End With
''    Next ii
''    '
''    '/* 접수내역의 FootNote체크 */
''    With Me
''        If .FootNoteFg = "0" Then
''            If .FootNote <> "" Then
''                .FootNoteFg = "1"
''            End If
''        End If
''    End With
''    '
''    '/* 결과등록에 대한 Status Follow -Up 처리(전체 Verify처리) */
''    '1)화면상에 보류여부가 없어야 하고,
''    '2)화면상의 Requried입력카운트에 접수내역에 보관되어 있는 Required입력 카운트를 더한 값이
''    '  접수내역의 총Required카운트와 같아야 한다.
''    '3)상세항목의 경우 보류여부가 하나도 없어야 한다.
''    '4)화면상의 Alternative입력카운트에 접수내역에 보관되어 있는 Required입력 카운트를 더한 값이
''    '  접수내역의 총Required카운트와 같아야 한다.
''    '위에서 입력카운트라 함은 Verify된 수를 의미한다.
''    blnAllVfy = False
''    With Me
''        If intScrReqTotCnt > 0 Then
''            '화면상에 그룹 혹은 단일 아이템이 존재
''            .ReqInputCnt = .ReqInputCnt + intScrReqInputCnt
''        End If
''        If blnAryDetailFirst = True Then
''            '화면상에 상세항목이 존재
''            For ii = 0 To UBound(aryDetail)
''                If aryDetail(ii) <> "" Then
''                    aryCnt = Split(aryDetail(ii), "$")
''                    If aryCnt(0) = "0" Then
''                        '상세항목별로 전체가 보류되지 않고 정상일때 접수내역에 카운트 반영.
''                        '주의:상세항목중 Required필드는 접수내역의 ReqInputCnt에 Add
''
''                        If Val(aryCnt(1)) = Val(aryCnt(2)) AND Val(aryCnt(3)) = Val(aryCnt(4)) Then
''                            '##############################################
''                            '#  1999.07.18  수정  by 김미경                                    #
''                            '#  상세항목의 대표코드만 입력 카운트에 반영한다.     #
''                            '##############################################
''                            .ReqInputCnt = .ReqInputCnt + 1
''                            '.ReqInputCnt = .ReqInputCnt + Val(aryCnt(2))
''                            '.AltInputCnt = .AltInputCnt + Val(aryCnt(4))
''                        End If
''                    End If
''                End If
''            Next
''
''
''        End If
''        '
''        If blnException = False Then
''            '1)번 조건 성립
''            'If .ReqTotCnt = .ReqInputCnt AND .AltInputCnt = .AltTotCnt Then
''            If .ReqTotCnt = .ReqInputCnt Then
''                '2)번 및 3)번 조건 성립
''                blnAllVfy = True   '전체 Verify : Status는 결과로 바뀜.
''                .StsCd = enStsCd.StsCd_LIS_FinRst
''            End If
''        End If
''    End With
''    '
''    '/* 결과등록 */
''    ' LAB102 : StsCd : 전체 Verify시 수정됨. :UPDATE
''    ' LAB201 : Status,ReqinputCnt,AltInputCnt,FootNote,FootNoteFg,VfyDt,VfyTm,VfyId,
''    '              *RmkCd : UPDATE
''    ' LAB302 : *RstVal,*RstCd,*HLDiv,*DPDiv,VfyDt,VfyTm,VfyId,TxtFg,*RstType,*EqpCd :UPDATE
''    ' LAB303 : *RstTxt : INSERT
''    ' LAB304 : *RstText : INSERT
''    ' *표시된 필드는 스프레드화면 동작시 이미 셋팅되어 있는 필드.
''    Set objRST = Me
''    If blnAllVfy = True Then
''        DataEntry = OraDS.ResultDataEntry(objRST, strVfyDt, strVfyTm, strVfyId, strVfyId)
''    Else
''        DataEntry = OraDS.ResultDataEntry(objRST, "", "", "", strVfyId)
''    End If
''    '
''    'With OraDS
''    '   If .LastServerErr <> 0 Then
''    '      glngErrorNo = .LastServerErr
''    '      gstrErrorMsg = .LastServerErrText
''    '      DataEntry = False
''    '      Exit Function
''    '   End If
''    'End With
''    '
''    'OraErr.Reset
''    '
'End Function
'
'Public Function ItemDataEntry() As Boolean
'
''    Dim ii As Integer
''    Dim jj As Integer
''    Dim blnSqlErr As Boolean
''    Dim strSQL As String
''    Dim objRST As clsPatientInfo
''    Dim strResult As String
''    Dim blnAllVfy As Boolean
''    Dim strVfyDt As String
''    Dim strVfyTm As String
''    Dim strVfyId As String
''    Dim aryCnt() As String
''    Dim aryPtInfo(0 To 3) As String
''
''    '
''    ItemDataEntry = True
''    '/* 결과 Item이 없는 경우 빠짐 */
''    If ssRst.MaxRows < 1 Then Exit Function
''    strVfyDt = Format(Date, "yyyymmdd")
''    strVfyTm = Format(Time, "hhmmss")
''    strVfyId = m_EmpId
''    '
''    For ii = 1 To ssRst.MaxRows
''        With Me.Result.Item(ii)
''            If .RstDiv <> "*" AND .RstDiv = "R" Then
''                '결과등록 화면상의 보류여부 Setting.
''                ssRst.Row = ii
''                ssRst.Col = SS_COL_EC
''                If Val(ssRst.Value) = 1 Then
''                  .ExcFg = "1"
''                Else
''                  .ExcFg = "0"
''                End If
''
''                '환자ID
''                ssRst.Row = ii
''                ssRst.Col = SS_COL_PTID
''                mvarPtId = Trim(ssRst.Value)
''                '
''                ssRst.Row = ii
''                '결과등록시 결과코드가 아닌결과 값을 보여주기 위해서
''                'COL=18번에 저장해야할값을 가지고 있고.
''                'COL=SS_COL_RESULT에 결과코드 값을 보여줌
''                '때문에 저장할때는 COL=18번 값을 저장해야함
''
''                'ssRst.Col = SS_COL_RESULT
''                ssRst.Col = 18
''
''                '-------------------------------------------------------------------------
''                ' Item별 결과등록시 보류는 Update대상에서 제외시킴.  -1999.08.13 by 김미경
''                ' ---
''                ' Item별 결과등록시 보류는 결과만 Update.  -2001.06.26 by 김미경
''                '-------------------------------------------------------------------------
''                'If Trim(ssRst.Value) <> "" AND .ExcFg = "0" Then
''                If Trim(ssRst.Value) <> "" Then
''                '결과가 입력 되었으면 결과유형대로 저장.
''                    ResultSet Trim(ssRst.Value), ii
''                    If .ExcFg = "0" Then
''                        .VfyDt = strVfyDt
''                        .VfyTm = strVfyTm
''                        .VfyId = strVfyId
''                        aryCnt = Split(.ReqAltCnt, vbTab)
''                        aryCnt(1) = CStr(Val(aryCnt(1)) + 1)
''                        If Val(aryCnt(0)) = Val(aryCnt(1)) Then
''                           .DataFg = "1"         '전체Verify
''                        Else
''                           .DataFg = "0"         'Partial Verify만
''                        End If
''                        .ReqTotCnt = aryCnt(0)
''                        .ReqAltCnt = aryCnt(1) '접수내역의 ReqInputCnt를 Update하기 위해 셋팅.
''                    Else
''                       aryCnt = Split(.ReqAltCnt, vbTab)
''                       .ReqAltCnt = aryCnt(1)
''                    End If
''                End If
''            '
''            End If
''        End With
''    Next ii
''    '
''    '
''    '/* 결과등록 */
''    ' LAB102 : StsCd : 전체 Verify시 수정됨. :UPDATE
''    ' LAB201 : Status,ReqinputCnt,AltInputCnt,FootNote,FootNoteFg,VfyDt,VfyTm,VfyId,
''    '              *RmkCd : UPDATE
''    ' LAB302 : *RstVal,*RstCd,*HLDiv,*DPDiv,VfyDt,VfyTm,VfyId,TxtFg,*RstType,*EqpCd :UPDATE
''    ' LAB303 : *RstTxt : INSERT
''    ' LAB304 : *RstText : INSERT
''    ' *표시된 필드는 스프레드화면 동작시 이미 셋팅되어 있는 필드.
''    Set objRST = Me
''    OraDS.ItemDataEntry objRST, strVfyDt, strVfyTm, strVfyId, strVfyId
''    '
''    With OraDS
''        If .LastServerErrText <> "" Then
''            ItemDataEntry = False
''            glngErrorNo = .LastServerErr
''            gstrErrorMsg = .LastServerErrText
''            Exit Function
''        End If
''    End With
''    '
''    OraErr.Reset
''    '
'End Function
'
'Public Function ModifyEntry() As Boolean
'
''    Dim ii As Integer
''    Dim jj As Integer
''    Dim blnSqlErr As Boolean
''    Dim strSQL As String
''    Dim objRST As clsPatientInfo
''    Dim strResult As String
''    Dim strMfyRsn As String
''    Dim strMfyTm As String
''    Dim strMfyId As String
''    Dim aryCnt() As String
''    Dim blnChangeEqpCd As Boolean
''    Dim blnChangeResult As Boolean
''    '
''    ModifyEntry = True
''    '/* 결과 Item이 없는 경우 빠짐 */
''    If ssRst.MaxRows < 1 Then Exit Function
''    strMfyRsn = Format(Date, "yyyymmdd")
''    strMfyTm = Format(Time, "hhmmss")
''    strMfyId = m_EmpId
''    '
''    '/* DataFg : 결과수정 화면에서(결과등록은 사용않함) 결과 혹은 장비가 수정되었는지
''    '                여부를 보관한다. '0':변화없음,'1':결과변화,'2':결과,장비변화,'3':장비변화 */
''    '/* 화면상의 결과가 수정되었지 여부를 체크 */
''    For ii = 1 To ssRst.MaxRows
''        With Me.Result.Item(ii)
''            ssRst.Row = ii
''            If .RstDiv <> "*" Then
''                blnChangeEqpCd = False
''                blnChangeResult = False
''                .DataFg = "0"              'Data Chnage Flag Default Setting
''                '
''                '결과코드 값을 저장하기 위해서 수정Col=18로
''                'ssRst.Col = SS_COL_RESULT
''
''                ssRst.Col = 18
''
''                '결과유형대로 저장.
''                MResultSet Trim(ssRst.Value), ii
''                If (Val(.rstval) <> Val(.MRstVal)) Or (.RstCd <> .MRstCd) Then
''                    '결과가 수정되었는지 여부를 체크한다.
''                    blnChangeResult = True
''                End If
''                If (.autofg <> "1") AND (.OTmpCd <> .EqpCd) Then
''                    '결과수정화면에서 장비가 바뀌었는지를 체크한다.
''                    blnChangeEqpCd = True
''                End If
''                If blnChangeResult = True AND blnChangeEqpCd = True Then
''                    .DataFg = "2"
''                ElseIf blnChangeResult = False AND blnChangeEqpCd = True Then
''                    .DataFg = "3"
''                ElseIf blnChangeResult = True AND blnChangeEqpCd = False Then
''                    .DataFg = "1"
''                End If
''            End If
''        End With
''    Next
''    '
''    '/* 결과수정 */
''    ' LAB102 : StsCd : 처방BODY별로 Stauts가 '4'이면 '5'로 Follow-Up
''    ' LAB302 : *RstVal,*RstCd,*HLDiv,*DPDiv,TxtFg,*RstType,*EqpCd :UPDATE
''    ' LAB304 : *RstText : INSERT
''    ' LAB305 : *RstText : INSERT
''    ' LAB308 :      *      : INSERT
''    ' *표시된 필드는 스프레드화면 동작시 이미 셋팅되어 있는 필드.
''    Set objRST = Me
''    OraDS.ResultModifyEntry objRST, strMfyRsn, strMfyTm, strMfyId
''    '
''    With OraDS
''        If .LastServerErr <> 0 Then
''            glngErrorNo = .LastServerErr
''            gstrErrorMsg = .LastServerErrText
''            ModifyEntry = False
''            Exit Function
''        End If
''    End With
''    '
''    OraErr.Reset
''    '
'End Function
'
'Private Sub ResultSet(ByVal Value As String, ByVal Row As Integer)
''    '
''    '/* 결과등록 유형에 따른 결과 저장 */
''    '   - RstVal : 수치형 결과, RstCd : 코드 혹은 문자형 결과
''    With Me.Result.Item(Row)
''        SELECT Case .rsttype
''            Case "N":            '수치결과는 RstCd에 결과값 저장
''                .rstval = Val(Value)
''                .RstCd = Value
''            Case "A":            '결과코드만 RstCd에 저장.
''                .rstval = "null"
''                .RstCd = Value
''            Case "R":            '비율이 수치로 들어가는 경우와 코드로 들어가는 경우가 틀림.
''                If IsNumeric(Value) Then
''                    .rstval = Value
''                    .RstCd = "1:" & Value
''                Else                 '문자형으로 자유로이 결과가 들어감.
''                    .rstval = "null"
''                    .RstCd = Value
''                End If
''            Case "F":
''                .rstval = "null"
''                .RstCd = Value
''            Case Else:
''                .rstval = "null"
''                .RstCd = Value
''        End SELECT
''    End With
'End Sub
'
'Private Sub MResultSet(ByVal Value As String, ByVal Row As Integer)
''    '
''    '/* 결과수정시 결과유형에 따른 결과 저장 */
''    '   - MRstVal : 수치형 결과, MRstCd : 코드 혹은 문자형 결과
''    With Me.Result.Item(Row)
''        If .MRstType = "" Then
''            If Trim(.MRstVal) <> "" AND Trim(.MRstCd) <> "" Then
''                .MRstType = "N"
''            End If
''            If Trim(.MRstVal) = "" AND Trim(.MRstCd) <> "" Then
''                .MRstType = "A"
''            End If
''        End If
''        SELECT Case .MRstType
''            Case "N":            '수치결과는 MRstCd에 결과값 저장
''                .MRstVal = Val(Value)
''                .MRstCd = Value
''            Case "A":            '결과코드만 MRstCd에 저장.
''                .MRstVal = ""
''                .MRstCd = Value
''            Case "R":            '비율이 수치로 들어가는 경우와 코드로 들어가는 경우가 틀림.
''                If IsNumeric(Value) Then
''                    .MRstVal = Value
''                    .MRstCd = "1:" & Value
''                Else                 '문자형으로 자유로이 결과가 들어감.
''                    .MRstVal = ""
''                    .MRstCd = Value
''                End If
''            Case "F":
''                .MRstVal = ""
''                .MRstCd = Value
''        End SELECT
''    End With
'End Sub
'
'Private Function VerifyOk() As Long
''    VerifyOk = MsgBox("모든 결과항목을 확인 하시겠읍니까?", _
''                        vbYesNo + vbQuestion, "결과확인")
'End Function
'
'Public Function SSCol(ByVal Fld As String) As Integer
''    SELECT Case UCase(Fld)
''        Case "TESTNM":
''            SSCol = SS_COL_TESTNM
''        Case "RESULT":
''            SSCol = SS_COL_RESULT
''        Case "EC":
''            SSCol = SS_COL_EC
''        Case "LASTRST":
''            SSCol = SS_COL_LASTRST
''        Case "RSTUNIT":
''            SSCol = SS_COL_RSTUNIT
''        Case "HLDIV":
''            SSCol = SS_COL_HLDIV
''        Case "DPDIV":
''            SSCol = SS_COL_DPDIV
''        Case "TXT":
''            SSCol = SS_COL_TXT
''        Case "GRPFG":
''            SSCol = SS_COL_GRPFG
''        Case "EQPNM":
''            SSCol = SS_COL_EQPNM
''        Case "REFVAL":
''            SSCol = SS_COL_REFVAL
''        Case "SEQ":
''            SSCol = SS_COL_SEQ
''        Case "ACCNO":
''            SSCol = SS_COL_ACCNO
''    End SELECT
'End Function
'
'Public Sub mnuUtil1_Click()
'    '
''    ItemRstCodePop
'    '
'End Sub
'
'Public Sub mnuUtil2_Click()
'    '
'    InstrumentListPop
'    '
'End Sub
'
'Public Function GetSqlTransferPop(ByVal EqpCd As String)
''    GetSqlTransferPop = OraDS.SqlStmtTransfer(EqpCd)
'End Function
'
'Private Sub InstrumentListPop()
''    Dim PosX As Long, PosY As Long
''    Dim Wid As Long, Hgt As Long
''   'kmk 여기 고쳐야 됨
''    Call ssRst.GetCellPos(SSCol("EQPNM"), ssRst.ActiveRow, PosX, PosY, Wid, Hgt)
''    Set objCodeList = New clsCodeList
''    With objCodeList
''      'Set .MyOraSE = OraDS.OraSe
''      Set .MyDB = DBConn
''      .ListCaption = "Instrument List"
''      .ListColHeader = "Name" & vbTab & "Code"
''      .Width = 3000
''      .Height = 3000
''      .Top = SsTop + PosY + 1600
''      .Left = SsLeft + PosX - .Width
''      .Tag = "Instrument"
''      .CaptionOn = True
''      .MultiSel = False
''      .PopupList OraDS.SqlStmtInstrument, 2
''    End With
'End Sub
'
'Private Sub ItemRstCodePop(ByVal rCol As Integer, ByVal sTestCD As String)
'
'    Dim strSql As String
'    Dim PosX As Long, PosY As Long
'    Dim Wid As Long, Hgt As Long
'
'    'kmk 여기 고쳐야 됨
'    Call ssRst.GetCellPos(rCol, ssRst.ActiveRow, PosX, PosY, Wid, Hgt)
'    Set objCodeList = New clsCodeList
'    With objCodeList
'        Set .MyDB = DBConn
'        .ListCaption = "Result Code List"
'        .ListColHeader = "Name" & vbTab & "Code"
'        .Width = 4050
'        .Height = 3000
'        .Top = SsTop + PosY + 1700
'        .Left = SsLeft + PosX - 810
'        .Tag = "ResultCode"
'        .CaptionOn = True
'        .MultiSel = False
'        strSql = SqlStmtRstCode(sTestCD)
'        .PopupList strSql, 2
'    End With
'End Sub
'
'Private Sub objCodeList_ListClick(ByVal SelList As String)
'    Dim sTestCD As String
'    Dim sLabNo  As String
'
'
'    If Not IsNull(SelList) Then
'        ssRst.Row = ssRst.ActiveRow
'        ssRst.Col = ssRst.ActiveCol
'        ssRst.Value = medGetP(SelList, 2, vbTab)
'        ssRst.Col = tblCol.tcLABNO:  sLabNo = Trim(ssRst.Value)
'        ssRst.Col = ssRst.ActiveCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'        If objdic.Exists(sLabNo & COL_DIV & sTestCD) Then
'            objdic.KeyChange sLabNo & COL_DIV & sTestCD
'            objdic.Fields("rstcd") = medGetP(SelList, 1, vbTab)
'        End If
'    End If
'    Set objCodeList = Nothing
'End Sub
'
'Private Sub OpenContextMenu(FormName As Form, MenuName As Menu)
'    '
''    Call SendMessage(FormName.hWnd, WM_RBUTTONDOWN, 0, 0&)
''    '
''    FormName.PopupMenu MenuName
'    '
'End Sub
'
'Public Function GetSqlRstCd(ByVal testcd As String, _
'                            Optional pOraSession As Object, _
'                            Optional pDatabase As Object) As String
'    '
''    If OraDS.IsRstCode(TestCd) = True Then
''        GetSqlRstCd = OraDS.SqlStmtRstCode(TestCd)
''    Else
''        GetSqlRstCd = ""
''    End If
'    '
'End Function
'
'Public Function GetTestItemDataEntry(ByVal testcd As String, _
'                                     Optional pOraSession As Object, _
'                                     Optional pDatabase As Object) As String
''    Dim strTestNm As String
'    '   If pDatabase Is Nothing Then
'    '      OraDS.InitConnect
'    '   Else
'    '      OraDS.InitConnect pOraSession, pDatabase
'    '   End If
'    '
''    GetTestItemDataEntry = OraDS.GetTestItemDataEntry(TestCd, Format(Date, "yyyymmdd"))
'    '
'End Function
'
'Public Function GetRmkCdNm(ByVal RmkCd As String) As String
''    '
''    If RmkCd <> "" Then
''        GetRmkCdNm = OraDS.GetTempleteText1(LC4_Remark, RmkCd)
''    Else
''        GetRmkCdNm = ""
''    End If
'    '
'End Function
'
'Public Function GetRstCdNm(ByVal testcd As String, ByVal RstCd As String) As String
'    '
''    If RstCd <> "" Then
''        GetRstCdNm = OraDS.GetRstCdNm(TestCd, RstCd)
''    Else
''        GetRstCdNm = ""
''    End If
'    '
'End Function
'
'
'
'Public Sub LoadUnverifiedList(ByVal pWorkArea As String, ByVal pRcvDt As String, _
'                                    ByVal pBuildingCd As String, ByRef tblLabList As Object)
'
''    Dim I As Integer
''    Dim SqlStmt As String
''    Dim tmpRs As RECORDSET
''    Dim strKey As String
''    Dim strSvKey As String
''    Dim lngTCount As Long
''
''    lngTCount = 0
''
''    Set tmpRs = OpenRecordSet(OraDS.GetUnverifiedList(pWorkArea, pRcvDt, pBuildingCd))
''
''    With tblLabList
''        .ReDraw = False
''        .MaxRows = 0
''        .MaxRows = tmpRs.RecordCount
'''        For i = 1 To tmpRs.RecordCount
''
''        strSvKey = ""
''        While Not tmpRs.EOF
''            strKey = Trim("" & tmpRs.Fields("WorkArea").Value) & "-" & _
''                     Mid("" & tmpRs.Fields("AccDt").Value, 3, 6) & "-" & _
''                     Trim("" & tmpRs.Fields("AccSeq").Value)
''            If strKey <> strSvKey Then
''                lngTCount = lngTCount + 1
''                I = I + 1
''                .Row = I
''                .Col = 1: .Value = strKey
''                .Col = 2: .Value = "" & tmpRs.Fields("ReqTotCnt").Value
''                .Col = 3: .Value = Val("" & tmpRs.Fields("ReqTotCnt").Value) - Val("" & tmpRs.Fields("ReqInputCnt").Value)
''                .Col = 5: .Value = Val("" & tmpRs.Fields("TestDiv").Value)
''                .Col = 6: .Value = "" & tmpRs.Fields("AbbrNm5").Value: .ForeColor = DCM_LightBlue
''                .Col = 7: .Value = IIf(tmpRs.Fields("statfg").Value & "" = "1", "Y", ""): .ForeColor = DCM_LightRed
''                .Col = 8:
''                        If tmpRs.Fields("wardid").Value & "" <> "" Then
''                            .Value = tmpRs.Fields("wardid").Value & ""
''                            If tmpRs.Fields("hosilid").Value & "" <> "" Then
''                                .Value = .Value & "-" & tmpRs.Fields("hosilid").Value & ""
''                            End If
''                        Else
''                            .Value = tmpRs.Fields("deptcd").Value & ""
''                        End If
''                .Col = 9: .Value = tmpRs.Fields("ptid").Value & "" & "(" & tmpRs.Fields("ptnm").Value & "" & ")"
''                          .ForeColor = DCM_LightBlue
''
''                .Col = 10: .Value = tmpRs.Fields("spcyy").Value & "" & "-" & tmpRs.Fields("spcno").Value & ""
''                strSvKey = strKey
''
''            Else
''                .Col = 6: .Value = .Value & "," & "" & tmpRs.Fields("AbbrNm5").Value
''            End If
''
''            tmpRs.MoveNext
''        Wend
'''        Next
''        .MaxRows = lngTCount
''
''        .ReDraw = True
''    End With
''
''    tmpRs.RsClose
''    Set tmpRs = Nothing
''
'End Sub
'
'Public Sub Load_WorkArea(ByRef cboWorkArea As Object)
'
''    Dim I%
''    Dim SqlStmt As String
''    Dim tmpRs As RECORDSET
''
''
''    Set tmpRs = OpenRecordSet(OraDS.GetWorkArea)
''
''    If tmpRs.EOF = True Then ' record가 존재하지 않을 경우
''        tmpRs.RsClose
''        Set tmpRs = Nothing
''        Exit Sub
''    End If
''
''    cboWorkArea.Clear
''    With tmpRs
''        .MoveFirst
'''        For i = 1 To .RecordCount
''        While Not .EOF
''            cboWorkArea.AddItem "" & .Fields("cdval1").Value & "  " & "" & .Fields("field1").Value
''            .MoveNext
''        Wend
'''        Next i
''    End With
''
''    tmpRs.RsClose
''    Set tmpRs = Nothing
'
'End Sub
'
'Public Sub LoadWorkSheetCode(ByVal pBuildingCd As String, ByRef lstWSCode As Object)
'
''    Dim rsWSCode As Object
''    Dim sSqlGetWSCode As String
''    Dim I%
''
''    Set rsWSCode = OpenRecordSet(OraDS.GetWorksheetCode(pBuildingCd))
''
''    If rsWSCode.EOF = True Then
''        MsgBox " worksheet code가 존재하지 않습니다."
''        GoTo NoData
''    End If
''
''    With lstWSCode
''        rsWSCode.MoveFirst
'''        For i = 0 To rsWSCode.RecordCount - 1
''        While Not rsWSCode.EOF
''            .AddItem "" & rsWSCode.Fields("WorkCd").Value & vbTab & _
''                     "" & rsWSCode.Fields("WorkNm").Value & vbTab & _
''                     "" & rsWSCode.Fields("WorkArea").Value & vbTab, I
''                     I = I + 1
''            rsWSCode.MoveNext
''        Wend
'''        Next i
''    End With
''
''NoData:
''    rsWSCode.RsClose
''    Set rsWSCode = Nothing
'
'End Sub
'
'Public Function GetLabNoBySpcNo(ByVal pSpcYy As String, ByVal pSpcNo As String) As String
'
''    Dim RS As RECORDSET
''    Dim SqlStmt As String
''    Dim strMsk As String
''
''    Set RS = OpenRecordSet(OraDS.GetLabNo(pSpcYy, pSpcNo))
''
''    GetLabNoBySpcNo = ""
''    If Not RS.EOF Then
''        GetLabNoBySpcNo = "" & RS.Fields("workarea").Value & "-" & _
''                          "" & RS.Fields("accdt").Value & "-" & _
''                          "" & RS.Fields("accseq").Value
''    End If
''
''    RS.RsClose
''    Set RS = Nothing
'
'End Function
'
'
'Public Sub LoadLstEQCode(ByVal pBuildingCd As String, ByRef lstEQCode As Object)
'
''    Dim rsEQCode As RECORDSET
''    Dim sSqlGetEQCode As String
''    Dim I%
''
''    Set rsEQCode = OpenRecordSet(OraDS.GetEqpList(pBuildingCd))
''
''    If rsEQCode.EOF = True Then
''        MsgBox " 장비내역이 존재하지 않습니다.", vbExclamation, "메세지"
''        GoTo NoData
''    End If
''
''    lstEQCode.Clear
''    With lstEQCode
''        rsEQCode.MoveFirst
'''        For i = 0 To rsEQCode.RecordCount - 1
''        While Not rsEQCode.EOF
''            .AddItem "" & rsEQCode.Fields("EqpCd").Value & vbTab & _
''                     "" & rsEQCode.Fields("EqpNm").Value & vbTab, I
''            I = I + 1
''            rsEQCode.MoveNext
''        Wend
'''        Next i
''    End With
''
''NoData:
''    rsEQCode.RsClose
''    Set rsEQCode = Nothing
'
'End Sub
'
'Public Function GetMaxWorkSeq(ByVal pWorkCd As String, ByRef pWorkDt As String) As Long
''
''    Dim tmpRs As New RECORDSET
''
''    Set tmpRs = OpenRecordSet(OraDS.GetMaxSeq(pWorkCd, pWorkDt))
''
''    If Not tmpRs.EOF Then
''        GetMaxWorkSeq = Val("" & tmpRs.Fields("MaxSeq").Value)
''    Else
''        GetMaxWorkSeq = 0
''    End If
''
''    tmpRs.RsClose
''    Set tmpRs = Nothing
''
'End Function
'
'Public Function GetEqpName(ByVal pEqpCd As String) As String
'
''    GetEqpName = OraDS.GetEqpNm(pEqpCd)
''
'End Function
'
'Public Function IsWorkSeqExists(ByVal pWorkCd As String, ByVal pWorkDt As String, _
'                                ByVal pWorkSeq As String) As Boolean
'
''    IsWorkSeqExists = OraDS.IsWorkSeqExists(pWorkCd, pWorkDt, pWorkSeq)
'
'End Function
'
'
'Public Sub LoadTestCode(ByRef lstList As Object, ByVal pWorkCd As String, ByVal pCase As String)
'
''    Dim rsTestCd As RECORDSET
''
''    Set rsTestCd = OpenRecordSet(OraDS.GetWorksheetTestCd(pWorkCd, pCase))
''
''    lstList.Clear
''    If rsTestCd.EOF Then GoTo NoData
''
''    While (Not rsTestCd.EOF)
''        lstList.AddItem "" & rsTestCd.Fields("TestCd").Value & vbTab & "" & rsTestCd.Fields("TestNm").Value
''        rsTestCd.MoveNext
''    Wend
''
''NoData:
''    rsTestCd.RsClose
''    Set rsTestCd = Nothing
'
'End Sub
'
'
'
'Public Function EqpHistoryDelete(ByVal EqpCd As String, ByVal TransDt As String) As Boolean
''2001-11-07 추가 : 기존 장비전송내역 삭제 (기간 : 1개월)
'
''    Dim SSQL As String
''
''On Error GoTo DELETE_ERROR
''    SSQL = " delete FROM " & T_LAB306 & _
''           " WHERE " & DBW("eqpcd=", EqpCd) & _
''           " AND " & DBW("transdt<=", TransDt)
''
''    DBConn.BeginTrans   '2001-11-07 추가
''    DBConn.Execute SSQL
''    EqpHistoryDelete = True
''    DBConn.CommitTrans
''    Exit Function
''
''DELETE_ERROR:
''    DBConn.RollbackTrans
''    DBConn.DisplayErrors
''
'End Function
'
'
''2002-01-04 추가 : 계산공식 적용
'Public Sub CalculateResult(ByVal Row As Long, ByVal sType As String, _
'                           ByVal dblTotVol As Double, ByVal dblSerumCrea As Double, ByVal dblUrineCrea As Double, _
'                           Optional ByVal dblPara1 As Double, Optional ByVal dblPara2 As Double)
''    Dim Value As String
''    Dim dblCalVal As Double
''
''    ssRst.Row = Row
''    ssRst.Col = SS_COL_RESULT
''    Value = Trim(ssRst.Value)
''    With Me.Result.Item(Row)
''        SELECT Case sType
''        'Urine Protain
''        Case "1":
''            dblCalVal = Val(Value) * 0.01 * dblTotVol * 0.001
''        'GFR/UCR
''        Case "2", "4":
''            dblCalVal = (dblUrineCrea / dblSerumCrea) * (dblTotVol / 1440)
''        End SELECT
''
''        If .AvalVal = 9 Then
''            ssRst.Value = dblCalVal
''        ElseIf .AvalVal = 0 Then
''            ssRst.Value = CLng(dblCalVal)
''        Else
''            ssRst.Value = Format(dblCalVal, "0." & String(.AvalVal, "0"))
''        End If
''        '실제 저장하는 필드...
''        ssRst.Row = Row
''        ssRst.Col = SS_COL_SCREEN_MAX
''        If .AvalVal = 9 Then
''            ssRst.Value = dblCalVal
''        ElseIf .AvalVal = 0 Then
''            ssRst.Value = CLng(dblCalVal)
''        Else
''            ssRst.Value = Format(dblCalVal, "0." & String(.AvalVal, "0"))
''        End If
''    End With
''    ssRst.Row = Row
''    ssRst.Col = SS_COL_RESULT
'End Sub
'
'Public Function DiffCountSQL()
''    DiffCountSQL = "SELECT cdval1 FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", LC3_WBCDiffCode)
'End Function
'
'Public Function GetRstCd(ByVal testcd As String, ByVal RstCd As String) As String
'    Dim RS As RECORDSET
'    Dim sSQL As String
'
'    sSQL = " SELECT a.cdval2,a.field1" & _
'                     " FROM   " & T_LAB031 & " a " & _
'                     " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
'                     " AND    " & DBW("a.cdval1", testcd, 2) & _
'                     " AND    " & DBW("a.cdval2", RstCd, 2) & _
'                     " ORDER BY a.cdval1"
'
'    Set RS = OpenRecordSet(sSQL)
'
'    If Not RS.EOF Then
'        GetRstCd = RS.Fields("field1").Value & ""
'    Else
'        GetRstCd = RstCd
'    End If
'    Set RS = Nothing
'
'
'End Function
'Public Function GetRstChk(ByVal testcd As String, ByVal RstCd As String) As Boolean
''    Dim RS As RECORDSET
''    Dim SSQL As String
''
''    SSQL = " SELECT a.cdval2,a.field1" & _
''                     " FROM   " & T_LAB031 & " a " & _
''                     " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
''                     " AND    " & DBW("a.cdval1", TestCd, 2) & _
''                     " AND    " & DBW("a.cdval2", RstCd, 2) & _
''                     " ORDER BY a.cdval1"
''
''    Set RS = OpenRecordSet(SSQL)
''
''    If Not RS.EOF Then
''        GetRstChk = True
''    Else
''        GetRstChk = False
''    End If
''    Set RS = Nothing
'End Function
'
'
'
'
'Public Sub PopUp(Optional ColorSet As Boolean = True, Optional ByVal iCol As Integer = 0)
'
'    Dim lngBkBackColor  As Long
'    Dim blnRstCd        As Boolean
'    Dim blnEqpCd        As Boolean
'    Dim sTestCD         As String
'
'
'
'    If ColorSet = True Then
'        With ssRst
'            .Row = ssRst.ActiveRow
'            .Col = iCol - 1
'            sTestCD = medGetP(Trim(.Value), 1, COL_DIV)
'            .Col = -1
'            lngBkBackColor = .BackColor
'            .BackColor = vb3DLight
'            DoEvents
'        End With
'    End If
'
'    Set mnuPopup = frmMenu.mnuPopup
'
'    blnRstCd = IsRstCode(sTestCD)
'    Set mnuUtil1 = frmMenu.mnuUtil1
'    mnuUtil1.Caption = "SELECT Result Code"
'
'    If blnRstCd = False Then
'        mnuUtil1.Enabled = False
'    Else
'        mnuUtil1.Enabled = True
'    End If
'
'    If mnuUtil1.Enabled Then
'        'Call mnuUtil1_Click
'        Call ItemRstCodePop(iCol, sTestCD)
'    End If
'
'
''    If iCol = Me.SSCol("RESULT") Then
''        If mnuUtil1.Enabled Then Call mnuUtil1_Click
''    ElseIf iCol = Me.SSCol("EQPNM") Then
''        If mnuUtil2.Enabled Then Call mnuUtil2_Click
''    Else
''        OpenContextMenu frmMenu, mnuPopup
''    End If
'    '
'    Set mnuPopup = Nothing
'    Set mnuUtil1 = Nothing
''    Set mnuUtil2 = Nothing
'    '
'    If ColorSet = True Then
'        With ssRst
'            .Row = ssRst.ActiveRow
'            .Col = -1
'            .BackColor = lngBkBackColor
'        End With
'    End If
'    '
'End Sub
'
'Public Function SetRST(ByRef objRST As Object, ByRef objDic1 As clsDictionary)
'    Set ssRst = objRST
'    Set objdic = objDic1
'End Function
'
'Public Function IsRstCode(ByVal Item As String) As Boolean
'    Dim strSql  As String
'    Dim RS      As RECORDSET
'
'    IsRstCode = False
'
'    strSql = " SELECT COUNT(a.cdval2) AS ITEMCNT" & _
'             " FROM   " & T_LAB031 & " a " & _
'             " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
'             " AND    " & DBW("a.cdval1", Item, 2)
'
'    Set RS = OpenRecordSet(strSql)
'    If Not RS.EOF Then
'        If Val("" & RS.Fields("ITEMCNT").Value) > 0 Then
'           IsRstCode = True
'        End If
'    End If
'    Set RS = Nothing
'
'End Function
'
'Public Function SqlStmtRstCode(ByVal Item As String) As String
'    SqlStmtRstCode = " SELECT a.cdval2,a.field1" & _
'                     " FROM   " & T_LAB031 & " a " & _
'                     " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
'                     " AND    " & DBW("a.cdval1", Item, 2) & _
'                     " ORDER BY a.cdval1"
'End Function
'
'Public Function GetRstCdSQL(ByVal cdindex As String, ByVal CdVal1 As String, _
'                         ByVal CdVal2 As String) As RECORDSET
'    Dim SqlStmt As String
'   '
'    SqlStmt = " SELECT * FROM " & T_LAB031 & _
'              " WHERE " & DBW("cdindex = ", cdindex) & _
'              " AND   " & DBW("cdval1  = ", CdVal1) & _
'              " AND   " & DBW("cdval2  = ", CdVal2)
'    Set GetRstCdSQL = OpenRecordSet(SqlStmt)
'
'   '
'End Function
'Public Function UpdatableLAB031(ByVal cdindex As String, ByVal CdVal1 As String, _
'                                ByVal CdVal2 As String) As Boolean
'    Dim SqlStmt As String
'    Dim RS      As RECORDSET
'
'    SqlStmt = " SELECT * FROM " & T_LAB031 & _
'              " WHERE " & DBW("cdindex = ", cdindex) & _
'              " AND   " & DBW("cdval1  = ", CdVal1) & _
'              " AND   " & DBW("cdval2  = ", CdVal2)
'
'    Set RS = OpenRecordSet(SqlStmt)
'    If RS.RecordCount > 0 Then
'        UpdatableLAB031 = True
'    Else
'        UpdatableLAB031 = False
'    End If
'
'    Set RS = Nothing
'End Function
Private Sub Class_Initialize()

End Sub
