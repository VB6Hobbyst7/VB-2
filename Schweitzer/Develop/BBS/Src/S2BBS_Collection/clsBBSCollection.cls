VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'병원별 처방상태 관리 =======KJG(2001/12/12)
Private stsORDER    As String
Private stsCOLLECT  As String
Private stsACCESS   As String
Private sWardId     As String   '일괄채혈병동
Private sWorkDt     As String   '일괄채혈일자
Private sWorkTm     As String   '일괄채혈시간
Private mvarWardID  As String   '병동코드(검체내역, BBS102)
Private mvarSpc     As String
Private iWorkSeq    As Long    '일괄채혈Seq
Private blnSpc72Chk As Boolean
Private mvarBldDic As New clsDictionary
Private mvarCheckCol As Boolean '채혈처리가 유효한지 여부 체크

Public Property Let Wardid(ByVal vData As String)
    mvarWardID = vData
End Property

Public Property Get SpcNum() As String
    SpcNum = mvarSpc
End Property

Public Property Get WardColSeq() As Long
    WardColSeq = iWorkSeq
End Property

Public Property Get BldDic() As clsDictionary
    Set BldDic = mvarBldDic
End Property

Public Property Get Spc72Chk() As Boolean
    Spc72Chk = blnSpc72Chk
End Property

Public Property Get CheckCol() As Boolean
    CheckCol = mvarCheckCol
End Property

Public Sub SetWardCol(ByVal pWardId As String, ByVal pWorkDt As String, ByVal pWorkTm As String)

    Dim RS      As Recordset
    Dim strSql  As String
    
    
'    blnBatchCol = True    '병동별 일괄채혈 구분
    
    sWardId = pWardId
    sWorkDt = pWorkDt   'Format(Now, CS_DateDbFormat)
    sWorkTm = pWorkTm   'Format(Now, CS_TimeDbFormat)
    
    strSql = " SELECT max(seq) as maxseq FROM " & T_LAB204 & _
             " WHERE " & _
                                 DBW("workdt", sWorkDt, 2) & " " & _
                       " AND " & DBW("wardid", pWardId, 2) & _
                       " AND " & DBW("worktm", sWorkTm, 2)
    Set RS = New Recordset
    RS.Open strSql, DBConn
    iWorkSeq = Val("" & RS.Fields("maxseq").Value & "")
    Set RS = Nothing
End Sub

Public Function Blood_Existence(ByVal ptid As String, ByVal coldt As String, ByVal coltm As String) As Boolean
    '검체의 존재여부를 확인한다.(true : 없다. flase : 있다)
    Dim RS              As Recordset
    Dim strSql          As String
    Dim strExistence    As String
    Dim strCompare      As String
    Dim strCompare1     As String
    Dim lngStoreHour    As Long
    Dim KeepTime        As Long            '검체보관시간...
    
    KeepTime = GetKeepHour
    strCompare1 = coldt & Mid(coltm, 1, 4)
    strCompare1 = Format(strCompare1, "####-##-## ##:##")
    strSql = " SELECT coldt,coltm" & _
             " FROM " & T_BBS201 & " a" & _
             " WHERE" & _
                      " a.coldt=(SELECT max(b.coldt) " & _
                                " FROM " & T_BBS201 & " b" & _
                                " WHERE " & DBW("b.ptid", ptid, 2) & "" & _
                                " AND a.ptid=b.ptid " & _
                                " AND " & DBW("b.coldt<", coldt, 2) & ") " & _
             " ORDER BY a.coltm desc "
    Set RS = New Recordset
    RS.Open strSql, DBConn
    
    If RS.EOF Then
        Blood_Existence = True 'true로 해줘야 없는걸로 되쥐.. 이걸 왜 빼먹은건지 ㅡㅡㅋ
'        dbconn.DisplayErrors
        Set RS = Nothing
        Exit Function
    End If
    If Not RS.EOF = True Then
        strCompare = RS.Fields("coldt").Value & "" & Mid(RS.Fields("coltm").Value & "", 1, 4)
        strCompare = Format(strCompare, "####-##-## ##:##")
        lngStoreHour = CLng(DateDiff("h", strCompare, strCompare1))
        If lngStoreHour <= KeepTime Then
             Blood_Existence = False
        Else
            Blood_Existence = True
        End If
    Else
        Blood_Existence = True
    End If
    Set RS = Nothing
End Function

Private Function GetKeepHour() As Long
    Dim objcom003   As clsCom003
    Dim RS          As Recordset
    
    Set objcom003 = New clsCom003
    Set RS = objcom003.OpenRecordSetDay(BC2_KEEP_HOUR)
    If RS Is Nothing Then
        GetKeepHour = 0
    End If
    With RS
        If .RecordCount < 1 Then
            GetKeepHour = 0
        Else
            GetKeepHour = .Fields("field1").Value & ""
        End If
    End With
    Set RS = Nothing
    Set objcom003 = Nothing
End Function

Private Function Collect(ByVal ptid As String, ByVal Bussdiv As String, ByVal colid As String, _
                         ByVal coldt As String, ByVal coltm As String, ByVal rcvid As String, _
                         ByVal rcvdt As String, ByVal rcvtm As String, _
                         ByVal statfg As String, ByVal build As String, _
                         ByVal savepos As String, ByVal Leg As String, _
                         ByVal Row As String, ByVal Col As String) As Boolean
    '----------------------------------------------------------------------------
    '처방을 읽으면서, 채혈처리를 한다.
    '----------------------------------------------------------------------------
    Dim objNumbers  As clsBBSNumbers
    Dim RS          As Recordset
    Dim sSql        As String
    Dim orgbuildcd  As String
    Dim TestBuild   As String
    Dim spcyy       As String
    Dim spcno       As Long
    Dim i           As Long
    
On Error GoTo CollectProcess_error
    '----------------------------------------------------------------------------
    '만일 이 환자에게 검체가 존재하면 채혈하지 않는다.
    '----------------------------------------------------------------------------
    If Blood_Existence(ptid, coldt, coltm) = False Then
        Collect = True
        Exit Function
    End If
    '----------------------------------------------------------------------------
    '채혈되지 않은 처방이 있고, 그 처방의 검사방법이 None이 아니면 채혈을 한다.
    '----------------------------------------------------------------------------
    sSql = " SELECT b.* " & _
           " FROM " & T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c " & _
           " WHERE " & DBW("a.orddiv", C_WORKAREA, 2) & _
           " AND " & DBW("a.ptid", ptid, 2) & _
           " AND " & DBW("a.donefg", stsORDER, 2) & _
           " AND " & DBW("a.bussdiv", Bussdiv, 2) & _
           " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
           " AND " & DBW("b.stscd", stsORDER, 2) & _
           " AND (b.dcfg is null or " & DBW("b.dcfg", "", 2) & ") " & _
           " AND b.ordcd=c.testcd " & _
           " AND c.applydt=( SELECT max(d.applydt) " & _
                           " FROM " & T_BBS001 & " d " & _
                           " WHERE d.testcd = b.ordcd " & _
                           " AND d.applydt<=b.orddt " & _
                           ") " & _
           " AND " & DBW("c.xmethod<>", XMethod.xmNONE)
           
    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    With RS
        If .EOF Then GoTo CollectProcess_error
        If .RecordCount <= 0 Then
            '채혈대상이 없다.----------------------------------------------------
            Collect = True
            Set RS = Nothing
            Exit Function
        End If
    End With
    Set RS = Nothing

    '--------------------------------------------------------------------------------------------
    '보관장소지정
    '검사건물 지정
    '--------------------------------------------------------------------------------------------
    If GetSpecSpace(build, savepos, Leg, Row, Col) = False Then Exit Function
    If GetOrgDestBuilding(Bussdiv, build, statfg, orgbuildcd, TestBuild) = False Then Exit Function

    '----------------------------------------------------------------------------
    '검체번호를 구한다.
    '----------------------------------------------------------------------------
     
    Set objNumbers = New clsBBSNumbers
    With objNumbers
        spcyy = medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 1, COL_DIV)
        spcno = Val(medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 2, COL_DIV))
    End With
    '----------------------------------------------------------------------------
    '채혈내역 생성
    '----------------------------------------------------------------------------
    sSql = Set_Collect201(spcyy, spcno, coldt, coltm, colid, rcvdt, rcvtm, rcvid, Leg, Row, Col, _
                          ptid, TestBuild, build, BBSSpcAddFg.stsORDER)
    DBConn.Execute sSql
    
    '----------------------------------------------------------------------------
    '보관장소 update
    '----------------------------------------------------------------------------
    sSql = Set_Position(build, Leg, Row, Col, spcyy, spcno)
    DBConn.Execute sSql
    
    '----------------------------------------------------------------------------
    '검체번호 update
    '----------------------------------------------------------------------------
    sSql = objNumbers.Set_NumbersCom099(BN_SPC_NO, spcyy, spcno)
    DBConn.Execute sSql
    
    '검체번호 저장(바코드 출력을 위해서)
    mvarSpc = spcyy & spcno
    
    Collect = True
    Set objNumbers = Nothing
    Exit Function
    
CollectProcess_error:
    Set objNumbers = Nothing
    Collect = False
End Function

Public Function Set_OutPatientCollect(ByVal objdic As clsDictionary) As Boolean
    Dim ptid    As String
    Dim Bussdiv As String
    Dim buildcd As String
    Dim ErChk   As String
    Dim colid   As String
    Dim coldt   As String
    Dim coltm   As String
    Dim rcvid   As String
    Dim rcvtm   As String
    Dim rcvdt   As String
    Dim SaveChk As String
    Dim Leg     As String
    Dim Row     As String
    Dim Col     As String
    
'순서(추가채혈 여부를 확인해서, 있으면, 추가채혈을 한다.)
'채혈여부를 판단해서 채혈을 한후 Update 한다..
'처방을 update update 해준다.
    
    DBConn.BeginTrans

On Error GoTo Set_Order_error
    If objdic.RecordCount < 1 Then Exit Function
    
    With objdic
        .MoveFirst
        ptid = .Fields("ptid")
        buildcd = .Fields("Buildcd")
        ErChk = IIf(.Fields("Erchk") = "Ok", "1", "")
        colid = .Fields("Colid")
        coldt = .Fields("ColDt")
        coltm = .Fields("ColTm")
        rcvid = .Fields("rcvid")
        rcvtm = .Fields("rcvtm")
        rcvdt = .Fields("rcvdt")
        SaveChk = IIf(.Fields("SaveChk") = "Ok", "1", "")
        Leg = .Fields("Leg")
        Row = .Fields("Row")
        Col = .Fields("Col")
        Bussdiv = BBSBUSSDIV.stsNotBed
    End With
'-------------------------------------------------
'1번:추가채혈여부를 판단해서 추가채혈 수행
'2번:채혈작업수행(추가채혈이 되었다면, 건너뛴다.)
'3번:처방에 관련되 테이블 Update 해준다.
'-------------------------------------------------
    '1번
    If Addspc_Exist(ptid, Bussdiv, colid, coldt, coltm, rcvid, rcvdt, rcvtm, ErChk, buildcd, SaveChk, Leg, Row, Col) = False Then
        GoTo Set_Order_error
    End If
    '2번
    If Collect(ptid, Bussdiv, colid, coldt, coltm, rcvid, rcvdt, rcvtm, ErChk, buildcd, SaveChk, Leg, Row, Col) = False Then
        GoTo Set_Order_error
    End If
    '3번
    If OrderAccess(ptid, Bussdiv, rcvdt, rcvtm, rcvid) = False Then
        GoTo Set_Order_error
    End If
                                      
    DBConn.CommitTrans
    Set_OutPatientCollect = True
    Exit Function
    
Set_Order_error:
    DBConn.RollbackTrans
    Set_OutPatientCollect = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Set_Collect(objdic As clsDictionary, Optional ByRef pToBuildCd As Variant, _
                            Optional ByRef objProgress As Object = Nothing, _
                            Optional ByVal blnWardCol As Boolean = False) As Boolean
    Dim objNumbers As clsBBSNumbers
    Dim blnOrder As Boolean       '처방에 의한 채혈대상
    Dim blnAdd   As Boolean       '추가채혈에 의한 채혈대상
    Dim sSpcYY   As String        '검체번호(년도)
    Dim sSpcno   As Long          '검체번호
    Dim Sptid    As String        '환자id
    Dim Sptnm    As String        '환자명
    Dim Scolid   As String        '채혈자id
    Dim Scoldt   As String        '채혈일
    Dim Scoltm   As String        '채혈시간
    Dim Tbuildcd As String        '검사수행될 건물코드
    Dim Cbuildcd As String        '채혈수행될 건물코드
    Dim Bussdiv  As String        '업무구분(입원외래)
    Dim G_year   As String        '검체번호를 따오기 위해서 사용
    Dim BarSpcno As String        '바코드 검체번호
    Dim HosilID  As String
    Dim strStatfg As String        '응급여부
    
    Dim OldSpcYy As String
    Dim OldSpcNo As String
    Dim OldColDt As String
    Dim OldColTm As String
    Dim OldColId As String
'sql문장 변수
    Dim strTmp   As String
    Dim strTmp1  As String
    Dim sSql     As String
    Dim i        As Long
    
    G_year = BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4, 1)
    Set objNumbers = New clsBBSNumbers
    
    OldSpcYy = ""
    OldSpcNo = ""
    blnSpc72Chk = False
    mvarCheckCol = False
    
    With objNumbers
        sSpcYY = medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 1, COL_DIV)
        sSpcno = Val(medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 2, COL_DIV))
    End With
    
    '--------------------------------
    '바코드 찍기 위해 Dictionary 선언
    '--------------------------------
    mvarBldDic.Clear
    mvarBldDic.FieldInialize "ptid", "ptnm,spcno,coldt,coltm,hosilid,rcvfg"
    
On Error GoTo Set_Collect_error
    
    DBConn.BeginTrans
    
    objdic.MoveFirst
    Do Until objdic.EOF
        Scoldt = objdic.Fields("coldt")
        Scoltm = objdic.Fields("coltm")
        Scolid = objdic.Fields("colid")
        Sptid = objdic.Fields("ptid")
        Sptnm = objdic.Fields("ptnm")
        Bussdiv = objdic.Fields("bussdiv")
        Cbuildcd = objdic.Fields("buildcd")
        HosilID = objdic.Fields("hosilid")
        strStatfg = objdic.Fields("statfg")
        '-------------------------------------------------
        '응급여부를 판단해서 실제 검사할 건물코드를 구한다
        '검사건물 코드가 없으면, 채혈처리를 하지 않는다
        '-------------------------------------------------
        Tbuildcd = ObjSysInfo.buildingcd
        If Not IsMissing(pToBuildCd) Then pToBuildCd = Tbuildcd    '2001.2.9 추가

        If Tbuildcd = "" Then
            MsgBox "데이터오류입니다. 검사를 수행할 건물정보가 누락되었습니다."
            GoTo Set_Collect_error
            Exit Function
        End If
        '----------------------------------------------------------
        '채혈의 종류를 구한다.(blnOrder:처방채혈,blnAdd:추가채혈요청
        '----------------------------------------------------------
        If Collect_Target(Sptid, Bussdiv) = True Then
            blnOrder = True
        End If
        If Addspc_Exist2(Sptid, Bussdiv) = True Then
            blnAdd = True
        End If
        '--------------------------------------------------------------------------------------
        'blnOrder = True AND blnAdd = False(처방에의한 채혈)
        'Spc_Exist(검체가 이미 존재하는지 Check(True이면, 채혈내역작성))
        '처방헤더,바디 Update
        '병동일괄채혈일경우(blnWardCol:True) lab204생성(Set_Collect204) : 2001.2.16 김미경 추가
        '--------------------------------------------------------------------------------------
        If blnOrder = True And blnAdd = False Then
            If Blood_Existence(Sptid, Scoldt, Scoltm) = True Then '검체 없는 경우
                sSql = Set_Collect201(sSpcYY, CStr(sSpcno), Scoldt, Scoltm, Scolid, "", "", "", "", "", "", Sptid, Tbuildcd, Cbuildcd, BBSSpcAddFg.stsORDER)
                DBConn.Execute sSql
                If blnWardCol Then
                    iWorkSeq = iWorkSeq + 1
                    sSql = Set_Collect204(sWorkDt, sWardId, sWorkTm, iWorkSeq, Scolid, Tbuildcd, Cbuildcd, sSpcYY, CStr(sSpcno), C_WORKAREA)
                    DBConn.Execute sSql
                End If
                
                BarSpcno = sSpcYY & Format(sSpcno, BC2_BARCODE)
                mvarBldDic.AddNew Sptid, Join(Array(Sptnm, BarSpcno, Scoldt, Scoltm, HosilID, strStatfg), COL_DIV)
                '-------------------------------
                '채혈내역 & 처방내역 연결 테이블
                '-------------------------------
                strTmp1 = Set_SpcAndOrddt(Sptid, Bussdiv, sSpcYY, CStr(sSpcno), Scoldt, Scoltm, Scolid)
                If strTmp1 <> "" Then
                    i = 0
                    Do
                        i = i + 1
                        sSql = medGetP(strTmp1, i, COL_DIV)
                        If sSql = "" Then Exit Do
                        
                        DBConn.Execute sSql
                    Loop
                End If
                sSpcno = CLng(sSpcno) + 1
            Else '이미 검체가 있는 경우
                '------------------------------------------------------------
                ' 72시간이 지나지 않은 검체에 대해서는 바코드를
                ' 출력하고, 채혈내역&처방내역 연결테이블을 생성시켜준다.
                '------------------------------------------------------------
                If Get_Collect201(OldSpcYy, OldSpcNo, OldColDt, OldColTm, OldColId, Sptid) = True Then
                    '검체가 존재할경우 채혈 바코드 출력여부를 결정해야한다.
                    BarSpcno = OldSpcYy & Format(OldSpcNo, BC2_BARCODE)
                    mvarBldDic.AddNew Sptid, Join(Array(Sptnm, BarSpcno, OldColDt, OldColTm, HosilID, strStatfg), COL_DIV)
                    '-------------------------------
                    '채혈내역 & 처방내역 연결 테이블
                    '-------------------------------
                    strTmp1 = Set_SpcAndOrddt(Sptid, Bussdiv, OldSpcYy, OldSpcNo, OldColDt, OldColTm, OldColId)
                    If strTmp1 <> "" Then
                        i = 0
                        Do
                            i = i + 1
                            sSql = medGetP(strTmp1, i, COL_DIV)
                            If sSql = "" Then Exit Do
                            
                            DBConn.Execute sSql
                        Loop
                    End If
                End If
                blnSpc72Chk = True
            End If
            'OCS에 상태 전송이 포함되어 있음
            strTmp = Set_CollectOrder(Sptid, Bussdiv)
            i = 0
            Do
                i = i + 1
                sSql = medGetP(strTmp, i, COL_DIV)
                If sSql = "" Then Exit Do
                DBConn.Execute sSql
            Loop
        End If
        '--------------------------------------------------------------------------------------
        'blnOrder = False AND blnAdd = True(추가처방에의한 채혈) 무조건채혈
        '채혈내역작성,추가요청내역 Update
        '병동일괄채혈일경우(blnWardCol:True) lab204생성(Set_Collect204) : 2001.2.16 김미경 추가
        '--------------------------------------------------------------------------------------
        If blnOrder = False And blnAdd = True Then
            sSql = Set_Collect201(sSpcYY, CStr(sSpcno), Scoldt, Scoltm, Scolid, "", "", "", "", "", "", Sptid, Tbuildcd, Cbuildcd, BBSSpcAddFg.stsCOLLECT)
            DBConn.Execute sSql
            
            If blnWardCol Then
                iWorkSeq = iWorkSeq + 1
                sSql = Set_Collect204(sWorkDt, sWardId, sWorkTm, iWorkSeq, Scolid, Tbuildcd, Cbuildcd, sSpcYY, CStr(sSpcno), C_WORKAREA)
                DBConn.Execute sSql
            End If
            
            '추가요청내역 Up
            sSql = Set_BBS207(Sptid, Bussdiv, sSpcYY, sSpcno)
            DBConn.Execute sSql
            
            BarSpcno = sSpcYY & Format(sSpcno, BC2_BARCODE)
            mvarBldDic.AddNew Sptid, Join(Array(Sptnm, BarSpcno, Scoldt, Scoltm, HosilID, strStatfg), COL_DIV)
            sSpcno = CLng(sSpcno) + 1
        End If
        '---------------------------------------------------------------------------------------
        'blnOrder = True AND blnAdd = True(처방채혈,추가요청 모두존재) 무조건채혈
        '채혈내역작성,처방헤더,바디 Update
        '추가요청내역 Update
        '병동일괄채혈일 경우(blnWardCol:True) lab204생성(Set_Collect204) : 2001.2.16 김미경 추가
        '---------------------------------------------------------------------------------------
        If blnOrder = True And blnAdd = True Then
            sSql = Set_Collect201(sSpcYY, CStr(sSpcno), Scoldt, Scoltm, Scolid, "", "", "", "", "", "", Sptid, Tbuildcd, Cbuildcd, BBSSpcAddFg.stsORDER)
            DBConn.Execute sSql
            
            If blnWardCol Then
                iWorkSeq = iWorkSeq + 1
                sSql = Set_Collect204(sWorkDt, sWardId, sWorkTm, iWorkSeq, Scolid, Tbuildcd, Cbuildcd, sSpcYY, CStr(sSpcno), C_WORKAREA)
                DBConn.Execute sSql
            End If
            '처방연결
            strTmp1 = Set_SpcAndOrddt(Sptid, Bussdiv, sSpcYY, CStr(sSpcno), Scoldt, Scoltm, Scolid)
            If strTmp1 <> "" Then
                i = 0
                Do
                    i = i + 1
                    sSql = medGetP(strTmp1, i, COL_DIV)
                    If sSql = "" Then Exit Do
                    
                    DBConn.Execute sSql
                Loop
            End If
            '헤더바디 업데이트
            'OCS에 상태 전송이 포함되어 있음
            strTmp = Set_CollectOrder(Sptid, Bussdiv)
            i = 0
            Do
                i = i + 1
                sSql = medGetP(strTmp, i, COL_DIV)
                If sSql = "" Then Exit Do
                
                DBConn.Execute sSql
            Loop
            
            sSql = Set_BBS207(Sptid, Bussdiv, sSpcYY, sSpcno)
            DBConn.Execute sSql
            
            BarSpcno = sSpcYY & Format(sSpcno, BC2_BARCODE)
            mvarBldDic.AddNew Sptid, Join(Array(Sptnm, BarSpcno, Scoldt, Scoltm, HosilID, strStatfg), COL_DIV)
            sSpcno = CLng(sSpcno) + 1
        End If
        
        If blnOrder = False And blnAdd = False Then
            mvarCheckCol = True
        End If
        
        objdic.MoveNext
    Loop
    '-------------------
    '번호부여정보 update
    '-------------------
    sSql = objNumbers.Set_NumbersCom099(BN_SPC_NO, sSpcYY, sSpcno - 1)
    DBConn.Execute sSql
    
    DBConn.CommitTrans
    Set_Collect = True
    Set objNumbers = Nothing
    Exit Function
Set_Collect_error:
    MsgBox Err.Description
    DBConn.RollbackTrans
    Set_Collect = False
    Set objNumbers = Nothing
End Function

Private Function GetSpecSpace(ByVal build As String, ByVal savepos As String, _
                               Leg As String, Row As String, Col As String) As Boolean
 
    '--------------------------------------------------------------------------------------------
    '보관장소를 구하거나, 입력된 보관장소가 사용가능한지 검사한다.
    '--------------------------------------------------------------------------------------------
    'savepos = "1"  : 자동지정
    '        = "0"  : 수동지정
    '--------------------------------------------------------------------------------------------
    '반환값 = True  : 보관장소를 사용할 수 있다.
    '       = False : 보관장소를 사용할 수 없다.
    '--------------------------------------------------------------------------------------------
    Dim objSpec As clsSpecManagement
    
    
    '수동 지정일 경우 입력값 확인----------------------------------------------------------------
    If savepos = "0" Then
        If Leg = "" Or Row = "" Or Col = "" Then
            Call MsgBox("보관장소 누락", vbCritical + vbOKOnly, "보관장소 누락")
            GetSpecSpace = False
            Exit Function
        End If
    End If
    '메모리 할당---------------------------------------------------------------------------------
    Set objSpec = New clsSpecManagement
    '자동지정일 경우 보관장소를 구한다.----------------------------------------------------------
    If savepos = "1" Then
        objSpec.Save_Spc_Search 1, build, Leg
        Leg = objSpec.Leg(1)
        Row = objSpec.Row(1)
        Col = objSpec.Col(1)
    End If
    '선택된(구해진) 보관장소가 사용가능한지 검사-------------------------------------------------
    If objSpec.SavePointChk(Leg, Row, Col, build) = False Then
        MsgBox "이미 사용중인 보관장소이거나,사용할수 없는 보관장소입니다." & vbNewLine & _
               "확인후 진행하십시요.", vbCritical + vbOKOnly, "보관장소입력에러"
        GetSpecSpace = False
    Else
        GetSpecSpace = True
    End If
    '메모리 해제---------------------------------------------------------------------------------
    Set objSpec = Nothing
End Function

Private Function Addspc_Exist(ByVal ptid As String, ByVal Bussdiv As String, ByVal colid As String, _
                         ByVal coldt As String, ByVal coltm As String, ByVal rcvid As String, _
                         ByVal rcvdt As String, ByVal rcvtm As String, _
                         ByVal statfg As String, ByVal build As String, _
                         ByVal savepos As String, ByVal Leg As String, _
                         ByVal Row As String, ByVal Col As String) As Boolean
    
    Dim RS          As Recordset
    Dim sSql        As String
    Dim orgbuildcd  As String
    Dim TestBuild   As String
    Dim spcyy       As String
    Dim spcno       As Long
    
On Error GoTo CollectProcess_error

    sSql = " SELECT * FROM " & T_BBS207 & _
           " WHERE " & _
                   "     " & DBW("ptid", ptid, 2) & _
                   " AND " & DBW("busidiv", Bussdiv, 2) & _
                   " AND " & DBW("dongfg", "0", 2)
                   
    Set RS = New Recordset
    RS.Open sSql, DBConn
    If RS.EOF = True Then
        Addspc_Exist = True
        Set RS = Nothing
        Exit Function
    End If
    '----------------------------------
    '추가 요청이 있어서 검체채혈을 한다.
    '----------------------------------
    '--------------------------------------------------------------------------------------------
    '보관장소지정
    '검사건물 지정
    '--------------------------------------------------------------------------------------------
    If GetSpecSpace(build, savepos, Leg, Row, Col) = False Then Exit Function
    If GetOrgDestBuilding(Bussdiv, build, statfg, orgbuildcd, TestBuild) = False Then Exit Function

    '----------------------------------------------------------------------------
    '검체번호를 구한다.
    '----------------------------------------------------------------------------
    Dim objNumbers As clsBBSNumbers
    Set objNumbers = New clsBBSNumbers
    
    With objNumbers
        spcyy = medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 1, COL_DIV)
        spcno = Val(medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 2, COL_DIV))
    End With
    
    '----------------------------------------------------------------------------
    '검체추가요청내역 Update
    '----------------------------------------------------------------------------
    sSql = Set_BBS207(ptid, Bussdiv, spcyy, spcno)
    DBConn.Execute sSql
    
    '----------------------------------------------------------------------------
    '채혈내역 생성
    '----------------------------------------------------------------------------
    sSql = Set_Collect201(spcyy, spcno, coldt, coltm, colid, rcvid, rcvtm, rcvid, Leg, Row, Col, _
                       ptid, TestBuild, build, BBSSpcAddFg.stsCOLLECT)
    DBConn.Execute sSql
    
    
    '----------------------------------------------------------------------------
    '보관장소 update
    '----------------------------------------------------------------------------
    sSql = Set_Position(build, Leg, Row, Col, spcyy, spcno)
    DBConn.Execute sSql
    
    
    '----------------------------------------------------------------------------
    '검체번호 update
    '----------------------------------------------------------------------------
    sSql = objNumbers.Set_NumbersCom099(BN_SPC_NO, spcyy, spcno)
    DBConn.Execute sSql
    
    
    '검체번호 저장(바코드 출력을 위해서)
    mvarSpc = spcyy & spcno
    
    
    Addspc_Exist = True
    Set objNumbers = Nothing
    Exit Function
    
CollectProcess_error:
    Addspc_Exist = False
    Set objNumbers = Nothing
    MsgBox Err.Description, vbExclamation
End Function

Private Function GetOrgDestBuilding(ByVal Bussdiv As String, _
                                    ByVal build As String, ByVal statfg As String, _
                                     orgbuildcd As String, TestBuild As String) As Boolean

    Dim ipos As Integer
    
    ipos = IIf(statfg = "1", 1, 2)
    
    
    
    '외래,병동에 따른 채혈 건물을 구한다.---------------------------------------------------------
    If Bussdiv = "1" Then
        orgbuildcd = build
    End If
    
    '채혈건물에 따른 검사건물을 구한다.-----------------------------------------------------------
    TestBuild = medGetP(Get_TestBuild, ipos, COL_DIV)
     
    '검사건물 정보가 정확하게 있는지 검사한다.----------------------------------------------------
    If TestBuild = "" Then
        Call MsgBox("해당건물정보가 누락되었습니다.", vbCritical + vbOKOnly, "건물정보누락")
        GetOrgDestBuilding = False
    Else
        GetOrgDestBuilding = True
    End If
End Function

Private Function Set_Collect201(ByVal spcyy As String, ByVal spcno As String, _
                                ByVal coldt As String, ByVal coltm As String, _
                                ByVal colid As String, ByVal rcvdt As String, _
                                ByVal rcvtm As String, ByVal rcvid As String, _
                                ByVal Leg As String, ByVal Row As String, ByVal Col As String, _
                                ByVal ptid As String, ByVal buildcd As String, _
                                ByVal orgbuildcd As String, ByVal Addfg As String) As String
        
'--------------------------
'채혈내역 작성(to:h7bbs201)
'--------------------------
    Set_Collect201 = " insert into " & T_BBS201 & "(" & _
                     " spcyy,spcno,coldt,coltm,colid,rcvdt,rcvtm,rcvid,storeleg,storerno,storecno," & _
                     " ptid,buildcd,orgbuildcd,localcd,expfg,addfg)" & _
                     " values(" & _
                               DBV("spcyy", spcyy, 1) & DBV("spcno", spcno, 1) & _
                               DBV("coldt", coldt, 1) & DBV("coltm", coltm, 1) & _
                               DBV("colid", colid, 1) & DBV("rcvdt", rcvdt, 1) & _
                               DBV("rcvtm", rcvtm, 1) & DBV("rcvid", rcvid, 1) & _
                               DBV("storeleg", Leg, 1) & DBV("storerno ", Row, 1) & DBV("storecno", Col, 1) & _
                               DBV("ptid", ptid, 1) & DBV("buildcd", buildcd, 1) & _
                               DBV("orgbuildcd", orgbuildcd, 1) & DBV("localcd     ", "", 1) & _
                               DBV("expfg     ", "0", 1) & DBV("addfg       ", Addfg) & _
                            ")"

End Function


Private Function Get_Collect201(ByRef spcyy As String, ByRef spcno As String, _
                                ByRef coldt As String, ByRef coltm As String, _
                                ByRef colid As String, ByVal ptid As String) As Boolean
        
    Dim strSql      As String
    Dim DrRS        As Recordset
    
    strSql = " SELECT spcyy,spcno,coldt,coltm,colid" & _
             " FROM " & T_BBS201 & _
             " WHERE " & DBW("ptid=", ptid) & _
             " ORDER BY spcyy desc, spcno desc"
    Set DrRS = New Recordset
    Call DrRS.Open(strSql, DBConn)
    If Not DrRS.EOF Then
        spcyy = "" & DrRS.Fields("spcyy").Value & ""
        spcno = "" & DrRS.Fields("spcno").Value & ""
        coldt = "" & DrRS.Fields("coldt").Value & ""
        coltm = "" & DrRS.Fields("coltm").Value & ""
        colid = "" & DrRS.Fields("colid").Value & ""
        Get_Collect201 = True
    Else
        Get_Collect201 = False
    End If
    Set DrRS = Nothing

End Function


Private Function Set_Position(ByVal cntercd As String, ByVal legcd As String, ByVal rowno As Long, _
                              ByVal colno As Long, ByVal spcyy As String, ByVal spcno As Long) As String
'보관장소 update
    Set_Position = " update " & T_BBS206 & " set" & _
                                             DBW("spcyy", spcyy, 3) & _
                                             DBW("spcno", spcno, 3) & _
                                             DBW("stscd", BBSSaveStatue.stsNotUsed, 2) & _
                    " WHERE " & _
                                      DBW("centercd", cntercd, 2) & _
                            " AND " & DBW("legcd", legcd, 2) & _
                            " AND " & DBW("rowno", rowno, 2) & _
                            " AND " & DBW("colno", colno, 2) & ""
End Function
Public Function SetAccessCheck(ByVal pPtid As String) As Boolean
    Dim sSql   As String
    Dim sSpcno As String
    Dim sSpcYY As String
    
    Dim RS   As Recordset
    sSpcYY = BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4, 1)
    
    sSql = " SELECT max(spcno) as spcno FROM " & T_BBS201 & _
           " WHERE " & DBW("ptid=", pPtid) & " AND " & DBW("spcyy=", sSpcYY)
    
    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    If Not RS.EOF Then
        sSpcno = RS.Fields("spcno").Value & ""
        
        sSql = " SELECT storeleg FROM " & T_BBS201 & " WHERE " & DBW("spcyy=", sSpcYY) & " AND " & DBW("spcno=", sSpcno)
        Set RS = New Recordset
        RS.Open sSql, DBConn
        If Not RS.EOF Then
            If (Trim(RS.Fields("storeleg").Value & "")) <> "" Then SetAccessCheck = True
        End If
    End If
    
    Set RS = Nothing
End Function
Public Function SetWardAccess(ByVal pPtid As String, ByVal pBussDiv As String, ByVal pRcvDt As String, _
                              ByVal pRcvTm As String, pRcviD As String) As Boolean
    Dim blnOkNot As Boolean
    
    
    DBConn.BeginTrans
    
    blnOkNot = OrderAccess(pPtid, pBussDiv, pRcvDt, pRcvTm, pRcviD)
    
    If blnOkNot = True Then
        DBConn.CommitTrans
        SetWardAccess = True
    Else
        DBConn.RollbackTrans
    End If
                              
End Function

Private Function OrderAccess(ByVal ptid As String, ByVal Bussdiv As String, _
                             ByVal rcvdt As String, ByVal rcvtm As String, _
                             ByVal rcvid As String) As Boolean
    '----------------------------------------------------------------------------
    '처방을 읽으면서, 접수처리를 한다.
    '----------------------------------------------------------------------------
    Dim objNumbers  As clsBBSNumbers
    Dim RsOrder     As Recordset
    Dim RS          As Recordset
    Dim sSql        As String
    Dim i           As Long
    
    Dim accdt       As String
    Dim accseq      As Long
    Dim SQLTmp      As String
On Error GoTo OrderAccessProcess_error
    
    '처방을 읽는다.----------------------------------------------------------------------
    'Set RsOrder = New recordset
    '접수번호를 구한다.-------------------------------------------------------------------
    Set objNumbers = New clsBBSNumbers
    With objNumbers
        accdt = .Get_AccdtFormat
        accseq = Val(.Get_AccDT_Seq(accdt))
    End With
    Set RsOrder = Get_Order_TarGet(ptid, Bussdiv)
    With RsOrder
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                '처방 바디 update
                sSql = Set_AddAcessBody(ptid, .Fields("orddt").Value & "", .Fields("ordno").Value & "", _
                                       .Fields("ordseq").Value & "", accdt, accseq, rcvdt, rcvtm)
                DBConn.Execute sSql
                '처방 접수내역 생성
                sSql = Set_TransAcess(accdt, accseq, ptid, .Fields("orddt").Value & "", _
                                      CLng(.Fields("ordno").Value & ""), CLng(.Fields("ordseq").Value & ""), rcvid, _
                                      rcvdt, rcvtm, .Fields("pherefg").Value & "")
                DBConn.Execute sSql
                
                sSql = Set_AddAcessOrder(ptid, .Fields("orddt").Value & "", CLng(.Fields("ordno").Value & ""))
                DBConn.Execute sSql

                SQLTmp = Set_AccUnitSQL_203(ptid, accdt, accseq)
                sSql = medGetP(SQLTmp, 1, COL_DIV)
                DBConn.Execute sSql
                sSql = medGetP(SQLTmp, 2, COL_DIV)
                DBConn.Execute sSql
                
'#############################################################################
'OCS에 상태 업데이트

'#############################################################################
                
                accseq = accseq + 1
                .MoveNext
            Next i
        End If
            '처방 헤더 update
           
        Set RS = GetRTSingleOrder(ptid, Bussdiv)
        If Not RS.EOF Then
            For i = 1 To RS.RecordCount
                sSql = Set_AddAcessBody(ptid, RS.Fields("orddt").Value & "", RS.Fields("ordno").Value & "", RS.Fields("ordseq").Value & "", accdt, accseq, rcvdt, rcvtm)
                DBConn.Execute sSql
                sSql = Set_AddAcessOrder(ptid, RS.Fields("orddt").Value & "", CLng(RS.Fields("ordno").Value & ""))
                DBConn.Execute sSql

                SQLTmp = Set_AccUnitSQL_203(ptid, accdt, accseq)
                sSql = medGetP(SQLTmp, 1, COL_DIV)
                DBConn.Execute sSql
                sSql = medGetP(SQLTmp, 2, COL_DIV)
                DBConn.Execute sSql
                
'#############################################################################
'OCS에 상태 업데이트

'#############################################################################
                
                accseq = accseq + 1
                RS.MoveNext
            Next
        End If
        Set RS = Nothing
        
            '접수번호 update-------------------------------------------------------------------
        sSql = objNumbers.Set_NumbersCom099(BN_ACC_NO, accdt, accseq - 1)
        DBConn.Execute sSql
            
    End With
    Set RsOrder = Nothing
    Set objNumbers = Nothing
    OrderAccess = True
    Exit Function
    
OrderAccessProcess_error:
    Set objNumbers = Nothing
    Set RsOrder = Nothing
    OrderAccess = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function Collect_Target(ByVal ptid As String, ByVal Bussdiv As String) As Boolean
'---------------------------------------------------------------------------
'해당환자에 대한 채혈 대상여부를 판단
'처방이 채혈을 해야하는 처방인지 아니면 채혈을 않해도 무방한 채혈인지를 판단
'---------------------------------------------------------------------------
    Dim RS   As Recordset
    Dim sSql As String
    
    sSql = " SELECT b.* " & _
           " FROM " & _
                    T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c " & _
           " WHERE " & _
                               DBW("a.orddiv=", C_WORKAREA) & _
                     " AND " & DBW("a.ptid=        ", ptid) & _
                     " AND " & DBW("a.bussdiv=  ", Bussdiv) & _
                     " AND " & DBW("a.orddt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
                     " AND a.ptid=b.ptid" & _
                     " AND a.orddt=b.orddt" & _
                     " AND a.ordno=b.ordno " & _
                     " AND " & DBW("b.stscd=", stsORDER) & _
                     " AND (b.dcfg is null or b.dcfg='')" & _
                     " AND b.ordcd=c.testcd " & _
                     " AND c.applydt=( SELECT max(d.applydt) " & _
                                     " FROM " & T_BBS001 & " d " & _
                                     " WHERE d.testcd = b.ordcd " & _
                                     " AND d.applydt<=b.orddt " & _
                                     ") " & _
                     " AND " & DBW("c.xmethod<>", XMethod.xmNONE)
    Set RS = New Recordset
    RS.Open sSql, DBConn
    If RS.EOF Then
        Collect_Target = False
    Else
        Collect_Target = True
    End If
    Set RS = Nothing
End Function

Private Function Addspc_Exist2(ByVal ptid As String, ByVal Bussdiv As String) As Boolean
'----------------------------------------
'해당환자에 대한 추가처방 여부를 판단한다
'----------------------------------------
    Dim RS   As New Recordset
    Dim sSql As String
    
'    sSql = " SELECT * FROM " & T_BBS207 & _
'           " WHERE " & _
'                   "     " & DBW("ptid=       ", ptid) & _
'                   " AND " & DBW("busidiv= ", Bussdiv) & _
'                   " AND " & DBW("dongfg=", "0")
'검체추가요청이 여러번 발생한 경우 가장 최근것을 가저오도록 한다.
    sSql = " SELECT * FROM " & T_BBS207 & _
           " WHERE " & _
                   "     " & DBW("ptid=       ", ptid) & _
                   " AND " & DBW("busidiv= ", Bussdiv) & _
                   " AND " & DBW("dongfg=", "0") & _
                   " and reqdt =(select max(b.reqdt) from " & T_BBS207 & " b " & _
                              " where b.ptid = ptid " & _
                              " and b.busidiv=busidiv " & _
                              " and b.dongfg=dongfg) "
                   
    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    If RS.EOF Then
        Addspc_Exist2 = False
    Else
        Addspc_Exist2 = True
    End If
    
    Set RS = Nothing

End Function

Private Function Set_Collect204(ByVal Workdt As String, ByVal Wardid As String, ByVal Worktm As String, _
                                ByVal Workseq As String, ByVal colid As String, _
                                ByVal buildcd As String, ByVal orgbuildcd As String, _
                                ByVal spcyy As String, ByVal spcno As String, _
                                Optional ByVal orddiv As String = "B") As String
        

    Set_Collect204 = " insert into " & T_LAB204 & "(" & _
                     " workdt,wardid,worktm,seq,workarea,accdt,accseq,colid," & _
                     " buildcd,orgbuildcd,spcyy,spcno,orddiv)" & _
                     " values(" & _
                                DBV("workdt        ", Workdt, 1) & DBV("wardid  ", Wardid, 1) & _
                                DBV("worktm        ", Worktm, 1) & DBV("seq    ", Workseq, 1) & _
                                "'','', 0," & _
                                DBV("colid          ", colid, 1) & DBV("buildcd", buildcd, 1) & _
                                DBV("orgbuildcd", orgbuildcd, 1) & DBV("spcyy    ", spcyy, 1) & _
                                DBV("spcno          ", spcno, 1) & DBV("orddiv     ", orddiv) & _
                            ")"

End Function

Private Function Set_SpcAndOrddt(ByVal ptid As String, ByVal Bussdiv As String, _
                                 ByVal spcyy As String, ByVal spcno As String, _
                                 ByVal coldt As String, ByVal coltm As String, _
                                 ByVal colid As String) As String
    Dim RS     As Recordset
    Dim strTmp As String
    Dim sSql   As String
    Dim ii     As Integer

    'dc된 처방은 제외하기 위해 s2ord101과 조인
    sSql = " SELECT distinct a.ptid,a.orddt,a.ordno,a.wardid" & _
           " FROM " & T_LAB102 & " b, " & T_LAB101 & " a " & _
           " WHERE " & DBW("a.ptid=", ptid) & _
           " AND " & DBW("a.bussdiv=", Bussdiv) & _
           " AND " & DBW("a.donefg= ", stsORDER) & _
           " AND " & DBW("a.orddiv=", C_WORKAREA) & _
           " AND " & DBW("a.orddt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
           " AND b.ptid = a.ptid " & _
           " AND b.orddt = a.orddt " & _
           " AND b.ordno = a.ordno " & _
           " AND (b.dcfg = '' or b.dcfg is null) "
           
    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    If RS.EOF Then
        Set_SpcAndOrddt = ""
        Set RS = Nothing
        Exit Function
    End If
    If Not RS.EOF Then
       Do Until RS.EOF
            strTmp = " insert into " & T_BBS102 & " (ptid,orddt,ordno,spcyy,spcno,coldt,coltm,colid) " & _
                            " values(" & _
                                       DBV("ptid", RS.Fields("ptid").Value & "", 1) & _
                                       DBV("orddt", RS.Fields("orddt").Value & "", 1) & _
                                       DBV("ordno", RS.Fields("ordno").Value & "", 1) & _
                                       DBV("spcyy", spcyy, 1) & _
                                       DBV("spcno", spcno, 1) & _
                                       DBV("coldt", coldt, 1) & _
                                       DBV("coltm", coltm, 1) & _
                                       DBV("colid", colid) & ")"
            
            If Set_SpcAndOrddt = "" Then
                Set_SpcAndOrddt = strTmp
            Else
                Set_SpcAndOrddt = Set_SpcAndOrddt & COL_DIV & strTmp
            End If
            
            RS.MoveNext
        Loop
    Else
        Set_SpcAndOrddt = ""
    End If
    Set RS = Nothing
End Function
Private Function Set_CollectOrder(ByVal ptid As String, ByVal Bussdiv As String) As String
'---------------------------------
'처방 바디를 업데이트한다.
'처방상태에 있는 모든처방이 대상임
'---------------------------------
    Dim RS          As Recordset
    Dim strPtid     As String
    Dim strOrdDt    As String
    Dim strOrdNo    As String
    Dim strOrdSeq   As String
    Dim strSql      As String
    Dim strTmp      As String
    Dim i           As Long
    
    strSql = " SELECT DISTINCT a.ptid,a.orddt,a.ordno,a.wardid,b.ordseq" & _
             " FROM " & T_LAB102 & " b," & T_LAB101 & " a " & _
             " WHERE " & DBW("a.ptid=", ptid) & _
             " AND " & DBW("a.bussdiv=", Bussdiv) & _
             " AND " & DBW("a.donefg= ", stsORDER) & _
             " AND " & DBW("a.orddiv=", C_WORKAREA) & _
             " AND " & DBW("a.orddt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
             " AND a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno " & _
             " AND (b.dcfg = '' or b.dcfg is null) "
    
'    MsgBox "처방내역 업데이트", vbExclamation
    
    Set RS = New Recordset
    Call RS.Open(strSql, DBConn)
    For i = 1 To RS.RecordCount
        strPtid = RS.Fields("ptid").Value & ""
        strOrdDt = RS.Fields("orddt").Value & ""
        strOrdNo = RS.Fields("ordno").Value & ""
        strOrdSeq = RS.Fields("ordseq").Value & ""
'원본
'        strTmp = " UPDATE " & T_LAB102 & _
'                 " SET " & _
'                           DBW("donefg", stsCOLLECT, 3) & _
'                           DBW("stscd ", stsCOLLECT, 2) & _
'                 " WHERE " & DBW("ptid=", strPtid) & _
'                 " AND   " & DBW("orddt=", strOrdDt) & _
'                 " AND   " & DBW("ordno=", strOrdNo) & _
'                 " AND   " & DBW("ordseq=", strOrdSeq) & _
'                 " AND   " & DBW("stscd=", stsORDER) & COL_DIV & _
'                 " UPDATE " & T_LAB101 & _
'                 " SET    " & DBW("donefg=", stsCOLLECT) & _
'                 " WHERE  " & _
'                                   DBW("ptid=   ", ptid) & _
'                         " AND " & DBW("bussdiv=", Bussdiv) & _
'                         " AND " & DBW("donefg= ", stsORDER) & _
'                         " AND " & DBW("orddiv=", C_WORKAREA) & _
'                         " AND " & DBW("orddt<=", Format(GetSystemDate, PRESENTDATE_FORMAT))
        '전주예수병원용
        Dim strBussdiv As String
        
        If Bussdiv = "1" Then
            strBussdiv = " and patsect in ('E','O') "
        ElseIf Bussdiv = "2" Then
            strBussdiv = " and patsect ='I' "
        End If
        
        
'상태전송 미포함
'        strTmp = " UPDATE mdbldort " & _
'                 " SET " & _
'                           DBW("donefg", stsCOLLECT, 3) & _
'                           DBW("stscd ", stsCOLLECT, 2) & _
'                 " WHERE " & DBW("patno=", strPtid) & _
'                 " AND   orddate=to_date(" & strOrdDt & ",'yyyymmdd')" & _
'                 " AND   " & DBW("ordseqno=", strOrdNo) & _
'                 " AND   " & DBW("stscd=", stsORDER) & COL_DIV & _
'                 " UPDATE mdbldort" & _
'                 " SET    " & DBW("donefg=", stsCOLLECT) & _
'                 " WHERE  " & _
'                                   DBW("patno=   ", ptid) & _
'                         strBussdiv & _
'                         " AND " & DBW("donefg= ", stsORDER) & _
'                         " AND orddate<=to_date(" & Format(GetSystemDate, "yyyymmdd") & ",'yyyymmdd')"
'OCS에 상태전송 포함
'procstat를 B로 상태 전송

        strTmp = " UPDATE mdbldort " & _
                 " SET " & _
                           DBW("donefg", stsCOLLECT, 3) & _
                           DBW("stscd ", stsCOLLECT, 3) & _
                           DBW("procstat", "B", 2) & _
                 " WHERE " & DBW("patno=", strPtid) & _
                 " AND   orddate=to_date(" & strOrdDt & ",'yyyymmdd')" & _
                 " AND   " & DBW("ordseqno=", strOrdNo) & _
                 " AND   " & DBW("stscd=", stsORDER) & COL_DIV & _
                 " UPDATE mdbldort" & _
                 " SET    " & DBW("donefg=", stsCOLLECT) & _
                 " WHERE  " & _
                                   DBW("patno=   ", ptid) & _
                         strBussdiv & _
                         " AND " & DBW("donefg= ", stsORDER) & _
                         " AND orddate<=to_date(" & Format(GetSystemDate, "yyyymmdd") & ",'yyyymmdd')"
        
        If Set_CollectOrder = "" Then
            Set_CollectOrder = strTmp
        Else
            Set_CollectOrder = Set_CollectOrder & COL_DIV & strTmp
        End If
        
        RS.MoveNext
        
    Next i
    Set RS = Nothing
End Function

Private Function Set_BBS207(ByVal ptid As String, ByVal Bussdiv As String, _
                            ByVal spcyy As String, ByVal spcno As String) As String
'----------------------
'검체추가 요청서 Update
'----------------------
    Set_BBS207 = " update " & T_BBS207 & " set " & _
                                           DBW("spcyy", spcyy, 3) & _
                                           DBW("spcno", spcno, 3) & _
                                           DBW("dongfg", stsCOLLECT, 2) & _
                   " WHERE " & _
                             "     " & DBW("ptid=", ptid) & _
                             " AND " & DBW("busidiv=", Bussdiv) & _
                             " AND " & DBW("dongfg=", stsORDER)

End Function

Private Function Get_TestBuild() As String
    Get_TestBuild = ObjSysInfo.buildingcd & COL_DIV & ObjSysInfo.buildingcd
End Function

Private Function Get_Order_TarGet(ByVal ptid As String, ByVal Bussdiv As String) As Recordset
    Dim strSql As String
    
    Set Get_Order_TarGet = New Recordset
    
    strSql = " SELECT a.ptid,a.orddt,a.ordno,a.ordseq,d.pherefg " & _
             " FROM " & T_LAB102 & " a," & T_LAB101 & " b," & T_BBS001 & " c," & T_BBS006 & " d" & _
             " WHERE " & _
             "     " & DBW("a.ptid", ptid, 2) & _
             " AND " & DBW("b.bussdiv", Bussdiv, 2) & _
             " AND " & DBW("b.orddiv", C_WORKAREA, 2) & _
             " AND " & DBW("b.orddt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
             " AND a.ptid=b.ptid" & _
             " AND a.orddt=b.orddt" & _
             " AND a.ordno=b.ordno" & _
             " AND (a.dcfg='' or a.dcfg is null)" & _
             " AND a.donefg in (" & _
                              DBV("donefg", stsORDER, 1) & _
                              DBV("donefg", stsCOLLECT) & ") " & _
             " AND a.ordcd=c.testcd" & _
             " AND c.compocd=d.compocd" & _
             " AND c.applydt=(SELECT max(z.applydt) FROM " & T_BBS001 & " z" & _
                             " WHERE " & _
                                           DBW("z.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
                                     " AND z.testcd=a.ordcd" & _
                             ")"
                             
    
    Set Get_Order_TarGet = New Recordset
    Get_Order_TarGet.Open strSql, DBConn
    
End Function
Private Function GetRTSingleOrder(ByVal ptid As String, ByVal Bussdiv As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT b.ptid,b.orddt,b.ordno,b.ordseq  " & _
         " FROM " & T_BBS001 & " c," & T_LAB102 & " b," & T_LAB101 & " a" & _
         " WHERE " & _
                   DBW("a.ptid=", ptid) & _
         " AND " & DBW("a.bussdiv=", Bussdiv) & _
         " AND " & DBW("a.orddiv", C_WORKAREA, 2) & _
         " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno" & _
         " AND (b.dcfg='' or b.dcfg is null)" & _
         " AND b.donefg in (" & _
                          DBV("donefg", stsORDER, 1) & _
                          DBV("donefg", stsCOLLECT) & ") " & _
         " AND b.ordcd=c.testcd" & _
         " AND " & DBW("c.xmethod=", XMethod.xmNONE) & _
         " AND c.applydt=(SELECT max(z.applydt) FROM " & T_BBS001 & " z" & _
                         " WHERE " & _
                                      DBW("z.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
                                " AND z.testcd=b.ordcd" & _
                        ")"
    Set GetRTSingleOrder = New Recordset
    GetRTSingleOrder.Open sSql, DBConn
End Function

Private Function Set_AddAcessBody(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As String, _
                                 ByVal ordseq As String, ByVal accdt As String, ByVal accseq As String, _
                                 ByVal rcvdt As String, ByVal rcvtm As String) As String
   
'    Set_AddAcessBody = " update " & T_LAB102 & " set " & _
'                                                DBW("rcvdt", rcvdt, 3) & _
'                                                DBW("rcvtm", rcvtm, 3) & _
'                                                DBW("donefg", stsACCESS, 3) & _
'                                                DBW("stscd", stsACCESS, 3) & _
'                                                DBW("workarea", C_WORKAREA, 3) & _
'                                                DBW("accdt", accdt, 3) & _
'                                                DBW("accseq", accseq, 2) & _
'                       " WHERE " & _
'                                "     " & DBW("ptid", ptid, 2) & _
'                                " AND " & DBW("orddt", orddt, 2) & _
'                                " AND " & DBW("ordno", OrdNo, 2) & _
'                                " AND " & DBW("ordseq", ordseq, 2)
    '예수병원용
''상태전송 미포함
'    Set_AddAcessBody = " update mdbldort set " & _
'                                                " acptdate=to_date(" & rcvdt & rcvtm & ",'yyyymmdd hh24:mi:ss')," & _
'                                                DBW("stscd", stsACCESS, 3) & _
'                                                DBW("workarea", C_WORKAREA, 3) & _
'                                                DBW("accdt", accdt, 3) & _
'                                                DBW("accseq", accseq, 2) & _
'                       " WHERE " & _
'                                "     " & DBW("patno", ptid, 2) & _
'                                " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
'                                " AND " & DBW("ordseqno", OrdNo, 2)
'상태전송 포함
'OCS에 상태를 E로 전송
    Set_AddAcessBody = " update mdbldort set " & _
                                                " acptdate=to_date(" & rcvdt & rcvtm & ",'yyyymmdd hh24:mi:ss')," & _
                                                DBW("stscd", stsACCESS, 3) & _
                                                DBW("workarea", C_WORKAREA, 3) & _
                                                DBW("accdt", accdt, 3) & _
                                                DBW("accseq", accseq, 3) & _
                                                DBW("procstat", "E", 2) & _
                       " WHERE " & _
                                "     " & DBW("patno", ptid, 2) & _
                                " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                                " AND " & DBW("ordseqno", OrdNo, 2)
                                
End Function

Private Function Set_TransAcess(ByVal accdt As String, ByVal accseq As Long, ByVal ptid As String, _
                               ByVal orddt As String, ByVal OrdNo As Long, _
                               ByVal ordseq As Long, ByVal rcvid As String, ByVal rcvdt As String, _
                               ByVal rcvtm As String, ByVal pheresis As String) As String
'처방 접수내역에 Insert 한다.
    
    Set_TransAcess = " insert into " & T_BBS202 & "(" & _
                 " workarea,accdt,accseq,ptid,orddt,ordno,ordseq,rcvdt,rcvtm,rcvid,stscd," & _
                 " canceldt,canceltm,cancelid,cancelfg,cancelrsn,pheresisfg)" & _
                 " values(" & _
                 DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & _
                 DBV("ptid", ptid, 1) & DBV("orddt", orddt, 1) & _
                 DBV("ordno", OrdNo, 1) & DBV("ordseq", ordseq, 1) & DBV("rcvdt", rcvdt, 1) & _
                 DBV("rcvtm", rcvtm, 1) & DBV("rcvid", rcvid, 1) & _
                 DBV("stscd", stsACCESS, 1) & DBV("canceldt", "", 1) & DBV("canceltm", "", 1) & _
                 DBV("cancelid", "", 1) & DBV("cancelfg", "0", 1) & _
                 DBV("cancelrsn", "", 1) & DBV("pheresisfg", pheresis) & ")"
End Function

Private Function Set_AddAcessOrder(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long) As String
    
'    Set_AddAcessOrder = " update " & T_LAB101 & " set" & _
'                                                  DBW("donefg", stsACCESS, 2) & _
'                        " WHERE " & _
'                        "       " & DBW("ptid", ptid, 2) & _
'                        " AND donefg in(" & _
'                                   DBV("donefg", stsORDER, 1) & _
'                                   DBV("donefg", stsCOLLECT) & ")" & _
'                        " AND " & DBW("orddt", orddt, 2) & _
'                        " AND " & DBW("ordno", OrdNo, 2)
    
    '예수병원용
    Set_AddAcessOrder = " update mdbldort set" & _
                                                  DBW("donefg", stsACCESS, 2) & _
                        " WHERE " & _
                        "       " & DBW("patno", ptid, 2) & _
                        " AND donefg in(" & _
                                   DBV("donefg", stsORDER, 1) & _
                                   DBV("donefg", stsCOLLECT) & ")" & _
                        " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                        " AND " & DBW("ordseqno", OrdNo, 2)
End Function

Public Function Get_BuildingCd(ByVal DeptCd As String) As String
'병동에 해당되는 건물코드를 가지고 온다.
    Get_BuildingCd = ObjSysInfo.buildingcd
End Function

'############################
'# 검체 추가요청 내역      ##
'############################
Public Function Get_SpcAdd(Wardid As String) As Recordset
    '검체추가 목록 조회
    'dongfg=채혈처리 여부임(0:처방,1:채혈)
    Dim strSql As String
    
    strSql = " SELECT distinct a.ptid,a.orddt,b." & F_PTNM & " as ptnm, c.orddiv, " & _
             "        c.bussdiv,a.reqdt,a.bedindt," & F_DOB2("b") & " as dob, " & _
             "        " & F_SEX2("b") & " as sex, c.orddoct, c.majdoct, c.deptcd, " & _
             "        c.roomid, c.hosilid " & _
             " FROM " & _
                      T_LAB102 & " d," & T_LAB101 & " c," & T_HIS001 & " b," & T_BBS207 & " a" & _
             " WHERE" & _
                      " " & DBW(" a.wardid", Wardid, 2) & _
                      " AND " & DBW("d.workarea=", "B") & _
                      " AND d.accdt=a.accdt" & _
                      " AND d.accseq=a.accno" & _
                      " AND c.ptid=d.ptid" & _
                      " AND c.orddt=d.orddt" & _
                      " AND c.ordno=d.ordno" & _
                      " AND a.ptid=b." & F_PTID & _
                      " AND d.ptid=a.ptid " & _
                      " AND " & DBW("a.dongfg", stsORDER, 2)
    
    Set Get_SpcAdd = New Recordset
    Get_SpcAdd.Open strSql, DBConn
End Function

'############################
'# 공통으로 쓰는 문장      ##
'############################

Public Function ER_Chk(ByVal ptid As String, ByVal orddt As String) As String
    '환자Id,처방일 처방 번호를 가지고환자의 응급여부를 체크한다.
    Dim strSql  As String
    Dim RS      As Recordset
    
    strSql = " SELECT distinct statfg" & _
             " FROM " & _
                      T_LAB102 & " a," & T_LAB101 & " b " & _
             " WHERE" & _
                      "     " & DBW("a.donefg", stsORDER, 2) & _
                      " AND " & DBW("b.orddiv", C_WORKAREA, 2) & _
                      " AND " & DBW("a.ptid", ptid, 2) & _
                      " AND a.ptid=b.ptid " & _
                      " AND " & DBW("a.orddt", orddt, 2) & _
                      " AND " & DBW("a.orddt<=", Format(GetSystemDate, "yyyymmdd"))
                      
                      

    Set RS = New Recordset
    RS.Open strSql, DBConn
    
    If RS.EOF = True Then
        ER_Chk = "0"
    Else
        Do Until RS.EOF = True
            ER_Chk = Trim(RS.Fields("STATFG").Value & "")
            If ER_Chk = "1" Then
                Set RS = Nothing
                Exit Function
            Else
                ER_Chk = "0"
            End If
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Function

Public Function TestBuildCd(ByVal blgcd As String) As String
'건물코드를 가지고 검사가 수행될 건물을 찾는다.
    TestBuildCd = ObjSysInfo.buildingcd & COL_DIV & ObjSysInfo.buildingcd
End Function

Public Function TestBldNm(ByVal bldcd As String) As String
'검사수행될 건물 코드 가지고온다.
    Dim strSql As String
    Dim DrRS As New Recordset
    
    strSql = " SELECT field1" & _
             " FROM " & T_COM003 & _
             " WHERE" & _
                      " " & DBW("cdindex", BC2_CENTER, 2) & _
                      "   AND " & DBW("cdval1", bldcd, 2)
    
    Set DrRS = New Recordset
    DrRS.Open strSql, DBConn
    
    If Not DrRS.EOF = True Then
        TestBldNm = DrRS.Fields("field1").Value & ""
    Else
        TestBldNm = ""
    End If
    
    Set DrRS = Nothing
End Function

'############################
'# 간호사/외래채혈(BBS104,5)##
'############################
Public Function Get_CollectOrder(ByVal Searchkey As String, ByVal TF As Boolean, _
                                 ByVal Bussdiv As String, ByVal orddt As String) As String
'간호사채혈에서 채혈대상이 되는 환자만 조회한다.
    Get_CollectOrder = " SELECT distinct a." & F_PTID & " as ptid," & _
                                " a." & F_PTNM & " as ptnm," & F_SSN2("a") & " as ssn " & _
                       " FROM " & _
                                T_HIS001 & " a," & T_LAB101 & " b" & _
                       " WHERE" & _
                                "      a." & F_PTID & " = b.ptid " & _
                                "  AND " & DBW("b.bussdiv", Bussdiv, 2) & _
                                "  AND " & DBW("b.donefg", stsORDER, 2) & _
                                "  AND " & DBW("b.orddt", orddt, 2) & _
                                "  AND " & DBW("b.orddiv", C_WORKAREA, 2)
   
   '환자ID는 존재여부만 확인한다.
    If TF = True Then
        Get_CollectOrder = Get_CollectOrder & _
                                            " AND " & DBW("a." & F_PTID, Searchkey, 2)
    Else
        Get_CollectOrder = Get_CollectOrder & _
                                            " AND a." & F_PTNM & " like " & "'" & Searchkey & "%" & "'"
    End If

End Function

'###############################
'# 채혈에서 환자조회(BBS104,5)##
'###############################


Public Function Get_SideCollect(ByVal Searchkey As String, ByVal blnSearch As Boolean, ByVal orddt As String) As String

'외래채혈 접수에서 처방 채혈상태의 환자를 조회한다.
    
    
    Get_SideCollect = " SELECT distinct a." & F_PTID & " as ptid,a." & F_PTNM & " as ptnm," & F_SSN2("a") & " as ssn " & _
                " FROM " & _
                         T_HIS001 & " a," & T_LAB101 & " b," & T_LAB102 & " c" & _
                " WHERE" & _
                         " a." & F_PTID & " = b.ptid " & _
                         " AND " & DBW("bussdiv", BBSBUSSDIV.stsNotBed, 2) & _
                         " AND b.donefg in(" & _
                                          DBV("donefg", stsORDER, 1) & _
                                          DBV("donefg", stsCOLLECT) & ") " & _
                         " AND " & DBW("b.orddt", orddt, 2) & _
                         " AND " & DBW("b.orddiv", C_WORKAREA, 2) & _
                         " AND (c.dcfg is null or " & DBW("c.dcfg<>", "1") & ")" & _
                         " AND b.ptid=c.ptid" & _
                         " AND b.orddt=c.orddt" & _
                         " AND b.ordno=c.ordno"
                        
   '환자ID는 존재여부만 확인한다.
    If blnSearch = True Then
        Get_SideCollect = Get_SideCollect & _
                                          " AND  " & DBW("a." & F_PTID, Searchkey, 2)
    Else
        Get_SideCollect = Get_SideCollect & _
                                          " AND  a." & F_PTNM & " like " & "'" & Searchkey & "%" & "'"
    End If
             
End Function

Public Function Get_AddSpcInFo(Search As String, ByVal blnSearch As Boolean) As String
'추가채혈여부 확인
    Get_AddSpcInFo = " SELECT a.ptid,a.accdt,a.accno,b." & F_PTNM & _
                            " as ptnm," & F_SSN2("b") & " as ssn,a.orddt " & _
                     " FROM " & _
                              T_BBS207 & " a," & T_HIS001 & " b" & _
                     " WHERE" & _
                              "     " & DBW("a.dongfg", stsORDER, 2) & _
                              " AND a.ptid=b." & F_PTID
            
    If blnSearch = True Then
        Get_AddSpcInFo = Get_AddSpcInFo & _
                                        " AND " & DBW("a.ptid", Search, 2)
    Else
        Get_AddSpcInFo = Get_AddSpcInFo & _
                                        " AND b." & F_PTNM & " like " & "'" & Search & "%" & "'"
    End If

End Function

Public Function Get_Order_107(ByVal ptid As String, ByVal orddt As String, _
                                Optional ByVal Search As String = "") As Recordset
    'frm107(외래채혈시 환자id를 넘기면 처방 정보를 보여준다.
    '처방(donefg=0),업무구분(bussdiv in(1))=외래

    Dim strSql As String
    '처방일,처방번호,처방Seq,검사코드,응급구분,수혈예정일,진료과코드,처방의코드,영수증번호,병동ID,처방리마크
    strSql = " SELECT b.orddt,b.ordno,b.ordseq,b.ordcd,b.statfg,b.donefg,b.unitqty,a.reqdt,a.reqtm," & _
                      " a.deptcd,a.orddoct,b.receptno,a.bussdiv,a.wardid,b.mesg " & _
             " FROM " & _
                      T_LAB101 & " a," & T_LAB102 & " b" & _
             " WHERE" & _
                      " " & DBW("a.ptid", ptid, 2) & _
                      " AND " & DBW("b.orddt", orddt, 2) & _
                      " AND a.bussdiv in ('1','3') " & _
                      " AND " & DBW("a.orddiv=", C_WORKAREA) & _
                      " AND a.ptid =b.ptid " & _
                      " AND a.orddt=b.orddt " & _
                      " AND a.ordno=b.ordno "
             
    If Search = "" Then
        strSql = strSql & _
                        " AND " & DBW("b.donefg", stsORDER, 2) & _
                        " ORDER BY a.ptid,b.ptid"
    Else
        strSql = strSql & _
                        " AND b.donefg in(" & _
                                        DBV("donefg", stsORDER, 1) & _
                                        DBV("donefg", stsCOLLECT) & ") " & _
                        " ORDER BY a.ptid,b.ptid"
    End If
             
    Set Get_Order_107 = New Recordset
    Get_Order_107.Open strSql, DBConn

End Function

Public Function Get_ADDSPC(ByVal ptid As String, ByVal accdt As String, accseq As String) As Recordset
    '외래채혈에서 접수일자와 접수번호를 가지고
    '채혈시 필요한 정보를 조회해온다.
    Dim strSql As String
    
    strSql = " SELECT b.donefg,b.orddt,b.ordno,b.ordseq,b.ordcd,b.statfg,a.reqdt,a.reqtm,a.deptcd," & _
                      " b.receptno,a.orddoct," & _
                      " a.bussdiv,a.wardid,b.mesg ,c.reqdt as reqdt1" & _
             " FROM " & _
                      T_LAB101 & " a," & T_LAB102 & " b," & T_BBS207 & " c" & _
             " WHERE" & _
                      "     " & DBW("b.workarea", C_WORKAREA, 2) & _
                      " AND " & DBW("b.accdt", accdt, 2) & _
                      " AND " & DBW("b.accseq", accseq, 2) & _
                      " AND " & DBW("b.ptid", ptid, 2) & _
                      " AND " & DBW("c.dongfg", stsORDER, 2) & _
                      " AND a.ptid =b.ptid " & _
                      " AND a.orddt=b.orddt " & _
                      " AND a.ordno=b.ordno " & _
                      " AND b.accdt=c.accdt" & _
                      " AND b.accseq=c.accno "
                      
                      
             
    Set Get_ADDSPC = New Recordset
    Get_ADDSPC.Open strSql, DBConn
End Function

'############################
'# 외래채혈(BBS104 )       ##
'############################

Public Function Get_ORDER_105(ByVal ptid As String, ByVal orddt As String) As Recordset
    'frm902(외래채혈 환자id를 넘기면 처방 정보를 보여준다.
    '처방(donefg=0),업무구분(bussdiv in(2,3))=병동,응급

    Dim sSql As String
    
    sSql = " SELECT b.orddt,b.ordno,b.ordseq,b.ordcd,b.statfg,a.reqdt,a.reqtm,a.deptcd,a.orddoct," & _
           "        a.bussdiv,a.wardid,b.mesg,b.receptno " & _
           " FROM " & _
                    T_LAB101 & " a," & T_LAB102 & " b" & _
           " WHERE" & _
                    "     " & DBW("a.orddiv=", C_WORKAREA) & " AND " & DBW("a.ptid", ptid, 2) & _
                    " AND " & DBW("b.orddt", orddt, 2) & _
                    " AND " & DBW("a.donefg", stsORDER, 2) & _
                    " AND a.bussdiv in('1','3') " & _
                    " AND a.ptid =b.ptid " & _
                    " AND a.orddt=b.orddt " & _
                    " AND a.ordno=b.ordno " & _
           " ORDER BY a.ptid,b.ordno "
             
    Set Get_ORDER_105 = New Recordset
    Get_ORDER_105.Open sSql, DBConn
    
    
End Function

Public Function Set_AccUnitSQL_203(ByVal ptid As String, ByVal accdt As String, ByVal accseq As String) As String
    Dim sSql    As String
    Dim sSql1   As String
    
    Dim RS   As Recordset

    '추가 필드 변수(BBS201에서 가지고온다)
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    Dim strRack         As String
    Dim strRow          As String
    Dim strCol          As String
    Dim strBuildcd      As String
    Dim strOrgBuildCd   As String

    '추가 필드 변수(BBS203에서 가지고온다)
    Dim strAssigncnt    As String
    Dim strDeliverycnt  As String
    Dim strretcnt       As String
    Dim strexpcnt       As String
    Dim strbagcnt       As String
    Dim strAssigncancelcnt As String
    Dim strYearChk As String
    Dim strYear    As String
    
    strYear = BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4, 1)
    
    '이미 검체가 있는경우
    sSql = " SELECT max(spcno) as spcno FROM " & T_BBS201 & _
           " WHERE " & DBW("ptid=", ptid) & _
           " AND " & DBW("spcyy=", strYear) & _
           " ORDER BY spcno desc"

    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    strYearChk = Mid(Format(GetSystemDate, "YYYYMMDD"), 5)
    
    If RS.EOF Or Val(RS.Fields("spcno").Value & "") = 0 Then
        If strYearChk >= "0101" And strYearChk <= "0110" Then
            strYear = Val(strYear) - 1
            sSql = " SELECT max(spcno) as spcno FROM " & T_BBS201 & _
                   " WHERE " & DBW("ptid=", ptid) & " AND " & DBW("spcyy=", strYear) & _
                   " ORDER BY spcno desc"
            Set RS = Nothing
            Set RS = New Recordset
            RS.Open sSql, DBConn
        End If
    End If
    
    If Not RS.EOF Then
        sSql = " SELECT * FROM " & T_BBS201 & " WHERE  " & DBW("spcyy=", BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4, 1)) & " AND " & DBW("spcno=", RS.Fields("spcno").Value & "")
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open sSql, DBConn
        
        If RS.EOF = False Then
            strSpcYY = RS.Fields("spcyy").Value & ""
            strSpcNo = RS.Fields("spcno").Value & ""
            strRack = RS.Fields("storeleg").Value & ""
            strRow = RS.Fields("storerno").Value & ""
            strCol = RS.Fields("storecno").Value & ""
            strBuildcd = RS.Fields("buildcd").Value & ""
            strOrgBuildCd = RS.Fields("orgbuildcd").Value & ""
        End If
    End If

    Set RS = Nothing

    '이미 접수되어 검사가 된것이있는지 체크
    sSql = " SELECT * FROM " & T_LAB203 & " WHERE " & _
                     DBW("workarea=", C_WORKAREA) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq)
    
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open sSql, DBConn

    '이미 작성되어 있는경우 update
    If Not RS.EOF Then
        strAssigncnt = RS.Fields("assigncnt").Value & ""
        strDeliverycnt = RS.Fields("deliverycnt").Value & ""
        strretcnt = RS.Fields("retcnt").Value & ""
        strexpcnt = RS.Fields("expcnt").Value & ""
        strbagcnt = RS.Fields("bagcnt").Value & ""
        strAssigncancelcnt = RS.Fields("assigncancelcnt").Value & ""

        sSql = " update " & T_BBS203 & " set " & _
                            DBW("assigncnt", strAssigncnt, 3) & _
                            DBW("deliverycnt", strDeliverycnt, 3) & _
                            DBW("retcnt", strretcnt, 3) & _
                            DBW("expcnt", strexpcnt, 3) & _
                            DBW("bagcnt", strbagcnt, 3) & _
                            DBW("assigncancelcnt", strAssigncancelcnt, 3) & _
                            DBW("spcyy", strSpcYY, 3) & _
                            DBW("spcno", strSpcNo, 3) & _
                            DBW("storeleg", strRack, 3) & _
                            DBW("storerno", strRow, 3) & _
                            DBW("storecno", strCol, 3) & _
                            DBW("buildcd", strBuildcd, 3) & _
                            DBW("orgbuildcd", strOrgBuildCd, 3) & _
                            DBW("ptid", ptid, 2) & _
              " WHERE " & _
              "       " & DBW("workarea", C_WORKAREA, 2) & _
              "  AND  " & DBW("accdt", accdt, 2) & _
              "  AND  " & DBW("accseq", accseq, 2)
    Else

        sSql = " insert into " & T_BBS203 & "(workarea,accdt,accseq,assigncnt,deliverycnt," & _
               " retcnt,expcnt,bagcnt,assigncancelcnt,spcyy,spcno,storeleg,storerno,storecno," & _
               " buildcd,orgbuildcd,ptid)" & _
               " values(" & _
                          DBV("workarea", C_WORKAREA, 1) & _
                          DBV("accdt", accdt, 1) & _
                          DBV("accseq", accseq, 1) & _
                          DBV("assigncnt", "", 1) & _
                          DBV("deliverycnt", "", 1) & _
                          DBV("retcnt", "", 1) & _
                          DBV("expcnt", "", 1) & _
                          DBV("bagcnt", "", 1) & _
                          DBV("assigncancelcnt", "", 1) & _
                          DBV("spcyy", strSpcYY, 1) & _
                          DBV("spcno", strSpcNo, 1) & _
                          DBV("storeleg", strRack, 1) & _
                          DBV("storerno", strRow, 1) & _
                          DBV("storecno", strCol, 1) & _
                          DBV("buildcd", strBuildcd, 1) & _
                          DBV("orgbuildcd", strOrgBuildCd, 1) & _
                          DBV("ptid", ptid) & ")"


    End If
    '환자의 최종 검체 정보를 업데이트 한다.
    sSql1 = " update " & T_BBS203 & " set " & _
                        DBW("spcyy", strSpcYY, 3) & _
                        DBW("spcno", strSpcNo, 3) & _
                        DBW("storeleg", strRack, 3) & _
                        DBW("storerno", strRow, 3) & _
                        DBW("storecno", strCol, 3) & _
                        DBW("buildcd", strBuildcd, 3) & _
                        DBW("orgbuildcd", strOrgBuildCd, 3) & _
                        DBW("ptid", ptid, 2) & _
          " WHERE " & _
          "       " & DBW("ptid", ptid, 2)
    

    Set_AccUnitSQL_203 = sSql & COL_DIV & sSql1

    Set RS = Nothing

End Function

Public Function Set_UpdateSQL_203(ByVal ptid As String, ByVal spcyy As String, ByVal spcno As String) As String
    Dim sSql    As String
    Dim RS      As Recordset
    
    '이미 검체가 있는경우
    sSql = " SELECT * FROM " & T_BBS201 & _
           " WHERE " & DBW("ptid=", ptid) & " AND " & DBW("spcyy=", spcyy) & " AND " & DBW("spcno=", spcno)
           

    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    If Not RS.EOF Then
        '환자의 최종 검체 정보를 업데이트 한다.
        sSql = " update " & T_BBS203 & " set " & _
                            DBW("spcyy", RS.Fields("spcyy").Value & "", 3) & _
                            DBW("spcno", RS.Fields("spcno").Value & "", 3) & _
                            DBW("storeleg", RS.Fields("storeleg").Value & "", 3) & _
                            DBW("storerno", RS.Fields("storerno").Value & "", 3) & _
                            DBW("storecno", RS.Fields("storecno").Value & "", 3) & _
                            DBW("buildcd", RS.Fields("buildcd").Value & "", 3) & _
                            DBW("orgbuildcd", RS.Fields("orgbuildcd").Value & "", 3) & _
                            DBW("ptid", RS.Fields("ptid").Value & "", 2) & _
              " WHERE " & _
              "       " & DBW("ptid", ptid, 2)
    
        Set_UpdateSQL_203 = sSql
    End If
    Set RS = Nothing
End Function

Private Sub Class_Initialize()
    
    If TRANS_REQUIRE_USED = True Then
        stsORDER = BBSOrdStatus.stsORDER
        stsCOLLECT = BBSOrdStatus.stsCOLLECT
        stsACCESS = BBSOrdStatus.stsACCESS
    Else
        stsORDER = BBSOrderStatus.stsORDER
        stsCOLLECT = BBSOrderStatus.stsCOLLECT
        stsACCESS = BBSOrderStatus.stsACCESS
    End If

End Sub

Private Sub Class_Terminate()
    Set mvarBldDic = Nothing
End Sub
