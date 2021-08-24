VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISCollectioin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+--------------------------------------------------------------------------------------+
'|  1. Class 명 : clsCollection(DLL용)
'|  2. 기    능 : 채혈/접수내역(LAB201)을 생성한다.
'|  3. 작 성 자 : 김미경
'|
'|  CopyRight(C) 1999 대련엠티에스
'|
'+--------------------------------------------------------------------------------------+

Option Explicit

'%  클래스 clsCollection의 Data Attributes
Public SpcYY As String      '/* 검체번호-년도   */
Public SpcNo As Long        '/* 검체번호        */
Public Ptid As String       '/* 환자 ID         */
Public PtNm As String       '/* 환자명         */
Public AgeDay As Long       '/* 환자 일령       */
Public Sex As String        '/* 환자 성별   */
Public BedInDt As String    '/* 입원일          */
Public DeptCd As String     '/* 진료과          */
Public OrdDoct As String    '/* 처방의          */
Public MajDoct As String    '/* 주치의          */
Public WorkArea As String   '/* Work Area(접수번호용) */
Public AccDt As String      '/* 접수일(접수번호용)    */
Public AccSeq As Long    '/* 접수순번(접수번호용)  */
Public StsCd As String      '/* STATUS - '0 ':처방,1:채혈,2:접수,3:In-Process,4:결과,5:확인,6:수정 */
Public ReqTotCnt As Integer '/* Required Item Total Count */
Public ReqInputCnt As Integer  '/* Required Item Input Count */
Public VfyDt As String      '/* 결과확인일자    */
Public VfyTm As String      '/* 결과확인시간    */
Public VfyId As String      '/* 결과확인자      */
Public ColDt As String      '/* 채혈일자        */
Public ColTm As String      '/* 채혈시간        */
Public ColId As String      '/* 채혈자          */
Public RcvDt As String      '/* 접수일자        */
Public RcvTm As String      '/* 접수시간        */
Public RcvId As String      '/* 접수자          */
Public EntDt As String      '/* 입력일자        */
Public EntTm As String      '/* 입력시간        */
Public EntId As String      '/* 입력자          */
Public SpcCd As String      '/* 검체코드        */
Public SpcNm As String      '/* 검체명        */
Public MultiFg As String    '/* 복수검체여부  */
Public OrgAccNo As String   '/* 원 접수번호     */
Public WardId As String     '/* 병동코드        */
Public RoomId As String     '/* 병실코드     */
Public BedID As String      '/* 침상코드    */
Public HosilId As String    '/* 호실코드    */
Public FootNoteFg As String '/* FOOTNOTE여부('0':부,'1':여) */
Public StoreCd As String    '/* 검체보관 코드   */
Public RptFg As String      '/* REPORT출력여부('0':부,'1':여) */
Public TestDiv As String    '/* 검사구분 - 1:기타검사,2:미생물검사 */
Public QcFg As String       '/* QC여부('0':부,'1':여) */
Public RmkCd As String      '/* 검체 Remark 코드 */
Public StatFg As String     '/* 응급여부    */
Public BuildCd As String    '/* 건물코드    */
Public BuildNm As String    '/* 건물코드명    */
Public OrgBuildCd As String '/* 채혈된 건물코드    */
Public TestNames As String  '/* Barcode에 출력될 검사명들 */

Public MornFg As String     '/* 아침채혈여부 */

Public OrdDt  As String
Public OrdNo  As Integer
Public OrdSeq As Integer
Public OrdCd  As String

Public SpcGrp As String     '/* 검체군  */
Public SpcSeq As String     '/* 복수검체의 Seq */

'= Lab 채혈접수 관련 Data Type ='

Private Type LabNumbers
   WorkArea As String
   AccDt    As String
   AccSeq   As Long
   BuildCd  As String
End Type
   
Private Type BarcodeLabel
   Location As String
   WorkArea As String
   AccDt    As String
   AccSeq   As Long
   OrdDt    As String
   ColDt    As String
   ColTm    As String
   SpcNo    As String
   PtNm     As String
   Ptid     As String
   SpcNm    As String
   StatFg   As String
   StoreCd  As String
   WardId   As String
   TestNames As String
   LabelCnt As String
End Type
   
   
Public ErrMsg       As String
Public ColCount     As Integer
Public CollectDone  As Boolean

Public WorkDt  As String
Public WorkTm  As String
Public WorkSeq As Integer

Private blnBatchCol    As Boolean
Private blnSetTrans    As Boolean
Private ColDone        As Boolean  '복수검체 채혈완료 여부
Private ReadyToCollect As Boolean
Private BarCount       As Integer
Private MultiSpcFlag   As String

Private CreateLabNo()   As LabNumbers
Private BarcodeBuffer() As BarcodeLabel

Private dicColList As New clsDictionary
Private objMySql   As New clsLISSqlStatement
Private objColSql  As New clsLISSqlCollection
'Private objBarPrt  As New clsBarcode

Private blnUseAutoPrint As Boolean

Private objBarCode As clsBarcode

'바코드라벨 자동출력기능 사용여부
Public Property Let SetTrans(ByVal vData As Boolean)
    blnSetTrans = vData
End Property

Public Property Let UseAutoPrint(ByVal vData As Boolean)
    blnUseAutoPrint = vData
End Property

'% 채혈내역을 생성하기 위한 기초데이타를 저장한다.
Public Sub AddLabCollect(ByRef tmpData() As String)
   
    Dim strKey As String, strData As String
    Dim strOrdInfo As String
    
    '순서 : 건물(0),WorkArea(1),검체코드(2),보관구분(3),응급여부(4),희망채혈일시(5),
    '       검사구분(6),복수검체여부(7),검체군(8),처방일(9),처방번호(10),처방Seq(11),검사코드(12),
    '       부서코드(13),처방의(14),주치의(15),검사명(16),라벨출력장수(17),LabDiv(18),검체약어명(19),
    '       미생물접수번호범위(20),처방정보(strOrdInfo)
    strKey = tmpData(0) & COL_DIV & tmpData(1) & COL_DIV & tmpData(2) & COL_DIV & tmpData(3) & COL_DIV & _
             tmpData(4) & COL_DIV & tmpData(5)
    '처방일,처방번호,처방Seq
    strOrdInfo = tmpData(9) & ":" & tmpData(10) & ":" & tmpData(11)
    
    If dicColList.Exists(strKey) Then
        dicColList.KeyChange strKey
        '** Barcode 라벨에 출력될 검사명들...
        dicColList.Fields("testnm") = dicColList.Fields("testnm") & "," & tmpData(16)
        '** 처방정보
        dicColList.Fields("ordinfo") = dicColList.Fields("ordinfo") & "," & strOrdInfo
    Else
        strData = tmpData(6) & COL_DIV & tmpData(7) & COL_DIV & tmpData(8) & COL_DIV & tmpData(9) & COL_DIV & tmpData(10) & COL_DIV & _
                  tmpData(11) & COL_DIV & tmpData(12) & COL_DIV & tmpData(13) & COL_DIV & tmpData(14) & COL_DIV & _
                  tmpData(15) & COL_DIV & tmpData(16) & COL_DIV & tmpData(17) & COL_DIV & tmpData(18) & COL_DIV & _
                  tmpData(19) & COL_DIV & tmpData(20) & COL_DIV & strOrdInfo
        dicColList.AddNew strKey, strData
    End If
    
    ReadyToCollect = True
    CollectDone = False
   
End Sub

'% 채혈내역을 생성하기 위한 기초데이타를 저장한다.
Public Sub SetAddLabCollect(ByRef tmpData() As String)
   
    Dim strKey As String, strData As String
    Dim strOrdInfo As String
    Dim AryTmp() As String
    Dim ii As Integer
    Dim blnChk As Boolean
    
    blnChk = False
    
    '순서 : 건물(0),WorkArea(1),검체코드(2),보관구분(3),응급여부(4),희망채혈일시(5),
    '       검사구분(6),복수검체여부(7),검체군(8),처방일(9),처방번호(10),처방Seq(11),검사코드(12),
    '       부서코드(13),처방의(14),주치의(15),검사명(16),라벨출력장수(17),LabDiv(18),검체약어명(19),
    '       미생물접수번호범위(20),처방정보(strOrdInfo)
    strKey = tmpData(0) & COL_DIV & tmpData(1) & COL_DIV & tmpData(2) & COL_DIV & tmpData(3) & COL_DIV & _
             tmpData(4) & COL_DIV & tmpData(5)
    '처방일,처방번호,처방Seq
    strOrdInfo = tmpData(9) & ":" & tmpData(10) & ":" & tmpData(11)
    
    If dicColList.Exists(strKey) Then
        dicColList.KeyChange strKey
        
        AryTmp = Split(dicColList.Fields("testnm"), ",")
        For ii = LBound(AryTmp) To UBound(AryTmp)
            If AryTmp(ii) <> tmpData(16) Then
                blnChk = True
            Else
                blnChk = False
                Exit For
            End If
        Next
        If blnChk = True Then
            '** Barcode 라벨에 출력될 검사명들...
            dicColList.Fields("testnm") = dicColList.Fields("testnm") & "," & tmpData(16)
            '** 처방정보
            dicColList.Fields("ordinfo") = dicColList.Fields("ordinfo") & "," & strOrdInfo
        End If
    Else
        strData = tmpData(6) & COL_DIV & tmpData(7) & COL_DIV & tmpData(8) & COL_DIV & tmpData(9) & COL_DIV & tmpData(10) & COL_DIV & _
                  tmpData(11) & COL_DIV & tmpData(12) & COL_DIV & tmpData(13) & COL_DIV & tmpData(14) & COL_DIV & _
                  tmpData(15) & COL_DIV & tmpData(16) & COL_DIV & tmpData(17) & COL_DIV & tmpData(18) & COL_DIV & _
                  tmpData(19) & COL_DIV & tmpData(20) & COL_DIV & strOrdInfo
        dicColList.AddNew strKey, strData
    End If
    
    ReadyToCollect = True
    CollectDone = False
   
End Sub

Public Sub SetColData(ByRef aryData As Variant)

    SpcYY = LIS_BarDiv & aryData(0)      '검체년도
    Ptid = aryData(1)       '환자ID
    PtNm = aryData(2)       '환자성명
    Sex = aryData(3)        '환자성별
    AgeDay = aryData(4)     '환자일령
    BedInDt = aryData(5)    '입원일
    EntDt = aryData(6)      '입력일
    EntTm = aryData(7)      '입력시간
    EntId = aryData(8)      '입력자
    OrgAccNo = aryData(9)   '원접수번호
    ColDt = aryData(10)     '채혈일
    ColId = aryData(11)     '채혈자
    WardId = aryData(12)    '병동ID
    HosilId = aryData(13)   '병실ID
    RoomId = aryData(14)    '병실ID
    BedID = aryData(15)     '침상ID
    OrgBuildCd = aryData(16) '** 채혈이 수행되는 건물코드

End Sub

Public Sub SetWardCol(ByRef pWorkDt As String, ByRef pWorkTm As String, _
                      Optional ByVal pWardId As String = "")

    Dim tmpRs As Recordset
    
    blnBatchCol = True    '병동별 일괄채혈 구분
    
    WorkDt = pWorkDt    'Format(Getsystemdate, CS_DateDbFormat)
    WorkTm = pWorkTm    'Format(Getsystemdate, CS_TimeDbFormat)
    If Trim(pWardId) <> "" Then WardId = pWardId

    Set tmpRs = New Recordset
    tmpRs.Open objColSql.SqlGetWardColSeq(WorkDt, WardId, WorkTm), DBConn
        
    If tmpRs.EOF Then
        WorkSeq = 0
    Else
        WorkSeq = Val("" & tmpRs.Fields("ColSeq").Value)
    End If
    Set tmpRs = Nothing
   
End Sub


'% 채혈Procedure를 수행하여 채혈접수내역을 생성한다.
'% ( 복수검체의 경우 한꺼번에 모두 채혈처리 한다 --> 아래 DoCollection_Backup은 한번에 하나씩.. )
Public Function DoCollection(Optional ByRef ProgressBar As Variant) As Boolean
   
    Dim i As Integer, j As Integer
    Dim tmpSql As String
    Dim tmpSpcRs As Recordset
    Dim tmpMultiRs As Recordset
    Dim blnMultiCol As Boolean
   
    If Not ReadyToCollect Then
       DoCollection = False
       Exit Function
    End If
       
    ColCount = 0
    ReDim BarcodeBuffer(0)
   
    StsCd = enStsCd.StsCd_LIS_Collection  'Status
    ReqTotCnt = 0       'Required Item Total Count
    ReqInputCnt = 0     'Required Item Input Count
    FootNoteFg = "0"    'FootNote구분
    RptFg = "0"         '레포트 출력 여부
    QcFg = "0"          'QC여부
    VfyDt = ""          '결과확인일
    VfyTm = ""          '결과확인시간
    VfyId = 0           '결과확인자
    RcvDt = ""          '접수일
    RcvTm = ""          '접수시간
    RcvId = 0           '접수자
    RmkCd = ""          '검체리마크
    TestNames = ""      '검사약어명String
    
'    Dim objBuild As clsBasisData
    Dim strBuild As String
    
    dicColList.Sort = True
    dicColList.MoveFirst
    
    While (Not dicColList.EOF)
      
        '>> 기준이 달라질때마다 채혈내역 한건 생성...
        BarCount = 0  '바코드출력장수 초기화
         
        With dicColList
          
            MultiSpcFlag = .Fields("MultiFg")
              
            WorkArea = .Fields("WorkArea")
            
            
'            ColDt = Format(Getsystemdate, CS_DateDbFormat)     '채혈일
'            ColTm = Format(Getsystemdate, CS_TimeDbFormat)     '채혈시간

            SpcCd = .Fields("SpcCd")                          '검체
            MultiFg = .Fields("MultiFg")                      '복수검체여부
            StoreCd = .Fields("StoreCd")                      '검체보관구분
            TestDiv = .Fields("TestDiv")                      '검사구분
            StatFg = .Fields("StatFg")                        '응급여부
            SpcGrp = .Fields("SpcGrp")                        '검체군
            BuildCd = .Fields("BuildCd")                      '** 건물코드
              
            TestNames = .Fields("testnm")
            SpcNm = .Fields("spcnm")
            
            OrdDt = .Fields("OrdDt")
            OrdNo = Val(.Fields("OrdNo"))
            OrdSeq = Val(.Fields("OrdSeq"))
              
            OrdDoct = Trim(.Fields("OrdDoct"))
            MajDoct = Trim(.Fields("MajDoct"))
            DeptCd = .Fields("DeptCd")
              
            '검체 전달 장소
            If P_ApplyBuildingInfo Then
'                Set objBuild = Nothing
'                Set objBuild = New clsBasisData
                strBuild = GetBuildNm(BuildCd)
'                Set objBuild = Nothing
                
                If strBuild <> "" Then
                    BuildNm = strBuild
                Else
                    BuildNm = ""
                End If
'                If ObjLISComCode.Building.Exists(BuildCd) Then
'                    ObjLISComCode.Building.KeyChange BuildCd
'                    BuildNm = ObjLISComCode.Building.Fields("buildnm")
'                Else
'                    BuildNm = ""
'                End If
            Else
                BuildNm = LABName
            End If
                
              
            '----> 복수검체일 경우
            If MultiSpcFlag = "Y" Then
                
                tmpSql = objColSql.SqlGetMultiSpc(SpcCd)
                Set tmpSpcRs = Nothing
                Set tmpSpcRs = New Recordset
                tmpSpcRs.Open tmpSql, DBConn
                 
                tmpSql = objColSql.SqlMultiSpcCnt2(Ptid, OrdDt, OrdNo, OrdSeq)
                Set tmpMultiRs = Nothing
                Set tmpMultiRs = New Recordset
                tmpMultiRs.Open tmpSql, DBConn
             
                Dim iMultiCnt As Integer  '복수검체 갯수
                Dim iDoneCnt As Integer   '생성된 갯수
                Dim sDoneSpc As String    '생성된 검체Seq
             
                iMultiCnt = tmpSpcRs.RecordCount
                iDoneCnt = tmpMultiRs.RecordCount    '이미 생성된 검체수
                For j = 1 To iDoneCnt
                    sDoneSpc = sDoneSpc & tmpMultiRs.Fields("SpcSeq").Value & ";"
                    tmpMultiRs.MoveNext
                Next
                Set tmpMultiRs = Nothing
                
                blnMultiCol = True
                While ((Not tmpSpcRs.EOF) And blnMultiCol)
                    If InStr(Trim(sDoneSpc), tmpSpcRs.Fields("SpcSeq").Value & "") = 0 Then
                        iDoneCnt = iDoneCnt + 1
                        ColDone = IIf(iDoneCnt = iMultiCnt, True, False)    '모두 채혈됬을 경우 True
                        
                        SpcNm = tmpSpcRs.Fields("BarNm").Value & ""  '검체명
                        
                        SpcNo = GetSpcNo(SpcYY)                 '검체번호
                        If SpcNo = 0 Then GoTo Err_Trap         '검체번호를 제대로 못가져온 경우...
                        AccSeq = GetLabNo(WorkArea, AccDt)      'AccSeq
                        If AccSeq = 0 Then GoTo Err_Trap        'AccSeq를 제대로 못가져온 경우...
                      
                        SpcSeq = "" & tmpSpcRs.Fields("SpcSeq").Value   '연속검사Seq
                        SpcCd = "" & tmpSpcRs.Fields("SpcCd").Value     '연속검사검체
                        
                        'BarCount = 1
                        If BarCount < Val(.Fields("labelcnt")) Then BarCount = Val(.Fields("labelcnt"))   '출력장수
                        
                        Call SaveData  '저장
                        If Not CollectDone Then GoTo Err_Trap
                        
                        '** Barcode Label을 출력한다.
                        If P_UseBarcodeSystem Then
                            Call PrintBarcode_New(.Fields("coldate"), Trim(frm168POCTCol.cboBarCode.Text))
                            '** 자동출력기능 사용(주로 ER(응급실))하면 Barcode출력내용을 Buffering 한다.
                            'If blnUseAutoPrint AND P_ErBarcodeAutoPrint Then Call SetBarcodeBuffer(.Fields("coldate"))
                        End If
                        
                        '연속검사를 시간차를 두고 채혈할 경우 한건씩 채혈처리...
                        If P_CollectMultiSpc = "1" Then blnMultiCol = False
                        
                    End If
                    tmpSpcRs.MoveNext
                Wend
                 
                Set tmpSpcRs = Nothing
                 
            '----> 일반검체일 경우
            Else
                
                SpcNo = GetSpcNo(SpcYY)                 '검체번호
                If SpcNo = 0 Then GoTo Err_Trap         '검체번호를 제대로 못가져온 경우...
                AccSeq = GetLabNo(WorkArea, AccDt)      'AccSeq
                If AccSeq = 0 Then GoTo Err_Trap        'AccSeq를 제대로 못가져온 경우...
                MultiFg = ""
                ColDone = True
                If BarCount < Val(.Fields("labelcnt")) Then BarCount = Val(.Fields("labelcnt"))   '출력장수
                Call SaveData  '저장
                If Not CollectDone Then GoTo Err_Trap
                
                '** Barcode Label을 출력한다.
                If P_UseBarcodeSystem Then
                    Call PrintBarcode_New(.Fields("coldate"), Trim(frm168POCTCol.cboBarCode.Text))
                    '** 자동출력기능 사용(주로 ER(응급실))하면 Barcode출력내용을 Buffering 한다.
                    'If blnUseAutoPrint AND P_ErBarcodeAutoPrint Then Call SetBarcodeBuffer(.Fields("coldate"))
                End If
                
            End If
            
        End With
      
        'Progress Bar 증가...
        If Not IsMissing(ProgressBar) Then
            If ProgressBar.Value = ProgressBar.Max Then ProgressBar.Max = ProgressBar.Max + 1
            ProgressBar.Value = ProgressBar.Value + 1
        End If
        DoEvents
        
        dicColList.MoveNext
        
    Wend
   
    DoCollection = True
    Exit Function
   
Err_Trap:
    DoCollection = False
    Set tmpSpcRs = Nothing
    Set tmpMultiRs = Nothing

End Function


'% 클래스의 채혈 내역을 DB로 저장한다.
Public Sub SaveData()
   
    Dim SqlStmt As String
    Dim tmpStr As String
    Dim tmpRs As Recordset
    Dim i As Integer
    Dim aryOrders() As String
    Dim aryOrdInfo() As String
    
    '/* 전주예수병원 추가변수 */ ====================================
    Dim clsTransfer As New clsLISTransfer
    '================================================================
    
    'ColCount 증가
    ColCount = ColCount + 1
    ReDim Preserve CreateLabNo(ColCount)
         
    '--채혈접수 동시 수행을 위해 생성된 LabNo를 Keep한다.
    CreateLabNo(ColCount).WorkArea = WorkArea
    CreateLabNo(ColCount).AccDt = AccDt
    CreateLabNo(ColCount).AccSeq = AccSeq
    CreateLabNo(ColCount).BuildCd = BuildCd
   
    On Error GoTo Err_Trap
   
    If blnSetTrans Then DBConn.BeginTrans   '-------------------------<< Transaction Start >>

    '--채혈접수내역 생성
    SqlStmt = objColSql.CreateSql_Collection(Me)    '// Sql문장 생성
    Call DBConn.Execute(SqlStmt)
    
    '--연속검사 내역 생성
    If MultiSpcFlag = "Y" Then       '---> 연속검사일 경우 LAB203 생성
        aryOrders = Split(dicColList.Fields("ordinfo"), ",")
        For i = LBound(aryOrders) To UBound(aryOrders)
            aryOrdInfo = Split(aryOrders(i), ":")
            OrdDt = aryOrdInfo(0)
            OrdNo = aryOrdInfo(1)
            OrdSeq = aryOrdInfo(2)
            SqlStmt = objColSql.CreateSql_MultiSpc(Me)   '// Sql문장 생성
            Call DBConn.Execute(SqlStmt)
        Next
    End If
    
    '--병동채혈 내역 생성
    If blnBatchCol Then              '---> 병동별 일괄채혈일 경우...
        WorkSeq = WorkSeq + 1
        SqlStmt = objColSql.CreateSql_WardCol(Me)    '// Sql문장 생성
        Call DBConn.Execute(SqlStmt)
    End If
   
    '--처방 Status Follow Up
    aryOrders = Split(dicColList.Fields("ordinfo"), ",")
    For i = LBound(aryOrders) To UBound(aryOrders)
        aryOrdInfo = Split(aryOrders(i), ":")
        If ColDone Then            '--> 연속검사일 경우 모두 채혈수행이 되야 처방Status Follow Up.
            'LAB102의 DoneFg = '1' : 채혈완료
            SqlStmt = objColSql.CreateSql_UpdateLabNo(Me, Ptid, aryOrdInfo(0), aryOrdInfo(1), _
                                                      aryOrdInfo(2), enStsCd.StsCd_LIS_Collection, "1")         '// Sql문장 생성
            Call DBConn.Execute(SqlStmt)
        Else
            'LAB102의 DoneFg = '0' : 채혈미완료
            SqlStmt = objColSql.CreateSql_UpdateLabNo(Me, Ptid, aryOrdInfo(0), aryOrdInfo(1), _
                                                      aryOrdInfo(2), enStsCd.StsCd_LIS_Collection)    '// Sql문장 생성
            Call DBConn.Execute(SqlStmt)
        End If
          
        SqlStmt = objColSql.CreateSql_HeaderUpdate(Ptid, aryOrdInfo(0), aryOrdInfo(1), _
                                                   enStsCd.StsCd_LIS_Collection, 1)
        Set tmpRs = Nothing
        Set tmpRs = New Recordset
        tmpRs.Open SqlStmt, DBConn
        
        If tmpRs.EOF Then
            Set tmpRs = Nothing
             
            SqlStmt = objColSql.CreateSql_HeaderUpdate(Ptid, aryOrdInfo(0), aryOrdInfo(1), _
                                                       enStsCd.StsCd_LIS_Collection, 2)
            Call DBConn.Execute(SqlStmt)
        Else
            Set tmpRs = Nothing
        End If
        
        '-- 예수병원 Status Follow Up ======================================================
        If clsTransfer.OCS_Status_Update(Ptid, aryOrdInfo(0), aryOrdInfo(1), _
                                         ColDt, ColTm, enStsCd.StsCd_LIS_Collection) = False Then
            
            MsgBox "검사진행상태 업로드 오류가 발생하였습니다.", vbCritical
            GoTo Err_Trap
            
        End If
        '===================================================================================
Skip:
    Next
    
    If blnSetTrans Then DBConn.CommitTrans        '-------------------------<< Transaction End >>
    
    CollectDone = True
    
    Set clsTransfer = Nothing
    
    Exit Sub

Err_Trap:
    If blnSetTrans Then MsgBox Err.Description, vbExclamation
    CollectDone = False
    Set tmpRs = Nothing
    Set clsTransfer = Nothing

End Sub

'% 검체번호를 부여한다.
Private Function GetSpcNo(ByVal SpcYY As String) As Long
   
    Dim SpcNoRs As Recordset
    Dim tmpSql As String
   
    GetSpcNo = 0
   
    '먼저 Lock을 걸기위해 의미없는 update문 실행.
    'select는 Lock이 안걸리기 때문에 극히 짧은 순간에도
    '검체번호가 중복될수 있음.. 을 방지.    2000.3.10  by 김미경
    tmpSql = objColSql.CreateSql_SpcNo(SpcYY, 4)
    If blnSetTrans Then DBConn.BeginTrans
    'dbconn.BeginTrans
   
    On Error GoTo Err_Trap
   
    DBConn.Execute tmpSql   'Lock 걸림
   
    tmpSql = objColSql.CreateSql_SpcNo(SpcYY, 1)
    Set SpcNoRs = New Recordset
    SpcNoRs.Open tmpSql, DBConn
    
    If SpcNoRs.EOF Then
        GetSpcNo = 1
        tmpSql = objColSql.CreateSql_SpcNo(SpcYY, 2, GetSpcNo)
    Else
        GetSpcNo = Val("" & SpcNoRs.Fields("Seq").Value) + 1
        tmpSql = objColSql.CreateSql_SpcNo(SpcYY, 3, GetSpcNo)
    End If
    Set SpcNoRs = Nothing
   
    DBConn.Execute tmpSql
    If blnSetTrans Then DBConn.CommitTrans
   
    Exit Function
   
Err_Trap:
    GetSpcNo = 0
    ErrMsg = "검체번호 부여 중 오류발생!"
    MsgBox Err.Description, vbExclamation
    Set SpcNoRs = Nothing
   
End Function

'% 접수번호를 부여한다.
Private Function GetLabNo(ByVal pWorkArea As String, ByRef pAccDt As String) As Long
   
    Dim tmpRs       As New Recordset
    Dim tmpSql      As String
    Dim LabDiv      As String
    Dim tmpStr      As String
'접수 Seq의 자릿수 증가로 Max값이 Integer범위를 벗어나므로 변경
'Modify By Legends 2003/12/02
    Dim tmpRng1     As Long
    Dim tmpRng2     As Long
    Dim tmpSpcGrp   As String
   
    GetLabNo = 0
    LabDiv = dicColList.Fields("labdiv")
   
    tmpRng1 = 1
'접수 Seq의 자릿수 증가로 수정
'Modify By legends 2003/12/02
'    tmpRng2 = 9999
    tmpRng2 = 99999
    tmpSpcGrp = "0"
    Select Case LabDiv
    Case LabDiv_ByDay:       '일단위
        pAccDt = Format(GetSystemDate, LIS_LabDayFormat)
    Case LabDiv_ByMonth:     '월단위
        pAccDt = Format(GetSystemDate, LIS_LabMonthFormat)
    Case LabDiv_ByYear:      '년단위
        pAccDt = Format(GetSystemDate, LIS_LabYearFormat)
    Case LabDiv_BySpc:       '검체단위
        pAccDt = Format(GetSystemDate, LIS_LabMonthFormat)
        tmpSpcGrp = SpcGrp
        '검체단위인데 검체군 마스터가 없는 경우... (이런일은 발생하면 안되지...)
        If dicColList.Fields("labrange") <> "" Then
            tmpStr = dicColList.Fields("labrange")      '검체군별 Lab No 범위
            tmpRng1 = Val(medGetP(tmpStr, 2, ";"))      '검체군별 Lab No 시작번호
            tmpRng2 = Val(medGetP(tmpStr, 3, ";"))      '검체군별 Lab No 끝번호
        End If
    Case Else:
        pAccDt = Format(GetSystemDate, LIS_LabDayFormat)
    End Select
   
    If blnSetTrans Then DBConn.BeginTrans
    
    '---------------------------------------------------------
'    gSql = "Lock Table TableName In Exclusive  Mode Nowate "
'    gclsDataRs.DataRs (gSql)
    '---------------------------------------------------------
    '먼저 Lock을 걸기위해 의미없는 update문 실행.
    'select는 Lock이 안걸리기 때문에 극히 짧은 순간에도
    '접수번호가 중복될수 있음.. 을 방지.    2000.3.10  by 김미경
    tmpSql = objColSql.CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 4)
   
    On Error GoTo Err_Trap
    
    DBConn.Execute tmpSql   'Lock 걸림
   
    '// Sql문장 생성
    tmpSql = objColSql.CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 1)
    '//
    tmpRs.Open tmpSql, DBConn
    
    If tmpRs.EOF Then
        GetLabNo = tmpRng1
        tmpSql = objColSql.CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 2, GetLabNo)
    Else
        GetLabNo = Val(tmpRs.Fields("Seq").Value & "")
        If GetLabNo < tmpRng1 Then
            GetLabNo = tmpRng1
        Else
            GetLabNo = GetLabNo + 1
        End If
        
        If GetLabNo > tmpRng2 Then
            ErrMsg = "접수번호가 Range(" & tmpRng1 & "-" & tmpRng2 & ")를 벗어났읍니다. : " & GetLabNo
            GoTo Err_Trap
        End If
        tmpSql = objColSql.CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 3, GetLabNo)
    End If
    Set tmpRs = Nothing
      
    DBConn.Execute tmpSql
    If blnSetTrans Then DBConn.CommitTrans
   
    Exit Function
    Resume
Err_Trap:
    DBConn.RollbackTrans
    Set tmpRs = Nothing
    GetLabNo = 0
    Exit Function

End Function


Public Sub GetLabNumbers(ByVal AryIndex As Integer, ByRef pWorkArea As String, _
                         ByRef pAccDt As String, ByRef pAccSeq As Long, _
                         Optional ByRef pBuildCd As Variant)
   
    If CollectDone Then
        With CreateLabNo(AryIndex)
            pWorkArea = .WorkArea
            pAccDt = .AccDt
            pAccSeq = .AccSeq
            pBuildCd = .BuildCd
        End With
    End If
   
End Sub


Public Sub PrintBarcode(ByVal pColTm As String)
    Dim objBar      As clsBarcode
    Dim strWardId   As String
    
'    strWardId = IIf(HosilId <> "", WardId & "/" & HosilId, WardId)
    strWardId = WardId
    
    Set objBar = New clsBarcode
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    '-- 변경 희망채혈일시 -> 실제채혈일시 By M.G.Choi 2006.01.19
    Dim strColDt    As String
    Dim strColTm    As String
    
    strColDt = Format(GetSystemDate, "yyyymmdd")
    strColTm = Format(GetSystemDate, "hh:mm")
    
    Call objBar.Label_PrintOut(BuildNm, WorkArea, Mid(AccDt, 3), Format(AccSeq, "#####"), _
                                  SpcYY & Format(SpcNo, LIS_BarFormat), Ptid, Mid(PtNm, 1, 10), _
                                  SpcNm, StoreCd, StatFg, IIf(Trim(WardId) <> "", strWardId, DeptCd), _
                                  Mid(strColDt, 5, 2) & "/" & Mid(strColDt, 7, 2), strColTm, _
                                  TestNames, BarCount, False, False)
    
    '** 원본 --------------------------------------------------------------------------------------
'    Call objBAR.Label_PrintOut(BuildNm, WorkArea, Mid(AccDt, 3), Format(AccSeq, "#####"), _
'                                  SpcYY & Format(SpcNo, LIS_BarFormat), Ptid, Mid(PtNm, 1, 10), _
'                                  SpcNm, StoreCd, StatFg, IIf(Trim(WardId) <> "", strWardId, DeptCd), _
'                                  Mid(OrdDt, 5, 2) & "/" & Mid(OrdDt, 7, 2), Mid(pColTm, 12, 5), _
'                                  TestNames, BarCount, False, False)
    '----------------------------------------------------------------------------------------------
    
'    Call medSleep(500)
    
    Set objBar = Nothing
        
                             
End Sub

Public Sub PrintBarcode_New(ByVal pColTm As String, ByVal pColSleep As String)
    Dim objBar      As clsBarcode
    Dim strWardId   As String
    
'    strWardId = IIf(HosilId <> "", WardId & "/" & HosilId, WardId)
    strWardId = WardId
    
    Set objBar = New clsBarcode
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    '-- 변경 희망채혈일시 -> 실제채혈일시 By M.G.Choi 2006.01.19
    Dim strColDt    As String
    Dim strColTm    As String
    
    strColDt = Format(GetSystemDate, "yyyymmdd")
    strColTm = Format(GetSystemDate, "hh:mm")
    
    Call objBar.Label_PrintOut(BuildNm, WorkArea, Mid(AccDt, 3), Format(AccSeq, "#####"), _
                                  SpcYY & Format(SpcNo, LIS_BarFormat), Ptid, Mid(PtNm, 1, 10), _
                                  SpcNm, StoreCd, StatFg, IIf(Trim(WardId) <> "", strWardId, DeptCd), _
                                  Mid(strColDt, 5, 2) & "/" & Mid(strColDt, 7, 2), strColTm, _
                                  TestNames, BarCount, False, False)
    
    '** 원본 --------------------------------------------------------------------------------------
'    Call objBAR.Label_PrintOut(BuildNm, WorkArea, Mid(AccDt, 3), Format(AccSeq, "#####"), _
'                                  SpcYY & Format(SpcNo, LIS_BarFormat), Ptid, Mid(PtNm, 1, 10), _
'                                  SpcNm, StoreCd, StatFg, IIf(Trim(WardId) <> "", strWardId, DeptCd), _
'                                  Mid(OrdDt, 5, 2) & "/" & Mid(OrdDt, 7, 2), Mid(pColTm, 12, 5), _
'                                  TestNames, BarCount, False, False)
    '----------------------------------------------------------------------------------------------
    
    If pColSleep > 0 Then
        Call medSleep(pColSleep)
    End If
    
    Set objBar = Nothing
        
                             
End Sub

Public Sub SetBarcodeBuffer(ByVal pColTm As String)

    ReDim Preserve BarcodeBuffer(UBound(BarcodeBuffer) + 1)
    
    With BarcodeBuffer(UBound(BarcodeBuffer))
        .Location = BuildNm
        .WorkArea = WorkArea
        .AccDt = Mid(AccDt, 3)
        .AccSeq = Format(AccSeq, "@@@@@")
        .SpcNo = SpcYY & Format(SpcNo, LIS_BarFormat)
        .Ptid = Ptid
        .PtNm = Mid(PtNm, 1, 10)
        .SpcNm = SpcNm
        .StoreCd = StoreCd
        .StatFg = StatFg
        .WardId = IIf(Trim(WardId) <> "", HosilId, DeptCd)
        .OrdDt = Mid(OrdDt, 5, 2) & "/" & Mid(OrdDt, 7, 2)
        .ColTm = Mid(pColTm, 12, 5)
        .TestNames = TestNames
        .LabelCnt = BarCount
    End With
    
End Sub

Public Function GetBarcodeLabel(ByVal AryIndex As Integer, Optional ByVal AccFg As Boolean = False) As String
   
    If CollectDone Then
        With BarcodeBuffer(AryIndex)
'            Call ObjLISComCode.BarInfo.Label_PrintOut(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .Ptid, .PtNm, _
'                              .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt, AccFg, False)
            Call objBarCode.Label_PrintOut(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .Ptid, .PtNm, _
                              .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt, AccFg, False)
        End With
    End If
   
End Function

Public Function GetBarcodeString(ByVal AryIndex As Integer, Optional ByVal AccFg As Boolean = False) As String
   
'    If CollectDone Then
'        With BarcodeBuffer(AryIndex)
'            GetBarcodeString = ObjLISComCode.BarInfo.Label_String(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .Ptid, .PtNm, _
'                              .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt, AccFg, False)
'        End With
'    End If
   If CollectDone Then
      With BarcodeBuffer(AryIndex)
         'Call MyBar.LabelPrint(.WorkArea, .ColDt, .AccSeq, .StatFg, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .WardId, .Location, .TestNames, .LabelCnt)
'         GetBarcodeString = ObjLISComCode.BarInfo.Label_String(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt)
         GetBarcodeString = objBarCode.Label_String(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .Ptid, .PtNm, .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt)
      End With
   End If
End Function

Public Function BarFormFeed()
'    Call ObjLISComCode.BarInfo.Label_FormFeed
    Call objBarCode.Label_FormFeed
End Function

'% Method 1 : CollectQuery
'%            Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%            클래스 clsCollection의 Data Attribute에 저장한다.

Public Function CollectQuery(ByVal pWorkArea As String, _
                             ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

    Dim SqlStmt As String
    Dim objRs As Recordset
   
    SqlStmt = "SELECT a.* " & _
              "FROM " & T_LAB201 & " a " & _
              "WHERE  " & DBW("a.workarea = ", pWorkArea) & " " & _
              "AND    " & DBW("a.accdt    = ", pAccDt) & " " & _
              "AND    " & DBW("a.accseq   = ", pAccSeq)
   
    Set objRs = New Recordset
    objRs.Open SqlStmt, DBConn
    
    If objRs.EOF Then
        CollectQuery = False
        Set objRs = Nothing
        Exit Function
    End If
   
    With objRs
      
        .MoveFirst
           
        SpcYY = Trim("" & .Fields("SpcYy").Value)
        SpcNo = Val("" & .Fields("SpcNo").Value)
        Ptid = Trim("" & .Fields("PtId").Value)
        Sex = Trim("" & .Fields("Sex").Value)
        AgeDay = Val("" & .Fields("AgeDay").Value)
        BedInDt = Trim("" & .Fields("BedInDt").Value)
        DeptCd = Trim("" & .Fields("DeptCd").Value)
        OrdDoct = Val("" & .Fields("OrdDoct").Value)
        MajDoct = Val("" & .Fields("MajDoct").Value)
        WorkArea = Trim("" & .Fields("WorkArea").Value)
        AccDt = Trim("" & .Fields("AccDt").Value)
        AccSeq = Val("" & .Fields("AccSeq").Value)
        StsCd = Trim("" & .Fields("StsCd").Value)
        ReqTotCnt = Val("" & .Fields("ReqTotCnt").Value)
        ReqInputCnt = Val("" & .Fields("ReqInputCnt").Value)
        VfyDt = Trim("" & .Fields("VfyDt").Value)
        VfyTm = Trim("" & .Fields("VfyTm").Value)
        VfyId = Val("" & .Fields("VfyId").Value)
        ColDt = Trim("" & .Fields("ColDt").Value)
        ColTm = Trim("" & .Fields("ColTm").Value)
        ColId = Val("" & .Fields("ColId").Value)
        RcvDt = Trim("" & .Fields("RcvDt").Value)
        RcvTm = Trim("" & .Fields("RcvTm").Value)
        RcvId = Val("" & .Fields("RcvId").Value)
        EntDt = Trim("" & .Fields("EntDt").Value)
        EntTm = Trim("" & .Fields("EntTm").Value)
        EntId = Val("" & .Fields("EntId").Value)
        SpcCd = Trim("" & .Fields("SpcCd").Value)
        MultiFg = Trim("" & .Fields("MultiFg").Value)
        OrgAccNo = Trim("" & .Fields("OrgAccNo").Value)
        WardId = Trim("" & .Fields("WardId").Value)
        RoomId = Trim("" & .Fields("RoomId").Value)
        BedID = Trim("" & .Fields("BedId").Value)
        HosilId = Trim("" & .Fields("HosilId").Value)
        FootNoteFg = Trim("" & .Fields("FootNoteFg").Value)
        StoreCd = Trim("" & .Fields("StoreCd").Value)
        RptFg = Trim("" & .Fields("RptFg").Value)
        TestDiv = Trim("" & .Fields("TestDiv").Value)
        QcFg = Trim("" & .Fields("QcFg").Value)
        RmkCd = Trim("" & .Fields("RmkCd").Value)
        StatFg = Trim("" & .Fields("StatFg").Value)
        BuildCd = Trim("" & .Fields("BuildCd").Value)
      
    End With
   
    Set objRs = Nothing
End Function

Private Sub Class_Initialize()
    blnBatchCol = False
    Erase BarcodeBuffer
    Set objMySql = New clsLISSqlStatement
    Set objColSql = New clsLISSqlCollection
    'Set objBarPrt = New clsBarcode
    Set dicColList = New clsDictionary
    Set objBarCode = New clsBarcode
    
    Call InitRtn
End Sub

Private Sub Class_Terminate()
    Set objMySql = Nothing
    Set objColSql = Nothing
    'Set objBarPrt = Nothing
    Set dicColList = Nothing
    Set objBarCode = Nothing
End Sub

Public Sub InitRtn()
    
    '** Dictionary : dicColList
    '** Key  : 건물,WorkArea,검체코드,보관구분,응급여부,희망채혈일시
    '** Data : 검사구분,복수검체여부,검체군,처방일,처방번호,처방Seq,검사코드,진료과
    '          처방의,주치의,검사명,라벨출력장수,접수번호부여기준,검체약어명,처방정보
    dicColList.Clear
    dicColList.DeleteAll
    dicColList.FieldInialize "buildcd,workarea,spccd,storecd,statfg,coldate", _
            "testdiv,multifg,spcgrp,orddt,ordno,ordseq,ordcd,deptcd," & _
            "orddoct,majdoct,testnm,labelcnt,labdiv,spcnm,labrange,ordinfo"
    dicColList.Sort = False
    
    ColCount = 0
    ReadyToCollect = False
    CollectDone = False
    ColDone = False
    ErrMsg = ""
    blnSetTrans = False
    
    Erase BarcodeBuffer
    
End Sub
Public Function ChkSpcnm(ByVal objordSheet As Object, ByVal FRowCnt As Integer, ByVal LRowCnt As Integer) As Integer
    Dim ii As Integer
    With objordSheet
        For ii = FRowCnt To LRowCnt
            .Row = ii
            .Col = enCOLLIST.tcSPCCD
            If Trim(.Value) = "" Then
                ChkSpcnm = ii
                Exit Function
            End If
        Next
    End With
End Function

Public Function CheckSameOrder(ByVal objordSheet As Object, ByVal intSel As Integer) As Integer

    Dim i As Integer, j As Integer
    Dim SaveCode As String
    Dim SaveSpc As String
    Dim SaveDate As String
    Dim SaveStatFg As String
    Dim strtmp  As VbMsgBoxResult
    Dim ChkBln  As Boolean
    Dim strOrdNo As String
    Dim strReqtm As String
    Dim strReqDt As String
    Dim strLastTm As String
    Dim k As Integer
    
    CheckSameOrder = 0
    ChkBln = False
    
    With objordSheet
   
        If intSel = 1 Then    '중복처방 Check
          
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = enCOLLIST.tcCHECK
                If .Value <> 1 Then GoTo OrdCheck1
                .Col = enCOLLIST.tcTESTCD:  SaveCode = .Value   '검사코드
                .Col = enCOLLIST.tcREQDTTM: SaveDate = .Value   '희망채혈일시
                .Col = enCOLLIST.tcSPCCD:   SaveSpc = .Value    '검체코드
                .Col = enCOLLIST.tcSTATFG:  SaveStatFg = .Value '응급여부
                
                For j = i + 1 To .DataRowCnt
                    .Row = j
                    .Col = enCOLLIST.tcCHECK
                    If .Value <> intSel Then GoTo OrdCheck2
                    .Col = enCOLLIST.tcORDNO:   strOrdNo = .Value
                    .Col = enCOLLIST.tcTESTCD
                    If .Value = SaveCode Then
                        .Col = enCOLLIST.tcREQDTTM
                        If .Value = SaveDate Then
                            .Col = enCOLLIST.tcSPCCD
                            If .Value = SaveSpc Then
                                .Col = enCOLLIST.tcSTATFG
                                If .Value = SaveStatFg Then
                                    .Col = enCOLLIST.tcORDDIV
                                    
                                    '-- 혈액은행은 중복처방 CHECK 불필요 By KJG
                                    If .Value <> BBS_ORDDIV Then
                                        If ChkBln = False Then
                                            strtmp = MsgBox("중복처방입니다. 계속 채취하시겠습니까?", vbCritical + vbOKCancel, "중복체크")
                                            If strtmp = vbCancel Then
                                                CheckSameOrder = j
                                                Exit Function
                                            Else
                                                ChkBln = True
                                            End If
                                        End If
                                        
                                       
                                        For k = j To .DataRowCnt
                                            .Row = k
                                            .Col = enCOLLIST.tcORDNO
                                            
                                            If .Value = strOrdNo Or Trim(.Value) = "" Then
                                            .Col = enCOLLIST.tcREQDTTM
                                                strReqDt = medGetP(.Value, 1, " ")
                                                strReqtm = Val(Mid(medGetP(.Value, 2, " "), 1, 2)) + 1
                                                strReqtm = strReqtm & Mid(medGetP(.Value, 2, " "), 3)
                                                strReqDt = strReqDt & " " & strReqtm
                                                .Value = strReqDt
                                            End If
                                        Next k
                                            
                                       
                                        
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
OrdCheck2:
                Next
OrdCheck1:
            Next
      
        ElseIf intSel = 2 Then    '지정검체 Check
      
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = 1
                If .Value <> 1 Then GoTo OrdCheck3
                .Col = enCOLLIST.tcSPCCD
                If Trim(.Value) = "" Then
                    CheckSameOrder = i
                    Exit Function
                End If
OrdCheck3:
            Next
        End If
    End With
            
End Function

Public Function Archive_WardColData(ByVal pDeptCd As String) As Boolean

    Dim strYesterday As String
    Dim SqlStmt As String
    Dim strtmp As String
    
    On Error GoTo Err_Trap
    
    strYesterday = Format(DateAdd("d", -30, GetSystemDate), CS_DateDbFormat)
    SqlStmt = objColSql.SqlArchiveWardColList(pDeptCd, strYesterday)
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    Archive_WardColData = True
    
    Exit Function
    
Err_Trap:
    Archive_WardColData = False
    MsgBox Err.Description, vbExclamation
End Function

'---------------------------------
' 전체 처방에대한 처방일 조회
'---------------------------------
Public Function LoadOrderDate(ByVal pPtid As String, ByRef objList As Object) As Long
    
    Dim Rs As Recordset
    Dim oDATE   As String
    Dim fDATE   As String
    Dim doday   As String
    Dim mDoCT   As String
    Dim jDoCT   As String
    
    Set Rs = New Recordset
    Rs.Open objColSql.SqlLoadOrderDate(pPtid, 0), DBConn
    
    If Rs.EOF Then
        LoadOrderDate = 0
    Else
        objList.Clear
        While (Not Rs.EOF)
            oDATE = Rs.Fields("orddt").Value & ""
            fDATE = Rs.Fields("bedindt").Value & ""
            doday = Rs.Fields("today").Value & ""
            mDoCT = Rs.Fields("orddoct").Value & ""
            jDoCT = Rs.Fields("majdoct").Value & ""
'            If fDATE > doday And mDoCT <> "" And jDoCT <> "" Then
'                MsgBox " 진료일자가 " & fDATE & " 인 처방은 미래처방이므로 접수가 불가능합니다 "
'            Else
                objList.AddItem Format(oDATE, CS_DateMask)
'            End If
'             objList.AddItem Format(Rs.Fields("orddt").Value, CS_DateMask)
            Rs.MoveNext
        Wend
        LoadOrderDate = objList.ListCount
        objList.ListIndex = 0
        DoEvents
    End If
    Set Rs = Nothing
    
End Function

'---------------------------------
' 수납 처방에대한 처방일 조회
'---------------------------------
Public Function LoadOrderDateYesPay(ByVal pPtid As String, ByRef objList As Object) As Long
' 2009.01.09 양성현 미래처방의 처리 현재일,REQDT, 처방의,주치의 추가
    Dim Rs As Recordset
    Dim oDATE   As String
    Dim fDATE   As String
    Dim doday   As String
    Dim mDoCT   As String
    Dim jDoCT   As String
    
    Set Rs = New Recordset
    Rs.Open objColSql.SqlLoadOrderDate(pPtid, 1), DBConn
    
    If Rs.EOF Then
        LoadOrderDateYesPay = 0
    Else
        objList.Clear
        While (Not Rs.EOF)
'            objList.AddItem Format(Rs.Fields("orddt").Value, CS_DateMask)
' 2009.01.09 양성현 미래처방의 처리
            oDATE = Rs.Fields("orddt").Value & ""
            fDATE = Rs.Fields("bedindt").Value & ""
            doday = Rs.Fields("today").Value & ""
            mDoCT = Rs.Fields("orddoct").Value & ""
            jDoCT = Rs.Fields("majdoct").Value & ""
'            If fDATE > doday And mDoCT <> "" And jDoCT <> "" Then
'                MsgBox " 진료일자가 " & fDATE & " 인 처방은 미래처방이므로 접수가 불가능합니다 "
'            Else
                objList.AddItem Format(oDATE, CS_DateMask)
'            End If
            
            Rs.MoveNext
        Wend
        LoadOrderDateYesPay = objList.ListCount
        If objList.ListCount > 0 Then objList.ListIndex = 0
        DoEvents
    End If
    Set Rs = Nothing
End Function
