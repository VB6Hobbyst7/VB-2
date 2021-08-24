Attribute VB_Name = "modVariableIF_YUMC"
Option Explicit

Public gfIFDisplayForm As Form

Public gsMachineCd As String
Public gsMachineNm As String
Public gsMachineExe As String
Public giTotIFItemCnt As Integer
Public giOriginIFItemCnt As Integer
Public giOriginCalItemCnt As Integer
Public gsOrdTestCdSeq As String
Public gsRstTestCdSeq As String
Public gsLastWSeq As String
Public giAddKey As Integer
Public giTestMode As Integer
Public gsClientServerMode As String
Public gsMachCd() As String
Public gsMachNm() As String

'For Server 등록 성공여부
Public giServerOK As Integer

'For MiddleWare
Public gObjMW As Object
Public gObjMW2 As Object

'인터페이스방식 모드
'0=단방향,
'1=양방향(Rack Or Tray 방식 지원안함, But Rack/Pos 표시)
'2=양방향(Rack Or Tray 방식 지원안함, But Tray/Pos 표시)
'3=양방향(Rack Or Tray 방식 지원안함, But Tray/Cup 표시)
'4=양방향(Rack/Pos 방식 지원),
'5=양방향(Tray/Pos 방식 지원),
'6=양방향(Tray/Cup 방식 지원),
Public gsIFMode$
Public gsINITMode$  'Initialize 버튼 사용 모드 - 0=사용안함, 1=사용함
Public gsTXMode$    '결과전송방식 모드 - 0=배치, 1=리얼타임(항목별 전송), 2=리얼타임(환자별 전송)
Public gsAPMode$    '자동출력 모드

Public gsIFVar1$, gsIFVar2$, gsIFVar3$, gsIFVar4$, gsIFVar5$

Public giBSRow%
Public giBERow%

'Comment관련
Public gCommentCd() As String
'Interface항목 검사항목코드 'H01001001NNNN
Public vIFItemCd() As Variant

Public iSpdBackColorOption As Integer

'ConverIFItemInfo시에 일치하는 개수
Public giMatchCnt%

Type CommTBL
    sPort As String
    sBaudRate As String
    sDataBit As String
    sStopBIt As String
    sParity As String
End Type

Type RackTBL
    sRackDigit As String
    sPosDigit As String
    sMaxRack As String
    sPosSetting As String
End Type

Type PosPerRackTBL
    sRackNo As String
    sPosMaxNo As String
End Type

Type IFItemTBL
    s01 As String:    s02 As String:    s03 As String
    s04 As String:    s05 As String:    s06 As String
    s07 As String:    s08 As String:    s09 As String
    s10 As String:    s11 As String:    s12 As String
    s13 As String:    s14 As String:    s15 As String
    s16 As String:    s17 As String:    s18 As String
    s19 As String:    s20 As String:    s21 As String
    s22 As String
End Type

Type CALITEMTBL
    s01 As String:    s02 As String:    s03 As String
    s04 As String:    s05 As String:    s06 As String
    s07 As String:    s08 As String:    s09 As String
    s10 As String:    s11 As String:    s12 As String
    s13 As String:    s14 As String:    s15 As String
    s16 As String
End Type

Type COMPIFSEQ
    sIFSeq() As String
    sResult() As String
    iSpdCol As Integer
End Type

Type CommentTBL
    CommentCd As String
    CommentNm As String
End Type

Type RSTTBL
    iCRow As Integer
    sLabInfo As String
    sSpcCd As String
    sTestCd(MAXIFITEM) As String
    sTRst As String
    sTRst2 As String
    sTAbl As String
    sTFlag As String
    sTCmt As String
End Type

Type ORDTBL
    iCRow As Integer
    sSampID As String
    sSampNo As String
    sIFSpcCd As String
    sOrdOpt As String
    iOrdCnt As Integer
    sIFSeq() As String
    sIFOrdCd() As String
    sServerCd() As String
    sIFRstCd() As String
    'IFRESULT
    sWDate As String
    sWSeq As String
    sJDate As String
    sJGbn As String
    sJNo As String
    sRack As String
    sPos As String
    sRegNo As String
    sName As String
    sSex As String
    sEmer As String
    sReRun As String
    sOther As String
End Type

Type ORDFIELDCFG
    sComponent As String
    sUse As String
    sStorageType As String
    sPath As String
    sFUse(MAXORDERFIELD) As String
    sFName(MAXORDERFIELD) As String
    sFOrd(MAXORDERFIELD) As String
    sFSize(MAXORDERFIELD) As String
End Type

Type RSTFIELDCFG
    sComponent As String
    sUse As String
    sStorageType As String
    sPath As String
    sFUse(MAXRESULTFIELD) As String
    sFName(MAXRESULTFIELD) As String
    sFOrd(MAXRESULTFIELD) As String
    sFSize(MAXRESULTFIELD) As String
End Type

Type USEFIELDCFG
    sSeq As String
    sFDispOrd As String
    sFOrd As String
    sFSize As String
    sText As String
End Type

Public gIFItem() As IFItemTBL
Public gCalItem() As CALITEMTBL
Public gIFRack As RackTBL
Public gCommInfo As CommTBL
Public gIFPosInfo() As PosPerRackTBL
Public gResultTable() As RSTTBL
Public gOrderTable As ORDTBL
Public gCommentTable() As CommentTBL
Public gOrdCfg As ORDFIELDCFG
Public gRstcfg As RSTFIELDCFG
Public gUseFieldCfg() As USEFIELDCFG

'
' 2002-05-24 JJH 추가
' 워크리스트내용병합 ( 조건: 같은병록번호, 결과받지않은것, 검사항목 중복되지않을 때 )
'
Public Function WMerge_SPD() As Long
    On Error GoTo ErrHandler
    
    Dim i%, j%
    Dim iDone%
    
    Dim sTmp1     As String
    Dim sTmp2     As String
    Dim sChk      As String
    Dim iAllCnt   As Integer
    
    Dim sWSEQ1    As String
    Dim sWSEQ2    As String
    
    Dim sOrdCd1() As String
    Dim sOrdCd2() As String
    
    Dim objld As Object
    
    iDone = 0
    
    With gfIFDisplayForm
        With .spdIntList
            If .IsBlockSelected = True Then
                
                '-- 선택한 데이타Row 체크
                If giBERow - giBSRow + 1 > 2 Then
                    MsgBox "3개 이상은 병합할 수 없습니다.", vbExclamation
                    Exit Function
                End If
                
                '--
                If giBERow <> .MaxRows Then
                    MsgBox "병합할 때 다음 리스트가 존재하면 안됩니다.", vbExclamation
                    Exit Function
                End If
                
                '--
                If MsgBox("선택된 데이타를 병합하시겠습니까?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Function
                End If
                
                '-- 작업번호
                .Row = giBSRow: .Col = 1: sWSEQ1 = Trim(.Text)
                .Row = giBERow: .Col = 1: sWSEQ2 = Trim(.Text)
                
                '-- 병록번호 체크
                .Row = giBSRow: .Col = 8: sTmp1 = Trim(.Text)
                .Row = giBERow: .Col = 8: sTmp2 = Trim(.Text)
                If sTmp1 <> sTmp2 Then
                    MsgBox "병록번호가 틀리므로 병합할 수 없습니다.", vbExclamation
                    Exit Function
                End If
                
                '-- 결과전송여부체크
                .Row = giBSRow: .Col = 15: sTmp1 = Trim(.Text)
                .Row = giBERow: .Col = 15: sTmp2 = .Text
                If sTmp1 <> "N" Or sTmp2 <> "N" Then
                    MsgBox "이미 결과전송이 완료된 데이타이므로 병합할 수 없습니다.", vbExclamation
                End If
                
                '-- 검사항목 중복체크
                .Row = giBSRow: .Col = 14: ReDim sOrdCd1(Val(Trim(.Text))) As String
                .Row = giBERow: .Col = 14: ReDim sOrdCd2(Val(Trim(.Text))) As String
                For i = 1 To UBound(sOrdCd1)
                    .Row = giBSRow: .Col = 16 + i: sOrdCd1(i) = Trim(.Text)
                Next
                    
                For i = 1 To UBound(sOrdCd2)
                    .Row = giBERow: .Col = 16 + i: sOrdCd2(i) = Trim(.Text)
                    
                    For j = 1 To UBound(sOrdCd1)
                        If sOrdCd1(j) = sOrdCd2(i) Then
                            MsgBox "검사항목이 중복되어 병합할 수 없습니다.", vbExclamation
                            Exit Function
                        End If
                    Next
                Next
                
                iAllCnt = UBound(sOrdCd1) + UBound(sOrdCd2)
                .Row = giBSRow: .Col = 14: .Text = CStr(iAllCnt)
                .Row = giBSRow: .Col = 16: .Text = CStr(iAllCnt)
                
                '검사항목 정보 숨기기
                For i = 1 To UBound(sOrdCd2)
                    .Row = giBSRow: .Col = 16 + i + UBound(sOrdCd1)
                    .Text = sOrdCd2(i)
                Next
                
                .Row = giBERow
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
                
                .Action = SS_ACTION_DESELECT_BLOCK
                
                iDone = 1
            End If
        End With
        
        '-- MDB Update
        Set objld = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
        Call objld.WMerge_MDB(gsMachineCd, Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD"), sWSEQ1, sWSEQ2)
        Set objld = Nothing

        If iDone = 1 Then
            .Hide: .Show
        End If
    
    End With

    Exit Function
    
ErrHandler:
    ViewMsg "WMerge_SPD 오류 - (" & Err.Description & ")"
End Function

Public Function AddIFList(ByVal sWDate As String, ByVal sWSeq As String, _
                                ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                                ByVal iRstCnt As Integer, ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                                ByVal sIFSpcCd As String, ByVal sCurRow As String) As String
    Dim i%
    Dim vWSeq, vCRstCnt
    Dim sCIFSeq$, sCIFRstCd$, sCRst1$, sCRst2$
    ReDim gResultTable(1)
    
    AddIFList = ""
    
    With gfIFDisplayForm
        If Len(sWDate) = 0 Then
            sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")
        Else
        End If
        
        If Len(sWSeq) = 0 Then
            With .spdIntList
                sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
                
                AddIFList = sWSeq
            End With
        Else
        End If
        
        With .spdIntList
            .MaxRows = .MaxRows + 1
            
            Call .SetText(1, .MaxRows, sWSeq & "")
            Call .SetText(2, .MaxRows, "1")     '체크
            Call .SetText(3, .MaxRows, sJDate & "")
            Call .SetText(4, .MaxRows, sJGbn & "")
            Call .SetText(5, .MaxRows, sJNo & "")
            Call .SetText(6, .MaxRows, "ADD")      'Rack
            Call .SetText(7, .MaxRows, "NO")      'Pos
            Call .SetText(8, .MaxRows, "")      'RegNo
            Call .SetText(9, .MaxRows, "")      'Name
            Call .SetText(10, .MaxRows, "")      'Sex
            Call .SetText(11, .MaxRows, "")      'Emer
            Call .SetText(12, .MaxRows, "")      'ReRun
            Call .SetText(13, .MaxRows, "")      'Other
            Call .SetText(14, .MaxRows, "N")     'Order
            
            Call .SetText(16, .MaxRows, CStr(iRstCnt) & "")      'IFCnt
            
            For i = 1 To iRstCnt
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                
                sCIFSeq = ConvertIFItemInfo(7, sCIFRstCd)
                
                If sCIFSeq = "" Then
                Else
                    Call .GetText(15, .MaxRows, vCRstCnt)      'Result
                    Call .SetText(15, .MaxRows, CStr(Val(vCRstCnt) + 1) & "")
                    
                    Call .SetText(16 + i, .MaxRows, sCIFSeq & "|" & sCRst1 & "|" & sCRst2 & "|")
                End If
            Next
            
            Call .GetText(15, .MaxRows, vCRstCnt)
            
            If Val(vCRstCnt) = 0 Then
                .MaxRows = .MaxRows - 1
            End If
            
            '현재 Row 기록
            gResultTable(1).iCRow = .MaxRows
        End With
        
        If Len(sJDate) = 0 Then
            sJDate = ""
        Else
            sJDate = sJDate & "-"
        End If
        
        If Len(sJGbn) = 0 Then
            sJGbn = ""
        Else
            sJGbn = sJGbn & "-"
        End If
        
        If Len(sJNo) > 0 Then
            .lblResult = sJDate & sJGbn & sJNo
        Else
            .lblResult = sWDate & "-" & sWSeq
        End If
    End With
End Function

Public Sub ClearAllList()
    Dim i%
    
    With gfIFDisplayForm
        If MsgBox("Interface List를 삭제하면 해당 List의 결과를 받지 못합니다." & vbCrLf & vbCrLf & _
            "결과를 아직 받지 않았다면 '아니오'를 선택하십시요." & vbCrLf & _
            "Interface List를 정말 삭제하시겠습니까?", vbYesNo, "전체리스트 화면 삭제 확인") = vbYes Then
            
            .spdIntList.MaxRows = 0
            
            With .spdRst
                .BlockMode = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = -1
                .Col2 = -1
                .Action = SS_ACTION_CLEAR_TEXT
                .BlockMode = False
            End With
            
            With .spdRst2
                .BlockMode = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = -1
                .Col2 = -1
                .Action = SS_ACTION_CLEAR_TEXT
                .BlockMode = False
            End With
            
            .lblResult.Caption = ""
            .lblOrder.Caption = ""
            .lblCSelList.Caption = ""
            
            .Hide
            .Show
        End If
    End With
End Sub

Public Sub ClearBlockedList()
    Dim i%
    Dim iDone%
    
    iDone = 0
    
    With gfIFDisplayForm
        With .spdIntList
            If .IsBlockSelected = True Then
                For i = giBSRow To giBERow
                    .Row = giBSRow
                    .Action = SS_ACTION_DELETE_ROW
                    
                    .MaxRows = .MaxRows - 1
                Next
                
                .Action = SS_ACTION_DESELECT_BLOCK
                
                iDone = 1
            End If
        End With
        
        If iDone = 1 Then
            .Hide
            .Show
        End If
    End With
End Sub

Public Sub DisplayAfterSendOrder()
    If gOrderTable.iCRow < 1 Then Exit Sub
    
    With gfIFDisplayForm
        With .spdIntList
            Call .SetText(2, gOrderTable.iCRow, CVar("0"))
            
            Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, gOrderTable.iCRow, gOrderTable.iCRow, _
                    RGB(0, 0, 0), 연노랑)
        End With
        
        .lblOrder = gOrderTable.sSampID
    End With
    
    'gOrderTable 초기화
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
End Sub

Public Function CFCompute(ByVal strInFormula As String, Optional ByRef nState As Variant) As Double
    Dim i               As Integer
    Dim j               As Integer
    Dim nStartPos       As Integer
    Dim strChar         As String
    Dim strFormula(99)  As String
    Dim nCnt            As Integer
    Dim strFormula2(99) As String
    Dim nCnt2           As Integer
    Dim nFlag           As Integer
    Dim nFlag2          As Integer
    Dim nFlag3          As Integer
    Dim nCurrPos        As Integer
    Dim nLeftPos        As Integer
    Dim nRightPos       As Integer
    Dim nLevel          As Integer
    Dim nOpLevel        As Integer
    Dim nOldCnt         As Integer
    Dim nOldStartPos    As Integer
    Dim nTop            As Integer  ' Stack Pointer
    
    If Not IsMissing(nState) Then nState = True
    ' 숫자, 연산자, 괄호 등을 나누어 배열에 저장
    nCnt = 0        ' 배열에 저장된 갯수
    nStartPos = 0   ' 숫자를 하나로 보기위해 숫자 시작위치 저장
    
    ' '**'를 '^'로 바꿈
    Do While InStr(strInFormula, "**") > 0
        nCurrPos = InStr(strInFormula, "**")
        strInFormula = Left(strInFormula, nCurrPos - 1) & "^" & Mid(strInFormula, nCurrPos + 2)
    Loop
    
    ' 'MOD'를 '%'로 바꿈
    Do While InStr(UCase(strInFormula), "MOD") > 0
        nCurrPos = InStr(UCase(strInFormula), "MOD")
        strInFormula = Left(strInFormula, nCurrPos - 1) & "%" & Mid(strInFormula, nCurrPos + 3)
    Loop
    
    nFlag = False   ' 수식이 정상적인 여부 (특수문자등)
    
    For i = 1 To Len(strInFormula)
        strChar = Mid$(strInFormula, i, 1)   ' 한글자씩 숫자인지 연산자인지 괄호인지 비교
        If Trim(strChar) <> "" Then
            If IsNumeric(strChar) Or (strChar = ".") Then   ' 숫자와 소수점만 수자로 취급
                If nStartPos = 0 Then nStartPos = i
                nFlag = True
            Else
                If nStartPos > 0 Then   ' 숫자 다음 숫자 외의 문자가 나올 경우 숫자를 저장
                    strFormula(nCnt) = Mid$(strInFormula, nStartPos, i - nStartPos)
                    nCnt = nCnt + 1
                    nStartPos = 0
                End If
                If (strChar Like "[()]") Or IsOp(strChar) Then
                    ' 괄호 및 연산자를 저장
                    strFormula(nCnt) = strChar
                    nCnt = nCnt + 1
                    nFlag = True
                End If
            End If
            If nFlag = True Then    ' 수치, 괄호, 연산자 외의 이상한 문자가 있는지 확인
                nFlag = False
            Else
                GoTo Err_Process
            End If
        End If
    Next i
    
    If nStartPos > 0 Then   ' 숫자가 마지막 으로 끝난 경우 숫자를 저장
        strFormula(nCnt) = Mid$(strInFormula, nStartPos, i - nStartPos)
        nCnt = nCnt + 1
        nStartPos = 0
    End If
    
    ' 부호(-)를 제거한다. ('(', '연산자' 다음에 나오는 '+', '-'는 부호임.)
    nFlag = True    ' '(', '연산자'가 나왔는지 여부
    
    For i = 0 To nCnt - 1
        If nFlag = True Then
            If strFormula(i) Like "[+-]" Then      ' 부호 발견
                If strFormula(i) = "-" Then
                    If IsNumeric(strFormula(i + 1)) Then
                        ' 부호(-)를 다음에 나오는 숫자에 포함.
                        strFormula(i + 1) = Trim(Str(Val(strFormula(i + 1)) * -1))
                    Else
                        '부호 다음에 연산자가 나옴
                        GoTo Err_Process
                    End If
                End If
                strFormula(i) = ""     ' 부호가 있던 자리를 Null로 체움
            End If
        End If
        ' '(', '연산자' 다음에 나오는 '+', '-'는 부호이므로 '(', '연산자' 확인
        If IsOp(strFormula(i)) Then
            nFlag = True
        Else
            nFlag = False
        End If
    Next i
    
    ' 부호(-)를 제거할때 발생한 Null 제거 (strFormula2로 옮긴후 다시 strFormula로 옮김)
    nCnt2 = 0
    
    For i = 0 To nCnt - 1
        If Trim(strFormula(i)) <> "" Then
            strFormula2(nCnt2) = strFormula(i)
            nCnt2 = nCnt2 + 1
        End If
    Next i
    
    nCnt = nCnt2
    
    For i = 0 To nCnt - 1
        strFormula(i) = strFormula2(i)
    Next i
    
    ' 식에 연산자가 없으면 모든 괄호 제거 (예:(1))
    For i = 0 To nCnt - 1
        If IsOp(strFormula(i)) Then
            Exit For
        End If
    Next i
    
    If i = nCnt Then
        ' 불필요한 괄호 제거
        For i = 0 To nCnt - 1
            If strFormula(i) Like "[()]" Then
                strFormula(i) = ""
            End If
        Next i
        
        ' 불필요한 괄호를 제거할때 발생한 Null 제거 (strFormula2로 옮긴후 다시 strFormula로 옮김)
        nCnt2 = 0
        
        For i = 0 To nCnt - 1
            If Trim(strFormula(i)) <> "" Then
                strFormula2(nCnt2) = strFormula(i)
                nCnt2 = nCnt2 + 1
            End If
        Next i
        
        nCnt = nCnt2
        
        For i = 0 To nCnt - 1
            strFormula(i) = strFormula2(i)
        Next i
    End If
    
    ' 괄호가 있는지 확인 및 우선 순위가 높은 연산자를 찾아서 괄호로 묶는다. (의미없는 괄호는 없엔다.)
    nStartPos = 0       ' 우선 순위 비교 시작 위치
    
    Do
        nFlag = True    ' 작업 중지 Flag
        nFlag2 = False  ' 괄호 삽입 Flag
        nFlag3 = False  ' nStartPos 변경여부
        nOldStartPos = nStartPos
        
        For i = nStartPos To nCnt - 1
            If IsOp(strFormula(i)) Then
                nOpLevel = Check_OpLevel(strFormula(i))    ' 연산 우선순위 Level
                nCurrPos = i
                nLevel = 0          ' 괄호로 묶인 내용을 하나로 보기위해 괄호를 열고 닫은 수
                
                For j = (i - 1) To 0 Step -1    ' 연산자의 좌측에 괄호가 들어갈 위치 확인
                    If nLevel = 0 Then
                        If IsOp(strFormula(j)) Then
                            nLeftPos = j + 1    ' 괄호가 삽입될 위치
                            Exit For
                        End If
                    End If
                    If strFormula(j) = ")" Then nLevel = nLevel + 1
                    If strFormula(j) = "(" Then nLevel = nLevel - 1
                Next j
                
                If j = -1 Then nLeftPos = 0     ' 괄호가 삽입될 위치
                
                nLevel = 0          ' 괄호로 묶인 내용을 하나로 보기위해 괄호를 열고 닫은 수
                
                For j = (i + 1) To (nCnt - 1)   ' 연산자의 우측에 우선순위가 더 높은 연산자가 있는지 확인
                    If nLevel = 0 Then
                        If IsOp(strFormula(j)) Then
                            If nOpLevel >= Check_OpLevel(strFormula(j)) Then
                                ' 우측의 연산자가 우선순위가 높지 않을 경우 현재 연산자를 괄호로 묶는다.
                                nRightPos = j    ' 괄호가 삽입될 위치
                                nFlag2 = True
                                nFlag3 = True
                                nFlag = False:  Exit For
                            Else
                                ' 우측의 연산자가 우선순위가 높으므로 현위치(J)에서 다시비교
                                nStartPos = j
                                nFlag3 = True
                                nFlag = False:  Exit For
                            End If
                        End If
                        If strFormula(j) = ")" Then
                            nRightPos = j    ' 괄호가 삽입될 위치
                            nFlag2 = True
                            nFlag3 = True
                            nFlag = False:  Exit For
                        End If
                    End If
                    If strFormula(j) = "(" Then nLevel = nLevel + 1
                    If strFormula(j) = ")" Then nLevel = nLevel - 1
                Next j
                
                If j = nCnt Then
                    nRightPos = nCnt   ' 괄호가 삽입될 위치
                    nFlag2 = True
                    Exit For
                End If
                
                If nFlag = False Then Exit For
            End If
        Next i
        
        nOldCnt = nCnt
        
        If nFlag2 = True Then
            If Not (strFormula(IIf(nLeftPos = 0, 0, nLeftPos - 1)) = "(" And strFormula(nRightPos) = ")") Then
                ' 괄호 삽입
                nCnt2 = 0
                For i = 0 To nLeftPos - 1
                    strFormula2(nCnt2) = strFormula(i)
                    nCnt2 = nCnt2 + 1
                Next i
                strFormula2(nCnt2) = "("
                nCnt2 = nCnt2 + 1
                For i = nLeftPos To nRightPos - 1
                    strFormula2(nCnt2) = strFormula(i)
                    nCnt2 = nCnt2 + 1
                Next i
                strFormula2(nCnt2) = ")"
                nCnt2 = nCnt2 + 1
                For i = nRightPos To nCnt - 1
                    strFormula2(nCnt2) = strFormula(i)
                    nCnt2 = nCnt2 + 1
                Next i
                nCnt = nCnt2
                For i = 0 To nCnt - 1
                    strFormula(i) = strFormula2(i)
                Next i
                nFlag2 = False
            End If
            nStartPos = nRightPos
          
        End If
        If nFlag3 = True Then
            nFlag3 = False
        Else
            ' 연산자 우선순위에 밀려 괄호로 묶이지 않은 연산자가 앞에 있을 수 있으므로
            If nOldStartPos <> 0 Then nStartPos = 0
        End If
    Loop Until (nLeftPos = 0 Or nLeftPos = 1) And (nRightPos = nOldCnt Or nRightPos = nOldCnt - 1)
    
    ' PreFix 로 바꾼다.
    For i = 0 To nCnt - 1
        If strFormula(i) = "(" Then
            nLevel = 0          ' 괄호로 묶인 내용을 하나로 보기위해 괄호를 열고 닫은 수
            strChar = ""
            nFlag = False
            For j = i + 1 To nCnt - 1
                If nLevel = 0 Then
                    If IsOp(strFormula(j)) Then
                        nCurrPos = j
                        strChar = strFormula(j)
                        nFlag = True
                        Exit For
                    End If
                End If
                If strFormula(j) = "(" Then nLevel = nLevel + 1
                If strFormula(j) = ")" Then nLevel = nLevel - 1
                If nLevel = -1 Then Exit For    ' 괄호안에 연산자 없음
            Next j
            If nFlag = True Then
                If Trim(strChar) <> "" Then
                    strFormula(i) = strChar    ' 괄호('(')를 연산자로 교체
                    strFormula(j) = ""
                End If
            End If
        End If
    Next i
    
    ' 불필요한 괄호 제거
    For i = 0 To nCnt - 1
        If strFormula(i) Like "[()]" Then
            strFormula(i) = ""
        End If
    Next i
    
    ' 불필요한 괄호를 제거할때 발생한 Null 제거 (strFormula2로 옮긴후 다시 strFormula로 옮김)
    nCnt2 = 0
    For i = 0 To nCnt - 1
        If Trim(strFormula(i)) <> "" Then
            strFormula2(nCnt2) = strFormula(i)
            nCnt2 = nCnt2 + 1
        End If
    Next i
    nCnt = nCnt2
    For i = 0 To nCnt - 1
        strFormula(i) = strFormula2(i)
    Next i
    
    ' 스텍에 넣으면서 계산
    nTop = -1
    If nCnt = 1 Then
        nTop = 0
        strFormula2(nTop) = strFormula(0)
    ElseIf nCnt = 2 Then
        GoTo Err_Process
    ElseIf nCnt > 2 Then
        For i = 0 To nCnt - 1
            nTop = nTop + 1
            strFormula2(nTop) = strFormula(i)
            If nTop >= 2 Then
                Do While IsNumeric(strFormula2(nTop - 1)) And IsNumeric(strFormula2(nTop))
                    ' 계산
                    If (strFormula2(nTop - 2) = "/" Or strFormula2(nTop - 2) = "\" Or strFormula2(nTop - 2) = "%") And Val(strFormula2(nTop)) = 0 Then
                        GoTo Err_Process
                    End If
                    strFormula2(nTop - 2) = Trim(Str(SubCompute(strFormula2(nTop - 2), Val(strFormula2(nTop - 1)), Val(strFormula2(nTop)))))
                    nTop = nTop - 2
                    If nTop < 2 Then Exit Do
                Loop
            End If
        Next i
    End If
    If nTop <> 0 Then GoTo Err_Process
    
    CFCompute = Val(strFormula2(0))
    
Exit Function

'/----------------------------------------------------------/

Err_Process:
    ViewMsg "계산식에 오류가 있습니다."
    
    If Not IsMissing(nState) Then nState = False
    
    Exit Function
    
End Function

Public Function ChkCalResult(ByVal iCRow As Integer, iRstCnt As Integer, sIFRstCd As String, sRst1 As String, sRst2 As String, sIFSpcCd As String) As String
    Dim i%, j%, k%, iPos%, iSPos%, iCnt%, iExist%, iAlready%
    Dim sCIFRstCd$, sCRst$, sCIFSeq$, sTmp$, sCF$, sCRst2$
    Dim vSex, vTmp, vIFCnt, vCRstCnt
    Dim sCompIFSeq As COMPIFSEQ
        
    With gfIFDisplayForm
        With .spdIntList
            Call .GetText(10, iCRow, vSex)
            
            Select Case vSex
                Case "M", "남", "1"
                    vSex = "M"
                Case "F", "여", "2"
                    vSex = "F"
                Case Else
                    vSex = "M"
            End Select
            
            Call .GetText(16, iCRow, vIFCnt)
            
            For j = 1 To CInt(Val(vIFCnt))
                Call .GetText(16 + j, iCRow, vTmp)
                
                sTmp = CStr(vTmp)
                sCIFSeq = GetByOne(sTmp, sTmp)
                gResultTable(1).sTestCd(j) = sCIFSeq
            Next
            
            For i = 1 To giOriginCalItemCnt
                iCnt = 0
                iExist = 0
                
                sCF = gCalItem(i).s04
                
                iSPos = 1
                
                Do
                    '알파벳 I(대문자 아이)를 찾는다.
                    iPos = InStr(iSPos, sCF, "I")
                    
                    If iPos = 0 Then
                        Exit Do
                    Else
                        iCnt = iCnt + 1
                        ReDim Preserve sCompIFSeq.sIFSeq(iCnt)
                        ReDim Preserve sCompIFSeq.sResult(iCnt)
                        sCompIFSeq.sIFSeq(iCnt) = Mid(sCF, iPos + 1, 3)
                                               
                        iSPos = iPos + 1
                    End If
                Loop
                
                '계산식에 필요한 Interface 결과가 전송되었는지 체크
                For j = 1 To iCnt
                    For k = 1 To CInt(Val(vIFCnt))
                        If gResultTable(1).sTestCd(k) = sCompIFSeq.sIFSeq(j) Then
                            Call .GetText(16 + k, iCRow, vTmp)
                            
                            sTmp = CStr(vTmp)
                            
                            Call GetByOne(sTmp, sTmp)
                            sCRst = GetByOne(sTmp, sTmp)
                            
                            sCompIFSeq.sResult(j) = sCRst
                            
                            If sCRst = "" Or IsNumeric(sCRst) = False Then
                            Else
                                iExist = iExist + 1
                            End If
                        End If
                    Next
                Next
                
                '계산을 위한 결과가 모두 전송되었다면
                '계산 결과를 스프레드에 나타냄
                If iCnt = iExist Then
                    For k = 1 To iCnt
                        sCF = Replace(sCF, "I" & sCompIFSeq.sIFSeq(k), sCompIFSeq.sResult(k))
                    Next
                    
                    sCRst = CFCompute(sCF)
                    
                    If Left(sCRst, 1) = "-" Then
                        sCRst = ConvertResult1("-", "0", sCRst, gCalItem(i).s01)
                    Else
                        sCRst = ConvertResult1("+", "0", sCRst, gCalItem(i).s01)
                    End If
                    
                    '계산 항목은 IFSEQ("C1"과 같은) 사용
                    sCRst = JudgeResultBySex(gCalItem(i).s01, sCRst, vSex, "", "", sCRst2, "", "")
                    
                    '이전에 전송되었는지 체크
                    Call .GetText(15, iCRow, vCRstCnt)
                    Call .GetText(16, iCRow, vIFCnt)
                    
                    For k = 1 To vIFCnt
                        Call .GetText(16 + k, iCRow, vTmp)
                        
                        sTmp = vTmp
                        
                        sTmp = GetByOne(sTmp, sTmp)
                        
                        If sTmp = gCalItem(i).s01 Then
                            sCompIFSeq.iSpdCol = k
                            iAlready = 1
                            Exit For
                        End If
                    Next
                    
                    If iAlready = 1 Then
                        Call .SetText(16 + sCompIFSeq.iSpdCol, iCRow, gCalItem(i).s01 & "|" & sCRst & "|" & sCRst2 & "|")
                    Else
                        Call .SetText(15, iCRow, CInt(Val(vCRstCnt) + 1) & "")
                        Call .SetText(16, iCRow, CInt(Val(vIFCnt) + 1) & "")
                        
                        Call .SetText(16 + CInt(Val(vIFCnt) + 1), iCRow, gCalItem(i).s01 & "|" & sCRst & "|" & sCRst2 & "|")
                    End If
                    
                    '결과등록을 위해 결과등록 파라미터 변환
                    iRstCnt = iRstCnt + 1
                    sIFRstCd = sIFRstCd & gCalItem(i).s01 & "|"
                    sRst1 = sRst1 & sCRst & "|"
                    'sCRst2는 JudgeResultBySex에서 받아 옴
                    sRst2 = sRst2 & sCRst2 & "|"
                    
                    If Len(sIFSpcCd) = 0 Then
                    Else
                        sIFSpcCd = sIFSpcCd & "|"
                    End If
                    
                    ChkCalResult = "1"
                End If
            Next
        End With
    End With
End Function

Public Function Check_OpLevel(ByVal strOp As String) As Integer

    Check_OpLevel = 0
    
    Select Case strOp
        Case "+", "-":      Check_OpLevel = 1
        Case "\", "%":      Check_OpLevel = 2
        Case "*", "/":      Check_OpLevel = 3
        Case "^":           Check_OpLevel = 4
    End Select

End Function

Public Function ConvertIFItemInfo2(ByVal iMode As Integer, ByVal sComp As String) As String
    Dim i%
    
    Select Case iMode
        '서버쪽코드를 IFSEQ로
        Case 1
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s06 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s01
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s03 = sComp Then
                    ConvertIFItemInfo2 = gCalItem(i).s01
                    Exit For
                End If
            Next
            
        'IFSEQ를 서버쪽코드로
        Case 2
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s06
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gCalItem(i).s03
                    Exit For
                End If
            Next
            
        '검사항목명을 IFSEQ로
        Case 3
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s02 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s01
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s02 = sComp Then
                    ConvertIFItemInfo2 = gCalItem(i).s01
                    Exit For
                End If
            Next
            
        'IFSEQ를 검사항목명으로
        Case 4
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s02
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gCalItem(i).s02
                    Exit For
                End If
            Next
        
        'IFORDCD를 IFSEQ로
        Case 5
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s01
                    Exit For
                End If
            Next
                  
        'IFSEQ를 IFORDCD로
        Case 6
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s03
                    Exit For
                End If
            Next
        
        'IFRSTCD를 IFSEQ로
        Case 7
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s04 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s01
                    Exit For
                End If
            Next
        
        'IFSEQ를 IFRSTCD로
        Case 8
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s04
                End If
            Next
        
        Case Else
        
    End Select
End Function

Public Function ConvertIFItemInfo(ByVal iMode As Integer, ByVal sComp As String) As String
    Dim i%
    
    Select Case iMode
        '서버쪽코드를 IFSEQ로
        Case 1
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s06 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gCalItem(i).s01
                    Exit For
                End If
            Next
            
        'IFSEQ를 서버쪽코드로
        Case 2
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s06
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gCalItem(i).s03
                    Exit For
                End If
            Next
            
        '검사항목명을 IFSEQ로
        Case 3
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s02 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s02 = sComp Then
                    ConvertIFItemInfo = gCalItem(i).s01
                    Exit For
                End If
            Next
            
        'IFSEQ를 검사항목명으로
        Case 4
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s02
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gCalItem(i).s02
                    Exit For
                End If
            Next
        
        'IFORDCD를 IFSEQ로
        Case 5
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
                  
        'IFSEQ를 IFORDCD로
        Case 6
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s03
                    Exit For
                End If
            Next
        
        'IFRSTCD를 IFSEQ로
        Case 7
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s04 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
        
        'IFSEQ를 IFRSTCD로
        Case 8
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s04
                    Exit For
                End If
            Next
        
        'IFORDCD를 IFSPECIMEN로
        Case 9
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s05
                End If
            Next
            
        'IFORDCD를 판정구분으로
        Case 10
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s09
                End If
            Next
        
        'IFSEQ를 IFSPCECIMEN로
        Case 11
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s05
                End If
            Next
            
        Case Else
        
    End Select
End Function

Public Function NewIFList(ByVal sWDate As String, ByVal sWSeq As String, _
                                ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                                ByVal sRack As String, ByVal sPos As String, _
                                ByVal sRegNo As String, ByVal sName As String, _
                                ByVal sSex As String, ByVal sEmer As String, ByVal sReRun As String, ByVal sOther As String, _
                                ByVal iRstCnt As Integer, ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                                ByVal sIFSpcCd As String, ByVal sCurRow As String, ByVal sFlag As String) As String
    On Error GoTo ErrRtn
    
    Dim i%
    Dim vWSeq, vCRstCnt
    Dim sCIFSeq$, sCIFRstCd$, sCRst1$, sCRst2$, sCFlag$
    Dim vSex
    
    NewIFList = ""
    
    With gfIFDisplayForm
        If Len(sWDate) = 0 Then
            sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")
        Else
        End If
        
        If Len(sWSeq) = 0 Then
            With .spdIntList
                sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
                
                NewIFList = sWSeq
            End With
        Else
        End If
        
        With .spdIntList
            .MaxRows = .MaxRows + 1
            
            Call .SetText(1, .MaxRows, sWSeq & "")
            'NewIFList와 OldIFList의 차이 = 체크 O, X
            Call .SetText(2, .MaxRows, "1")     '체크
            '<--- 나중에 데이터 삭제시에 편리함..
            Call .SetText(3, .MaxRows, sJDate & "")
            Call .SetText(4, .MaxRows, sJGbn & "")
            Call .SetText(5, .MaxRows, sJNo & "")
            Call .SetText(6, .MaxRows, sRack & "")      'Rack
            Call .SetText(7, .MaxRows, sPos & "")       'Pos
            Call .SetText(8, .MaxRows, sRegNo & "")     'RegNo
            Call .SetText(9, .MaxRows, sName & "")      'Name
            Call .SetText(10, .MaxRows, sSex & "")      'Sex
            Call .SetText(11, .MaxRows, sEmer & "")     'Emer
            Call .SetText(12, .MaxRows, sReRun & "")    'ReRun
            Call .SetText(13, .MaxRows, sOther & "")    'Other
            Call .SetText(14, .MaxRows, "N")     'Order
            
            Call .SetText(16, .MaxRows, CStr(iRstCnt) & "")      'IFCnt
            
            For i = 1 To iRstCnt
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                sCFlag = GetByOne(sFlag, sFlag)
                
                sCIFSeq = ConvertIFItemInfo(7, sCIFRstCd)
                
                If sCIFSeq = "" Then
                Else
                    Call .GetText(15, .MaxRows, vCRstCnt)      'Result
                    Call .SetText(15, .MaxRows, CStr(Val(vCRstCnt) + 1) & "")
                    
                    Call .SetText(16 + Val(vCRstCnt) + 1, .MaxRows, _
                                sCIFSeq & "|" & sCRst1 & "|" & sCRst2 & "|" & sCFlag & "|")
                End If
            Next
            
            Call .GetText(15, .MaxRows, vCRstCnt)
            
            If Val(vCRstCnt) = 0 Then
                .MaxRows = .MaxRows - 1
            End If
            
            '현재 Row 기록
            gResultTable(1).iCRow = .MaxRows
        End With
        
        If Len(sJDate) = 0 Then
            sJDate = ""
        Else
            sJDate = sJDate & "-"
        End If
        
        If Len(sJGbn) = 0 Then
            sJGbn = ""
        Else
            sJGbn = sJGbn & "-"
        End If
        
        If Len(sJNo) > 0 Then
            .lblResult = sJDate & sJGbn & sJNo
        Else
            .lblResult = sWDate & "-" & sWSeq
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        NewIFList = "NO"
        ViewMsg "NewIFList - Err(" & Err.Description & ")"
    End If
End Function

Public Function NewIFListBySex(ByVal sWDate As String, ByVal sWSeq As String, _
                                ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                                ByVal sRack As String, ByVal sPos As String, _
                                ByVal sRegNo As String, ByVal sName As String, _
                                ByVal sSex As String, ByVal sEmer As String, ByVal sReRun As String, ByVal sOther As String, _
                                ByVal iRstCnt As Integer, ByVal sIFRstCd As String, sRst1 As String, sRst2 As String, _
                                ByVal sIFSpcCd As String, ByVal sCurRow As String) As String
    Dim i%
    Dim vWSeq, vCRstCnt, vSex
    Dim sCIFSeq$, sCIFRstCd$, sCRst1$, sCRst2$
    Dim sNRst1$, sNRst2$
    
    NewIFListBySex = ""
    
    With gfIFDisplayForm
        If Len(sWDate) = 0 Then
            sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")
        Else
        End If
        
        If Len(sWSeq) = 0 Then
            With .spdIntList
                sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
                
                NewIFListBySex = sWSeq
            End With
        Else
        End If
        
        With .spdIntList
            .MaxRows = .MaxRows + 1
            
            Call .SetText(1, .MaxRows, sWSeq & "")
            Call .SetText(2, .MaxRows, CVar("0"))
            Call .SetText(3, .MaxRows, sJDate & "")
            Call .SetText(4, .MaxRows, sJGbn & "")
            Call .SetText(5, .MaxRows, sJNo & "")
            Call .SetText(6, .MaxRows, sRack & "")      'Rack
            Call .SetText(7, .MaxRows, sPos & "")       'Pos
            Call .SetText(8, .MaxRows, sRegNo & "")     'RegNo
            Call .SetText(9, .MaxRows, sName & "")      'Name
            Call .SetText(10, .MaxRows, sSex & "")      'Sex
            Call .SetText(11, .MaxRows, sEmer & "")     'Emer
            Call .SetText(12, .MaxRows, sReRun & "")    'ReRun
            Call .SetText(13, .MaxRows, sOther & "")    'Other
            Call .SetText(14, .MaxRows, "N")     'Order
            
            Call .SetText(16, .MaxRows, CStr(iRstCnt) & "")      'IFCnt
            
            For i = 1 To iRstCnt
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                
                sCIFSeq = ConvertIFItemInfo(7, sCIFRstCd)
                
                If sCIFSeq = "" Then
                Else
                    'sIFSeq에 따른 정확한 소수자리 처리, 성별에 따른 참고치 처리
                    Call .GetText(10, .MaxRows, vSex)
                    
                    Select Case vSex
                        Case "M", "남", "1"
                            vSex = "M"
                        Case "F", "여", "2"
                            vSex = "F"
                        Case Else
                            vSex = "M"
                    End Select
                    
                    '결과값2(판정)의 특수한 역할을 수행하는 함수
                    Call gfIFDisplayForm.SpecificProcessResult(sCIFRstCd, sCRst1, sCRst2, sCIFSeq, CStr(vSex))
                    
                    Call .GetText(15, .MaxRows, vCRstCnt)      'Result
                    Call .SetText(15, .MaxRows, CStr(Val(vCRstCnt) + 1) & "")
                    
                    Call .SetText(16 + Val(vCRstCnt) + 1, .MaxRows, sCIFSeq & "|" & sCRst1 & "|" & sCRst2 & "|")
                    
                    sNRst1 = sNRst1 & sCRst1 & Chr(124)
                    sNRst2 = sNRst2 & sCRst2 & Chr(124)
                End If
            Next
            
            sRst1 = sNRst1
            sRst2 = sNRst2
            
            Call .GetText(15, .MaxRows, vCRstCnt)
            
            If Val(vCRstCnt) = 0 Then
                .MaxRows = .MaxRows - 1
            End If
            
            '현재 Row 기록
            gResultTable(1).iCRow = .MaxRows
        End With
    End With
End Function

Public Function OldIFList(ByVal iCRow%, ByVal iRstCnt%, _
            ByVal sIFRstCd$, ByVal sRst1$, ByVal sRst2$, ByVal sIFSpcCd$, _
            ByVal sRack$, ByVal sPos$, ByVal sRegNo$, ByVal sName$, _
            ByVal sSex$, ByVal sEmer$, ByVal sReRun$, ByVal sOther$, ByVal sFlag$) As String
    On Error GoTo ErrHandler
    
    Dim i%, j%, k%, iAdd%, iCCol%, iCompCnt%, iAllCnt%, iExist%
    Dim aIFSeq()    As String
    Dim sCIFRstCd$, sCRst1$, sCRst2$, sCFlag$, sCIFSeq$, sTmp$
    Dim sPIFSeq$, sPRst1$, sPRst2$, sPFlag$
    Dim sTIFRstCd$, sTRst1$, sTRst2$, sTFlag$
    Dim vTmp, vIFCnt, vCRstCnt, vRack, vPos, vTTestInfo
    
    OldIFList = "OK"
    
    'iAdd = True : 새로운 항목의 결과 추가, iAdd = False : 기존항목의 결과 재전송
    iAdd = True
    
    With gfIFDisplayForm
        With .spdIntList
            Call .GetText(2, iCRow, vTmp)
                
            For i = 1 To iRstCnt
                Call .GetText(15, iCRow, vCRstCnt)
                Call .GetText(16, iCRow, vIFCnt)
                
                If vCRstCnt = "N" Then
                '처음 결과가 전송되었을 때
                    vCRstCnt = 0
                End If
                
                iCompCnt = Val(vIFCnt)
                
                '전체 sIFRstCd 중 하나씩 가져옴
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                sCFlag = GetByOne(sFlag, sFlag)
                
                'IFSeq로 변환 - 중복되는 항목중 실제 오더경우의 IFSeq를 가져옴
                iAllCnt = 0
                
                'sCIFRstCd와 일치하는 모든 IFSeq 구함
                For j = 1 To giOriginIFItemCnt
                    If gIFItem(j).s04 = sCIFRstCd Then
                        iAllCnt = iAllCnt + 1
                        ReDim Preserve aIFSeq(iAllCnt)
                        aIFSeq(iAllCnt) = gIFItem(j).s01
                    End If
                Next
                
                iExist = 0
                
                For j = 1 To iCompCnt
                '현재 Row의 모든 IFSeq에 대해 실제 IFSeq와 조사
                    Call .GetText(16 + j, iCRow, vTmp)
                    
                    For k = 1 To iAllCnt
                        If vTmp = "" Then
                        Else
                            sTmp = CStr(vTmp)
                            
                            sTmp = GetByOne(sTmp, sTmp)
                            
                            If sTmp = aIFSeq(k) Then
                                iExist = 1
                                sCIFSeq = aIFSeq(k)
                            End If
                        End If
                    Next
                    
                    If iExist = 1 Then
                        Exit For
                    End If
                Next
                
                If iExist = 0 Then
                    sCIFSeq = ConvertIFItemInfo(7, sCIFRstCd)
                End If
                                   
                If sCIFSeq = "" Then
                Else
                    If vCRstCnt = 0 Then
                        '처음 전송
                        If sRack = "" Then
                        Else
                            Call .SetText(6, iCRow, sRack & "")
                        End If
                        
                        If sPos = "" Then
                        Else
                            Call .SetText(7, iCRow, sPos & "")
                        End If
                        
                        If sRegNo = "" Then
                        Else
                            Call .SetText(8, iCRow, sRegNo & "")
                        End If
                        
                        If sName = "" Then
                        Else
                            Call .SetText(9, iCRow, sName & "")
                        End If
                        
                        If sSex = "" Then
                        Else
                            Call .SetText(10, iCRow, sSex & "")
                        End If
                        
                        If sEmer = "" Then
                        Else
                            Call .SetText(11, iCRow, sEmer & "")
                        End If
                        
                        If sReRun = "" Then
                        Else
                            Call .SetText(12, iCRow, sReRun & "")
                        End If
                        
                        If sOther = "" Then
                        Else
                            Call .SetText(13, iCRow, sOther & "")
                        End If
                        
                        Call .SetText(15, iCRow, "1")
                    End If
                    
                    iCCol = 0
                    
                    '기존의 전송받은 검사항목의 IFSeq와 비교함
                    For j = 1 To iCompCnt
                        Call .GetText(16 + j, iCRow, vTmp)
                        sTmp = CStr(vTmp)
                        
                        '이전 검사항목의 IFSeq를 가져옴
                        sPIFSeq = GetByOne(sTmp, sTmp)
                        sPRst1 = GetByOne(sTmp, sTmp)
                        sPRst2 = GetByOne(sTmp, sTmp)
                        sPFlag = GetByOne(sTmp, sTmp)
                        
                        If sPIFSeq = "" Then
                            iCCol = j
                            Exit For
                        Else
                            If sPIFSeq = sCIFSeq Then
                                iAdd = False
                            '현재의 칼럼정보를 넘김
                                iCCol = j
                                Exit For
                            Else
                                iAdd = True
                            End If
                        End If
                    Next
                    
                    '새로이 전송된 항목인지, 재전송된 항목인지에 따라
                    If iAdd = True Then
                        If iCCol = 0 Then
                            Call .SetText(15, iCRow, CVar(Val(vCRstCnt) + 1) & "")
                            Call .SetText(16, iCRow, CVar(Val(vIFCnt) + 1) & "")
                            
                            Call .SetText(16 + Val(vIFCnt) + 1, iCRow, _
                                    sCIFSeq & Chr(124) & sCRst1 & Chr(124) & sCRst2 & Chr(124) & sCFlag & Chr(124) & "")
                        Else
                            Call .SetText(16 + iCCol, iCRow, _
                                    sCIFSeq & Chr(124) & sCRst1 & Chr(124) & sCRst2 & Chr(124) & sCFlag & Chr(124) & "")
                        End If
                    Else
                        If sPRst1 = "" And sPRst2 = "" Then
                            Call .SetText(15, iCRow, CVar(Val(vCRstCnt) + 1) & "")
                        End If
                        
                        Call .SetText(16 + iCCol, iCRow, _
                                sCIFSeq & Chr(124) & sCRst1 & Chr(124) & sCRst2 & Chr(124) & sCFlag & Chr(124) & "")
                    End If
                End If
            Next
        End With
        
        gResultTable(1).iCRow = iCRow
    End With
    
    Exit Function
    
ErrHandler:
    OldIFList = "NO"
    ViewMsg "OldIFList - Err(" & Err.Description & ")"
End Function

Public Function OldIFListBySex(ByVal iCRow%, ByVal iRstCnt%, _
            ByVal sIFRstCd$, sRst1$, sRst2$, ByVal sIFSpcCd$, _
            ByVal sRack$, ByVal sPos$, ByVal sRegNo$, ByVal sName$, _
            ByVal sSex$, ByVal sEmer$, ByVal sReRun$, ByVal sOther$) As String
    
    On Error GoTo ErrHandler
    
    Dim i%, j%, k%, iAdd%, iCCol%, iCompCnt%, iAllCnt%, iExist%
    Dim aIFSeq()    As String
    Dim sCIFRstCd$, sCRst1$, sCRst2$, sCIFSeq$, sTmp$, sPIFSeq$, sPRst1$, sPRst2$, sTIFRstCd$, sTRst1$, sTRst2$
    Dim vTmp, vIFCnt, vCRstCnt, vRack, vPos, vTTestInfo, vSex
    Dim sNRst1$, sNRst2$
    
    OldIFListBySex = "OK"
    sNRst1 = ""
    sNRst2 = ""
    
    'iAdd = True : 새로운 항목의 결과 추가, iAdd = False : 기존항목의 결과 재전송
    iAdd = True
    
    With gfIFDisplayForm
        With .spdIntList
            Call .GetText(2, iCRow, vTmp)
                
            For i = 1 To iRstCnt
                Call .GetText(15, iCRow, vCRstCnt)
                Call .GetText(16, iCRow, vIFCnt)
                
                If vCRstCnt = "N" Then
                '처음 결과가 전송되었을 때
                    vCRstCnt = 0
                End If
                
                iCompCnt = Val(vIFCnt)
                
                '전체 sIFRstCd 중 하나씩 가져옴
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                
                'IFSeq로 변환 - 중복되는 항목중 실제 오더경우의 IFSeq를 가져옴
                iAllCnt = 0
                
                'sCIFRstCd와 일치하는 모든 IFSeq 구함
                For j = 1 To giOriginIFItemCnt
                    If gIFItem(j).s04 = sCIFRstCd Then
                        iAllCnt = iAllCnt + 1
                        ReDim Preserve aIFSeq(iAllCnt)
                        aIFSeq(iAllCnt) = gIFItem(j).s01
                    End If
                Next
                
                iExist = 0
                
                For j = 1 To iCompCnt
                '현재 Row의 모든 IFSeq에 대해 실제 IFSeq와 조사
                    Call .GetText(16 + j, iCRow, vTmp)
                    
                    For k = 1 To iAllCnt
                        If vTmp = "" Then
                        Else
                            sTmp = CStr(vTmp)
                            
                            sTmp = GetByOne(sTmp, sTmp)
                            
                            If sTmp = aIFSeq(k) Then
                                iExist = 1
                                sCIFSeq = aIFSeq(k)
                            End If
                        End If
                    Next
                    
                    If iExist = 1 Then
                        Exit For
                    End If
                Next
                
                If iExist = 0 Then
                    sCIFSeq = ConvertIFItemInfo(7, sCIFRstCd)
                End If
                                   
                If sCIFSeq = "" Then
                Else
                    If vCRstCnt = 0 Then
                        '처음 전송
                        If sRack = "" Then
                        Else
                            Call .SetText(6, iCRow, sRack & "")
                        End If
                        
                        If sPos = "" Then
                        Else
                            Call .SetText(7, iCRow, sPos & "")
                        End If
                        
                        If sRegNo = "" Then
                        Else
                            Call .SetText(8, iCRow, sRegNo & "")
                        End If
                        
                        If sName = "" Then
                        Else
                            Call .SetText(9, iCRow, sName & "")
                        End If
                        
                        If sSex = "" Then
                        Else
                            Call .SetText(10, iCRow, sSex & "")
                        End If
                        
                        If sEmer = "" Then
                        Else
                            Call .SetText(11, iCRow, sEmer & "")
                        End If
                        
                        If sReRun = "" Then
                        Else
                            Call .SetText(12, iCRow, sReRun & "")
                        End If
                        
                        If sOther = "" Then
                        Else
                            Call .SetText(13, iCRow, sOther & "")
                        End If
                        
                        Call .SetText(15, iCRow, "1")
                    End If
                    
                    iCCol = 0
                    
                    '기존의 전송받은 검사항목의 IFSeq와 비교함
                    For j = 1 To iCompCnt
                        Call .GetText(16 + j, iCRow, vTmp)
                        sTmp = CStr(vTmp)
                        
                        '이전 검사항목의 IFSeq를 가져옴
                        sPIFSeq = GetByOne(sTmp, sTmp)
                        sPRst1 = GetByOne(sTmp, sTmp)
                        sPRst2 = GetByOne(sTmp, sTmp)
                        
                        If sPIFSeq = "" Then
                            iCCol = j
                            Exit For
                        Else
                            If sPIFSeq = sCIFSeq Then
                                iAdd = False
                            '현재의 칼럼정보를 넘김
                                iCCol = j
                                Exit For
                            Else
                                iAdd = True
                            End If
                        End If
                    Next
                    
                    'sIFSeq에 따른 정확한 소수자리 처리, 성별에 따른 참고치 처리
                    Call .GetText(10, iCRow, vSex)
                    
                    Select Case vSex
                        Case "M", "남", "1"
                            vSex = "M"
                        Case "F", "여", "2"
                            vSex = "F"
                        Case Else
                            vSex = "M"
                    End Select
                    
                    Call gfIFDisplayForm.SpecificProcessResult(sCIFRstCd, sCRst1, sCRst2, sCIFSeq, CStr(vSex))
                    
                    '새로이 전송된 항목인지, 재전송된 항목인지에 따라
                    If iAdd = True Then
                        If iCCol = 0 Then
                            Call .SetText(15, iCRow, CVar(Val(vCRstCnt) + 1) & "")
                            Call .SetText(16, iCRow, CVar(Val(vIFCnt) + 1) & "")
                            
                            Call .SetText(16 + Val(vIFCnt) + 1, iCRow, sCIFSeq & Chr(124) & sCRst1 & Chr(124) & sCRst2 & Chr(124) & "")
                        Else
                            Call .SetText(16 + iCCol, iCRow, sCIFSeq & Chr(124) & sCRst1 & Chr(124) & sCRst2 & Chr(124) & "")
                        End If
                    Else
                        If sPRst1 = "" And sPRst2 = "" Then
                            Call .SetText(15, iCRow, CVar(Val(vCRstCnt) + 1) & "")
                        End If
                        
                        Call .SetText(16 + iCCol, iCRow, sCIFSeq & Chr(124) & sCRst1 & Chr(124) & sCRst2 & Chr(124) & "")
                    End If
                    
                    sNRst1 = sNRst1 & sCRst1 & Chr(124)
                    sNRst2 = sNRst2 & sCRst2 & Chr(124)
                End If
            Next
        End With
        
        gResultTable(1).iCRow = iCRow
    End With
    
    '결과1, 결과2 바꾸기
    sRst1 = sNRst1
    sRst2 = sNRst2
    
    Exit Function
    
ErrHandler:
    OldIFListBySex = "NO"
    ViewMsg "OldIFListBySex 오류 - (" & Err.Description & ")"
End Function

Public Function ConvertResult1(ByVal sSign As String, ByVal sExp As String, ByVal sRst As String, ByVal sIFRstCd As String, Optional ByVal sIFSeq As String) As String
    Dim i%
    Dim sDot$, sDotGbn$
    Dim sValue$, sTmpVal$
    
    For i = 1 To giOriginIFItemCnt
        If sIFSeq = "" Then
            If sIFRstCd = gIFItem(i).s04 Then
                sDot = gIFItem(i).s07
                sDotGbn = gIFItem(i).s08
                            
                Exit For
            End If
        Else
            If sIFSeq = gIFItem(i).s01 Then
                sDot = gIFItem(i).s07
                sDotGbn = gIFItem(i).s08
                            
                Exit For
            End If
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If sIFRstCd = gCalItem(i).s01 Then
            sDot = gCalItem(i).s05
            sDotGbn = gCalItem(i).s06
                        
            Exit For
        End If
    Next
    
    Select Case sDotGbn
        Case "0"
            sDotGbn = "L"
        Case "1"
            sDotGbn = "H"
        Case "2"
            sDotGbn = "U"
    End Select
    
    If IsNumeric(sRst) = False Then
        ConvertResult1 = sRst
        Exit Function
    End If
    
    If sSign = "" Then
        sSign = "+"
    End If
    
    If sSign = "+" Then
        If sExp = "" Then
            sValue = sRst
        Else
            sValue = CStr(Val(sRst) * (10 ^ Val(sExp)))
        End If
    ElseIf sSign = "-" Then
        If sExp = "" Then
            sValue = sRst
        Else
            sValue = CStr(Val(sRst) / (10 ^ Val(sExp)))
        End If
    End If
    
    If Left(sValue, 1) = "." Then
        sValue = "0" & sValue
    End If
    
'실제 sRst값을 소수점 설정에 따라 바꿈
    If sDot = "" Or IsNumeric(sDot) = False Or sDotGbn = "" Or IsNumeric(sDotGbn) = True Then
        ConvertResult1 = sValue
        Exit Function
    Else
        Select Case sDot
            Case "0"
                sTmpVal = Format$(sValue, "0")
            Case "1"
                sTmpVal = Format$(sValue, "0.0")
            Case "2"
                sTmpVal = Format$(sValue, "0.00")
            Case "3"
                sTmpVal = Format$(sValue, "0.000")
            Case "4"
                sTmpVal = Format$(sValue, "0.0000")
            Case "5"
                sTmpVal = Format$(sValue, "0.00000")
            Case "6"
                sTmpVal = Format$(sValue, "0.000000")
            Case "7"
                sTmpVal = Format$(sValue, "0.0000000")
            Case "8"
                sTmpVal = Format$(sValue, "0.00000000")
            Case "9"
                sTmpVal = Format$(sValue, "0.000000000")
        End Select
        
        Select Case sDotGbn
            '올림
            Case "U"
                Select Case sDot
                    Case "0"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 1)
                        End If
                    Case "1"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.1)
                        End If
                    Case "2"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.01)
                        End If
                    Case "3"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.001)
                        End If
                    Case "4"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.0001)
                        End If
                    Case "5"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.00001)
                        End If
                    Case "6"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.000001)
                        End If
                    Case "7"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.0000001)
                        End If
                    Case "8"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.00000001)
                        End If
                    Case "9"
                        '올림이 아니라 반대로 내림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.000000001)
                        End If
                End Select
                
            '반올림
            Case "H"
                
            '내림
            Case "L"
                Select Case sDot
                    Case "0"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 1)
                        End If
                    Case "1"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.1)
                        End If
                    Case "2"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.01)
                        End If
                    Case "3"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.001)
                        End If
                    Case "4"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.0001)
                        End If
                    Case "5"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.00001)
                        End If
                    Case "6"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.000001)
                        End If
                    Case "7"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.0000001)
                        End If
                    Case "8"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.00000001)
                        End If
                    Case "9"
                        '내림이 아니라 반대로 올림이 된 경우
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.000000001)
                        End If
                End Select
        End Select
        
        ConvertResult1 = sTmpVal
    End If
End Function

Public Sub CurRstDisplay(ByVal iRow As Integer, ByVal sTestNm As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                            ByVal sFcolor As String, ByVal sBcolor As String)
    Dim i%
    Dim vTestNm
    
    With gfIFDisplayForm.spdIntList
        For i = 16 To .MaxCols
            Call .GetText(i, 0, vTestNm)
            
            If CStr(vTestNm) = sTestNm Then
                Call SpdForeBack(gfIFDisplayForm.spdIntList, i, i, iRow, iRow, sFcolor, sBcolor)
                Call .SetText(i, iRow, sRst1 & " " & sRst2 & "")
            End If
        Next
    End With
End Sub

Public Sub DisplayInit()
    Dim i%
        
    'Title 결정
    gfIFDisplayForm.Caption = "   " & UCase$(gsMachineNm) & " 인터페이스 화면 - BY ACK Co., Ltd."
    
    With gfIFDisplayForm.spdIntList
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        .Lock = True
        .NoBeep = True
        .BlockMode = False
        
        .Col = 6
        .Col2 = 7
        .Row = -1
        .Row2 = -1
        .BlockMode = True
        .Lock = False
        .BlockMode = False
            
        Call SetSpdIntLIstColHidden
         
        'Rack, Pos 사용여부
        If Val(gIFRack.sMaxRack) = 0 Then
            For i = 6 To 7
                .Col = i
                .ColHidden = True
            Next
        Else
            For i = 6 To 7
                .Col = i
                .ColHidden = False
            Next
        End If
        
        .MaxRows = 0
    End With
        
    With gfIFDisplayForm.spdRst
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        .EditModePermanent = True
        .NoBeep = True
        .BlockMode = False
                
        .BlockMode = True
        .Col = 1
        .Col2 = 4
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
    End With
    
     With gfIFDisplayForm.spdRst2
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        .EditModePermanent = True
        .NoBeep = True
        .BlockMode = False
                
        .BlockMode = True
        .Col = 1
        .Col2 = 4
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
    End With

'Interface Mode에 따른 Display
    If gsIFMode = "0" Then
    'Uni-Direction
        With gfIFDisplayForm
            .fraSendOrd.Visible = False
            .fraBarCd.Top = 7500
        End With
    Else
    'Bi-Direction
        '1=양방향(Rack Or Tray 방식 지원안함, But Rack/Pos 표시)
        '2=양방향(Rack Or Tray 방식 지원안함, But Tray/Pos 표시)
        '3=양방향(Rack Or Tray 방식 지원안함, But Tray/Cup 표시)
        '4=양방향(Rack/Pos 방식 지원),
        '5=양방향(Tray/Pos 방식 지원),
        '6=양방향(Tray/Cup 방식 지원),

        With gfIFDisplayForm
            .fraBarCd.Visible = False
            
            If gsIFMode = "1" Then
            'Rack Or Tray 방식 지원안함, But Rack/Pos 표시
                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Rack"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "2" Then
            'Rack Or Tray 방식 지원안함, But Tray/Pos 표시
                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "3" Then
            'Rack Or Tray 방식 지원안함, But Tray/Cup 표시
                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Cup"))
            ElseIf gsIFMode = "4" Then
            'Rack/Pos 방식 지원
                .pnlRackTray = "Rack"
                .pnlPosCup = "Pos"
                
                Call .spdIntList.SetText(6, 0, CVar("Rack"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "5" Then
            'Tray/Pos 방식 지원
                .pnlRackTray = "Tray"
                .pnlPosCup = "Pos"
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "6" Then
            'Tray/Cup 방식 지원
                .pnlRackTray = "Tray"
                .pnlPosCup = "Cup"
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Cup"))
            End If
        End With
    End If
    
'Transmit Mode에 따른 Display
    If gsTXMode = "0" Then
    'Batch
        '등록 Option을 Client로 하면 OK
    ElseIf gsTXMode = "1" Then
    'RealTime 한 항목씩
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    ElseIf gsTXMode = "2" Then
    'RealTime  한 환자씩
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    End If
    
'Initialize mode에 따른 Display
    If gsINITMode = "0" Then
    'Not Use
        gfIFDisplayForm.cmdInitial.Visible = False
    Else
    'Use
        gfIFDisplayForm.cmdInitial.Visible = True
    End If
    
'MaxLength Check
    With gfIFDisplayForm
        .txtRack.MaxLength = CInt(Val(gIFRack.sRackDigit))
        .txtPos.MaxLength = CInt(Val(gIFRack.sPosDigit))
        .txtOrdNo.MaxLength = CInt(Val(gOrdCfg.sFSize(3)))
        .txtBarCd.MaxLength = CInt(Val(gOrdCfg.sFSize(3)))
    End With
    
'APMode에 따른 결과 Display
    If gsAPMode = "1" Then
        With gfIFDisplayForm.spdRst
            .ColWidth(1) = 10#
            .ColWidth(2) = 7.5
            .ColWidth(3) = 0#
            .ColWidth(4) = 4#
        End With
        
        With gfIFDisplayForm.spdRst2
            .ColWidth(1) = 10#
            .ColWidth(2) = 7.5
            .ColWidth(3) = 0#
            .ColWidth(4) = 4#
        End With
    End If
End Sub

Public Sub DisplayNextRackPos(ByVal sRackFlag$)
    On Error GoTo ErrHandler
    
    With gfIFDisplayForm
        'Alphabet
        If gIFRack.sRackDigit = 1 And Val(gIFRack.sMaxRack) = 26 Then
            If Val(.txtPos) = Val(sRackFlag) Then
                'Pos 초기화
                If .txtPos.MaxLength >= 1 And .txtPos.MaxLength <= 10 Then
                    .txtPos = Format("1", RackFormat(.txtPos.MaxLength))
                Else
                    .txtPos = "0"
                End If
                
                'Rack 증가
                Select Case .txtRack.MaxLength
                    Case 1
                        If .txtRack = "" Then
                            .txtRack = "A"
                        Else
                            .txtRack = Chr(Asc(.txtRack) + 1)
                        End If
                End Select
                
                Exit Sub
            End If
        '1, 01, 001, 0001, ....
        Else
            If Val(.txtPos) = Val(sRackFlag) Then
                'Pos 초기화
                If .txtPos.MaxLength >= 1 And .txtPos.MaxLength <= 10 Then
                    .txtPos = Format("1", RackFormat(.txtPos.MaxLength))
                Else
                    .txtPos = "0"
                End If
                
                'Rack 증가
                If .txtRack.MaxLength >= 1 And .txtRack.MaxLength <= 10 Then
                    .txtRack = Format(Val(.txtRack) + 1, RackFormat(.txtRack.MaxLength))
                Else
                    .txtRack = "0"
                End If
                
                Exit Sub
            End If
        End If
        
        'Pos 증가
        If .txtPos.MaxLength >= 1 And .txtPos.MaxLength <= 10 Then
            .txtPos = Format(Val(.txtPos) + 1, RackFormat(.txtPos.MaxLength))
        Else
            .txtPos = "0"
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayNextRackPos 오류 - (" & Err.Description & ")"
End Sub

Public Sub DisplayResult2(ByVal iRow As Integer)
    On Error GoTo ErrHandler
    
    Dim vIFCnt, vTmp, vJDate, vJGbn, vJNo
    Dim i%, j%, k%
    Dim sTmp$, sTestCd$, sCRst1$, sTestNm$, sCRst2$, sCFlag$
    
    Dim tmpData()   As String
    
    
    With gfIFDisplayForm
        .lblCSelList = ""
    End With
    
    Call ResultSpdClear2
    
    If iRow <= 0 Then
        Exit Sub
    End If
    
    With gfIFDisplayForm.spdIntList
        Call .GetText(3, iRow, vJDate)
        Call .GetText(4, iRow, vJGbn)
        Call .GetText(5, iRow, vJNo)
        Call .GetText(16, iRow, vIFCnt)
        
        If Len(vJDate) = 0 Then
            vJDate = ""
        Else
            vJDate = vJDate & "-"
        End If
        
        If Len(vJGbn) = 0 Then
            vJGbn = ""
        Else
            vJGbn = vJGbn & "-"
        End If
                    
        gfIFDisplayForm.lblCSelList = "결과조회 : " & CStr(vJDate) & CStr(vJGbn) & CStr(vJNo)
        
        For i = 1 To CInt(vIFCnt)
            sTestCd = "": sCRst1 = "": sCRst2 = "": sTestNm = ""
            Call .GetText(16 + i, iRow, vTmp)
            
            sTmp = CStr(vTmp)
            
            If Trim(sTmp) <> "" Then
                tmpData() = Split(sTmp, Chr(124))
                
                sTestCd = tmpData(0)    'GetByOne(sTmp, sTmp)
                sCRst1 = tmpData(1)     'GetByOne(sTmp, sTmp)
                sCRst2 = tmpData(2)     'GetByOne(sTmp, sTmp)
                sCFlag = tmpData(3)
            Else
                sTestCd = ""
                sCRst1 = ""
                sCRst2 = ""
                sCFlag = ""
            End If
            
            sTestNm = ConvertIFItemInfo(4, sTestCd)
                            
            If i > 15 Then
                Call gfIFDisplayForm.spdRst2.SetText(1, i - 15, sTestNm & "")
                Call gfIFDisplayForm.spdRst2.SetText(2, i - 15, sCRst1 & "")
                Call gfIFDisplayForm.spdRst2.SetText(3, i - 15, sCRst2 & "")
                Call gfIFDisplayForm.spdRst2.SetText(4, i - 15, sCRst2 & "")
                Call gfIFDisplayForm.spdRst2.SetText(5, i - 15, sCFlag & "")
                
                If sCRst2 = "High" Or sCRst2 = "Positive" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst2, 1, 4, i - 15, i - 15, RGB(0, 0, 0), RGB(255, 220, 220))
                ElseIf sCRst2 = "Low" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst2, 1, 4, i - 15, i - 15, RGB(0, 0, 0), RGB(220, 220, 255))
                ElseIf sCRst2 = "H" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst2, 1, 4, i - 15, i - 15, RGB(0, 0, 0), RGB(255, 220, 220))
                ElseIf sCRst2 = "L" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst2, 1, 4, i - 15, i - 15, RGB(0, 0, 0), RGB(220, 220, 255))
                ElseIf sCRst2 <> "" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst2, 1, 4, i - 15, i - 15, RGB(0, 0, 0), RGB(230, 230, 230))
                End If
            Else
                Call gfIFDisplayForm.spdRst.SetText(1, i, sTestNm & "")
                Call gfIFDisplayForm.spdRst.SetText(2, i, sCRst1 & "")
                Call gfIFDisplayForm.spdRst.SetText(3, i, sCRst2 & "")
                Call gfIFDisplayForm.spdRst.SetText(4, i, sCRst2 & "")
                Call gfIFDisplayForm.spdRst.SetText(5, i, sCFlag & "")
                
                If sCRst2 = "High" Or sCRst2 = "Positive" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(255, 220, 220))
                ElseIf sCRst2 = "Low" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(220, 220, 255))
                ElseIf sCRst2 = "H" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(255, 220, 220))
                ElseIf sCRst2 = "L" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(220, 220, 255))
                ElseIf sCRst2 <> "" Then
                    Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(230, 230, 230))
                End If
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayResult2 에러발생" & "(" & CStr(Err.Description) & ")"
End Sub

Public Sub EditRegState(ByVal iPersonCnt As Integer, ByVal sWDate As String, ByVal sTWSeq As String)
    On Error GoTo ErrHandler
    
    Dim objld As Object
    
    Set objld = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    Call objld.Edit_IFResult(gsMachineCd, 2, sWDate, sTWSeq, "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", iPersonCnt)
    
    Set objld = Nothing
      
    Exit Sub
      
ErrHandler:
    Set objld = Nothing
    ViewMsg "EditRegState 오류 - (" & Err.Description & ")"
End Sub

Public Function FindIFListRow(ByVal iMode As Integer, sJDate As String, sJGbn As String, sJNo As String, sWSeq As String, Optional ByVal sCurRow As String) As Integer
    Dim i%
    Dim vJDate, vJGbn, vJNo, vWSeq, vTmp
    
    FindIFListRow = 0
    
    If Trim(sJDate) = "" And Trim(sJGbn) = "" And Trim(sJNo) = "" Then
        Exit Function
    End If
    
    With gfIFDisplayForm.spdIntList
        Select Case iMode
            'sJDate, sJGbn, sJNo Match
            Case 0
                For i = 1 To .MaxRows
                    Call .GetText(3, i, vJDate)
                    Call .GetText(4, i, vJGbn)
                    Call .GetText(5, i, vJNo)
                    
                    If CStr(vJDate) = sJDate And CStr(vJGbn) = sJGbn And CStr(vJNo) = sJNo Then
                        sWSeq = ""
                        FindIFListRow = i
                    End If
                Next
                
            'sWSeq Match
            Case 1
                For i = 1 To .MaxRows
                    Call .GetText(1, i, vWSeq)
                    
                    If CStr(vWSeq) = sWSeq Then
                        sJDate = ""
                        sJGbn = ""
                        sJNo = ""
                        FindIFListRow = i
                    End If
                Next
                
            'sCurRow Match
            Case 2
                Call .GetText(1, CInt(sCurRow), vWSeq)
                Call .GetText(3, CInt(sCurRow), vJDate)
                Call .GetText(4, CInt(sCurRow), vJGbn)
                Call .GetText(5, CInt(sCurRow), vJNo)
                
                sWSeq = vWSeq
                sJDate = ""
                sJGbn = ""
                sJNo = ""
                
                FindIFListRow = CInt(sCurRow)
            
            'sJNo Match
            Case 3
                For i = 1 To .MaxRows
                    Call .GetText(5, i, vJNo)
                    
                    If CStr(vJNo) = sJNo Then
                        sWSeq = ""
                        FindIFListRow = i
                    End If
                Next
                
            'Worklist와 순서로 Match
            Case 4
                If .MaxRows = 0 Then
                    FindIFListRow = 0
                Else
                    For i = 1 To .MaxRows
                        Call .GetText(15, i, vTmp)
                        
                        If vTmp = "N" Then
                            sWSeq = ""
                            FindIFListRow = i
                            Exit For
                        Else
                            FindIFListRow = 0
                        End If
                    Next
                End If
                
            Case Else
            
        End Select
        
    End With
End Function

Public Function FindIFListWithJ(ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String) As Integer
    Dim i%
    Dim vJDate, vJGbn, vJNo
    
    FindIFListWithJ = 0
    
    If Trim(sJDate) = "" And Trim(sJGbn) = "" And Trim(sJNo) = "" Then
        Exit Function
    End If
    
    With gfIFDisplayForm.spdIntList
        For i = 1 To .MaxRows
            Call .GetText(3, i, vJDate)
            Call .GetText(4, i, vJGbn)
            Call .GetText(5, i, vJNo)
            
            If CStr(vJDate) = sJDate And CStr(vJGbn) = sJGbn And CStr(vJNo) = sJNo Then
                FindIFListWithJ = i
            End If
        Next
    End With
End Function

Public Function FindIFListWithJNo(ByVal sJNo As String) As Integer
    Dim i%
    Dim vJNo
    
    FindIFListWithJNo = 0
    
    If Trim(sJNo) = "" Then
        Exit Function
    End If
    
    With gfIFDisplayForm.spdIntList
        For i = 1 To .MaxRows
            Call .GetText(5, i, vJNo)
            
            If CStr(vJNo) = sJNo Then
                FindIFListWithJNo = i
            End If
        Next
    End With
End Function

Public Function FindIFListWithW(ByVal sWSeq As String) As Integer
    Dim i%
    Dim vWSeq
    
    FindIFListWithW = 0
    
    If Trim(sWSeq) = "" Then
        Exit Function
    End If
    
    With gfIFDisplayForm.spdIntList
        For i = 1 To .MaxRows
            Call .GetText(1, i, vWSeq)
            
            If CStr(vWSeq) = sWSeq Then
                FindIFListWithW = i
            End If
        Next
    End With
End Function

Public Function GetByOne(ByVal tStr As String, sOriginal As String) As String
    Dim Pos%
    
    Pos = InStr(tStr, Chr$(124))
    
    If Pos = 0 Then
    Else
        GetByOne = Trim$(Mid$(tStr, 1, Pos - 1))
        sOriginal = Trim$(Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos))
    End If
End Function

Public Function GetByOneRow(ByVal tStr As String, sOriginal As String) As String
    Dim Pos%
    
    Pos = InStr(tStr, Chr$(13))
    
    If Pos = 0 Then
    Else
        GetByOneRow = Trim$(Mid$(tStr, 1, Pos - 1))
        sOriginal = Trim$(Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos))
    End If
End Function

Public Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim Pos%

    Pos = InStr(tStr, sUserSymbol)

    If Pos = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, Pos - 1))
        sOriginal = Trim$(Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos))
    End If
End Function

Public Sub GetInterfaceCd()
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Interface.Mode")
        
    gsIFMode = sBuf
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Transmit.Mode")
        
    gsTXMode = sBuf
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Auto.P.Mode")
        
    gsAPMode = sBuf
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Initialize.Mode")
        
    gsINITMode = sBuf
End Sub

Public Function GetItemColWidth()
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Item.ColWidth")
        
    GetItemColWidth = Val(sBuf)
End Function

Public Function GetLastWorkSeq(ByVal sWDate As String) As String
    Dim objDB As Object
    Dim sRtnVal$
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    sRtnVal = objDB.Get_LastIFResult(gsMachineCd, sWDate)
    
    gsLastWSeq = Format(GetByOneUserSymbol(sRtnVal, sRtnVal, Chr(3)), "0000")
    
    Set objDB = Nothing
End Function

Public Function GetCurLastWSeq() As String
    GetCurLastWSeq = ""
    
    With gfIFDisplayForm.spdIntList
        Call GetLastWorkSeq(Format(frmInterface.dtpLabDate.Value, "YYYYMMDD"))
        GetCurLastWSeq = gsLastWSeq
            
        Exit Function
    End With
End Function

Public Sub GetMachineInfo()
    Dim retval As Long
    Dim sBuf As String
    Dim bRetVal As Boolean
    Dim i%

'Comm Info
'Port Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "ComPort")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "ComPort", "1")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gCommInfo.sPort = "1"
    Else
        gCommInfo.sPort = sBuf
    End If

'BaudRate Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "BaudRate")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "BaudRate", "9600")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gCommInfo.sBaudRate = "9600"
    Else
        gCommInfo.sBaudRate = sBuf
    End If
    
'DataBit Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "DataBit")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "DataBit", "8")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gCommInfo.sDataBit = "8"
    Else
        gCommInfo.sDataBit = sBuf
    End If

'StopBit Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "StopBit")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "StopBit", "1")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gCommInfo.sStopBIt = "1"
    Else
        gCommInfo.sStopBIt = sBuf
    End If

'Parity Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity", "N")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gCommInfo.sParity = "N"
    Else
        gCommInfo.sParity = sBuf
    End If
    
'Rack Info
'RackDigit
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "RackDig")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "RackDig", "3")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gIFRack.sRackDigit = "3"
    Else
        gIFRack.sRackDigit = sBuf
    End If

'PosDigit
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "PosDig")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "PosDig", "2")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gIFRack.sPosDigit = "2"
    Else
        gIFRack.sPosDigit = sBuf
    End If

'MaxRack
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "MaxRack")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "MaxRack", "20")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gIFRack.sMaxRack = "20"
    Else
        gIFRack.sMaxRack = sBuf
    End If

'PosSetting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "PosSetting")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "PosSetting", "||||||||||||||||||||")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        gIFRack.sPosSetting = "||||||||||||||||||||"
    Else
        gIFRack.sPosSetting = sBuf
    End If
    
    Erase gIFPosInfo
    
    ReDim gIFPosInfo(CInt(gIFRack.sMaxRack))
    
    sBuf = gIFRack.sPosSetting
    
    For i = 1 To CInt(gIFRack.sMaxRack)
        'ALPHABET
        If Val(gIFRack.sRackDigit) = 1 And Val(gIFRack.sMaxRack) = 26 Then
            gIFPosInfo(i).sRackNo = Chr(Asc("A") + i - 1)
            gIFPosInfo(i).sPosMaxNo = GetByOne(sBuf, sBuf)
        'NUMERIC
        Else
            gIFPosInfo(i).sRackNo = Format$(i, RackFormat(gIFRack.sRackDigit))
            gIFPosInfo(i).sPosMaxNo = GetByOne(sBuf, sBuf)
        End If
    Next
    
'Path & Exe Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "App.Exe")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "App.Exe", App.Path & "\" & gsMachineNm & ".exe")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
    Else
    End If

'Interface Mode
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Interface.Mode")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Interface.Mode", "0")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
        gsIFMode = "0"   'Default 단방향
    Else
        gsIFMode = sBuf
    End If
    
'Initialize Mode
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Initialize.Mode")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Initialize.Mode", "0")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
        gsINITMode = "0"   'Default 사용안함
    Else
        gsINITMode = sBuf
    End If
    
'Transmit Mode
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Transmit.Mode")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Transmit.Mode", "0")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
        gsTXMode = "0"   'Default Batch
    Else
        gsTXMode = sBuf
    End If

'AutoPrint Mode
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Auto.P.Mode")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Auto.P.Mode", "0")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
        gsAPMode = "0"   'Default No
    Else
        gsAPMode = sBuf
    End If

'gsIFVar1
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "if.var1")
    
    gsIFVar1 = sBuf
    
'gsIFVar2
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "if.var2")
    
    gsIFVar2 = sBuf
    
'gsIFVar3
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "if.var3")
    
    gsIFVar3 = sBuf
    
'gsIFVar4
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "if.var4")
    
    gsIFVar4 = sBuf
    
'gsIFVar5
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "if.var5")
    
    gsIFVar5 = sBuf
End Sub

Public Sub GetOrdRstCfg()
    Dim sBuf$
    Dim i%
    
'Order.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Use")
        
    gOrdCfg.sUse = sBuf
        
'Order.Component
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Component")
        
    gOrdCfg.sComponent = sBuf
        
'Order.Storage.Type
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Storage.Type")
        
    gOrdCfg.sStorageType = sBuf
    
'Order.Storage.Path
    If gOrdCfg.sStorageType = "" Then
        gOrdCfg.sPath = ""
    ElseIf gOrdCfg.sStorageType = "File" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.FILE.Path")
            
        gOrdCfg.sPath = sBuf
    ElseIf gOrdCfg.sStorageType = "Database" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.DB.DSN")
            
        gOrdCfg.sPath = sBuf
    Else
    End If
    
'Result.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Use")
        
    gRstcfg.sUse = sBuf
    
'Result.Component
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Component")
        
    gRstcfg.sComponent = sBuf
    
'Result.Storage.Type
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Storage.Type")
        
    gRstcfg.sStorageType = sBuf
    
'Result.Storage.Path
    If gRstcfg.sStorageType = "" Then
        gRstcfg.sPath = ""
    ElseIf gRstcfg.sStorageType = "File" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.FILE.Path")
            
        gRstcfg.sPath = sBuf
    ElseIf gRstcfg.sStorageType = "Database" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.DB.DSN")
            
        gRstcfg.sPath = sBuf
    Else
    End If

'Order.Field.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Use")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFUse(i) = GetByOne(sBuf, sBuf)
    Next
    
'Order.Field.FName
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FName")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFName(i) = GetByOne(sBuf, sBuf)
    Next
    
'Order.Field.FOrder
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FOrder")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFOrd(i) = Val(GetByOne(sBuf, sBuf))
    Next

'Order.Field.Size
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Size")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFSize(i) = Val(GetByOne(sBuf, sBuf))
    Next

'Result.Field.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Use")
    
    For i = 1 To MAXRESULTFIELD
        gRstcfg.sFUse(i) = GetByOne(sBuf, sBuf)
    Next

'Result.Field.FName
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FName")
    
    For i = 1 To MAXRESULTFIELD
        gRstcfg.sFName(i) = GetByOne(sBuf, sBuf)
    Next
    
'Result.Field.FOrder
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FOrder")
    
    For i = 1 To MAXRESULTFIELD
        gRstcfg.sFOrd(i) = Val(GetByOne(sBuf, sBuf))
    Next

'Result.Field.Size
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Size")
    
    For i = 1 To MAXRESULTFIELD
        gRstcfg.sFSize(i) = Val(GetByOne(sBuf, sBuf))
    Next
End Sub

Public Sub GetTestCdSeq()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.TestCd.Seq")
    
    gsOrdTestCdSeq = sBuf
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.TestCd.Seq")
    
    gsRstTestCdSeq = sBuf
    
    Exit Sub
    
ErrHandler:
    ViewMsg "GetTestCdSeq - Err(" & Err.Description & ")"
End Sub

Public Sub GetTestItem()
    On Error GoTo ErrHandler
    
    Dim objDB As Object
    Dim sRetVal1$, sRetVal2$
    Dim iItemCnt%
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    sRetVal1 = objDB.Get_IFTestItem(gsMachineCd, 0)
    
    sRetVal2 = objDB.Get_CalTestItem(gsMachineCd, 0)
        
    If sRetVal1 = "NONE" Then
    Else
        iItemCnt = GetByOneUserSymbol(sRetVal1, sRetVal1, Chr$(3))
        giOriginIFItemCnt = iItemCnt
        Call MakeIFItemStruct(sRetVal1, iItemCnt)
    End If
    
    If sRetVal2 = "NONE" Then
    Else
        iItemCnt = GetByOneUserSymbol(sRetVal2, sRetVal2, Chr$(3))
        giOriginCalItemCnt = iItemCnt
        Call MakeCalItemStruct(sRetVal2, iItemCnt)
    End If
    
    giTotIFItemCnt = giOriginIFItemCnt + giOriginCalItemCnt
    
    Exit Sub
    
ErrHandler:
    Set objDB = Nothing
    ViewMsg "GetIFTestItem - Local DB 연결 실패!!"
End Sub

Public Sub GetTestMode()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Test.Mode")
    
    giTestMode = Val(sBuf)
    
    Exit Sub
    
ErrHandler:
    ViewMsg "GetTestMode - Err(" & Err.Description & ")"
End Sub

Public Sub GetCSMode()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "C/S.Mode")
    
    gsClientServerMode = sBuf
    
    Exit Sub
    
ErrHandler:
    ViewMsg "GetCSMode - Err(" & Err.Description & ")"
End Sub

Public Function GetExcelExePath()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.Exe.Path")
    
    GetExcelExePath = sBuf
    
    Exit Function
    
ErrHandler:
    ViewMsg "GetExcelExePath - Err(" & Err.Description & ")"
End Function

Public Function GetExcelFilePath()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.File.Path")
    
    GetExcelFilePath = sBuf
    
    Exit Function
    
ErrHandler:
    ViewMsg "GetExcelFilePath - Err(" & Err.Description & ")"
End Function

Public Function ifFileExists(ByVal strfilename As String) As Integer
    Dim i As Integer
    On Error Resume Next
    
    i = Len(Dir$(strfilename))
    
    If Err Or i = 0 Then
        ifFileExists = False
    Else
        ifFileExists = True
    End If
End Function

Public Function IsOp(ByVal Op As String) As Boolean
    IsOp = False
    
    If (Op Like "[+-]") Or _
        (Op = "\") Or _
        (Op = "%") Or _
        (Op = "*") Or _
        (Op = "/") Or _
        (Op = "^") Then
        IsOp = True
    End If
End Function

Public Sub IFMachineCd()
    Dim sBuf$
    Dim retval As Long
    
'Machine Code
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineCd", "", sBuf, 255, App.Path & "\장비코드.ini")
    
    If retval = 0 Then
        MsgBox "장비코드 설정이 되어 있지 않습니다. 프로그램을 제대로 실행할 수 없습니다!!", vbCritical, "장비코드.ini 설정"
    End If
    
    gsMachineCd = Left(sBuf, retval) 'Machine Name
    
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineNm", "", sBuf, 255, App.Path & "\장비코드.ini")
    
    If retval = 0 Then
        MsgBox "장비코드 설정이 되어 있지 않습니다. 프로그램을 제대로 실행할 수 없습니다!!", vbCritical, "장비코드.ini 설정"
    End If
    
    gsMachineNm = Left(sBuf, retval)
    
'Machine Exe
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineExe", "", sBuf, 255, App.Path & "\장비코드.ini")
    
    If retval = 0 Then
        MsgBox "장비코드 설정이 되어 있지 않습니다. 프로그램을 제대로 실행할 수 없습니다!!", vbCritical, "장비코드.ini 설정"
    End If
    
    gsMachineExe = Left(sBuf, retval)
End Sub

Public Function JudgeResultBySex$(ByVal sIFSeq$, ByVal sCompRst$, ByVal sSex$, ByVal sPrevRst$, ByVal sDateDiff$, sRefFlag$, sPanFlag$, sDelFlag$)
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sJGbn$, sRef1$, sRef2$, sPanL$, sPanH$, sDelGbn$, sDelL$, sDelH$
    
    If Len(sIFSeq) = 3 Then
        For i = 1 To giOriginIFItemCnt
            If sIFSeq = gIFItem(i).s01 Then
                sJGbn = gIFItem(i).s09
                
                If sSex = "M" Then
                    sRef1 = gIFItem(i).s10
                    sRef2 = gIFItem(i).s11
                ElseIf sSex = "F" Then
                    sRef1 = gIFItem(i).s12
                    sRef2 = gIFItem(i).s13
                End If
                
                sPanL = gIFItem(i).s14
                sPanH = gIFItem(i).s15
                
                sDelGbn = gIFItem(i).s16
                sDelL = gIFItem(i).s17
                sDelH = gIFItem(i).s18
                
                Exit For
            End If
        Next
    ElseIf Len(sIFSeq) = 2 Then
        For i = 1 To giOriginCalItemCnt
            If sIFSeq = gCalItem(i).s01 Then
                sJGbn = gCalItem(i).s07
                
                If sSex = "M" Then
                    sRef1 = gCalItem(i).s08
                    sRef2 = gCalItem(i).s09
                ElseIf sSex = "F" Then
                    sRef1 = gCalItem(i).s10
                    sRef2 = gCalItem(i).s11
                End If
                
                sPanL = gCalItem(i).s12
                sPanH = gCalItem(i).s13
                
                sDelGbn = gCalItem(i).s14
                sDelL = gCalItem(i).s15
                sDelH = gCalItem(i).s16
                
                Exit For
            End If
        Next
    End If
    
    sRefFlag = "": sPanFlag = "": sDelFlag = ""
    
    If Trim(sCompRst) = "" Then Exit Function
    
    If IsNumeric(sCompRst) = False Then
        If sJGbn <> "0" Then
            sRefFlag = "H"
        End If
        
        If sPanL <> "" Or sPanH <> "" Then
            sPanFlag = "P"
        End If
        
        JudgeResultBySex = sCompRst
    
        Exit Function
    End If
        
    Select Case sJGbn
        Case "0"
            JudgeResultBySex = sCompRst
            sRefFlag = ""
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultBySex = sCompRst
                sRefFlag = "L"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                JudgeResultBySex = sCompRst
                sRefFlag = ""
            Else
                JudgeResultBySex = sCompRst
                sRefFlag = "H"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultBySex = "NEGATIVE"
                sRefFlag = ""
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeResultBySex = "POSITIVE"
                sRefFlag = ""
            Else
                JudgeResultBySex = "TRACE"
                sRefFlag = ""
            End If
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultBySex = sCompRst
                sRefFlag = "N"
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeResultBySex = sCompRst
                sRefFlag = "P"
            Else
                JudgeResultBySex = sCompRst
                sRefFlag = "T"
            End If
        Case "4"
            
        Case "5"
        'QAL P/N
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultBySex = "POSITIVE"
                sRefFlag = ""
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeResultBySex = "NEGATIVE"
                sRefFlag = ""
            Else
                JudgeResultBySex = "TRACE"
                sRefFlag = ""
            End If
        Case "6"
        'QAN P/N
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultBySex = sCompRst
                sRefFlag = "P"
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeResultBySex = sCompRst
                sRefFlag = "N"
            Else
                JudgeResultBySex = sCompRst
                sRefFlag = "T"
            End If
        
        Case Else
        
    End Select
    
    
    'PANIC
    If Val(sCompRst) < Val(sPanL) And Trim(sPanL) <> "" Then
        sPanFlag = "P"
    End If
    
    If Val(sCompRst) > Val(sPanH) And Trim(sPanH) <> "" Then
        sPanFlag = "P"
    End If
    
    'DELTA
    If Trim(sPrevRst) <> "" And IsNumeric(sPrevRst) = True Then
        If Val(sDateDiff) = 0 Then sDateDiff = "1"
        
        Select Case sDelGbn
            '사용안함
            Case "0"
                sDelFlag = ""
            '변화차
            Case "1"
                If Val(sCompRst) >= Val(sPrevRst) Then
                    If Val(sCompRst) - Val(sPrevRst) > sDelH Then
                        sDelFlag = "D"
                    End If
                Else
                    If Abs(Val(sCompRst) - Val(sPrevRst)) > Abs(sDelL) Then
                        sDelFlag = "D"
                    End If
                End If
            '변화비율
            Case "2"
                If Val(sCompRst) >= Val(sPrevRst) Then
                    If Abs((Val(sCompRst) - Val(sPrevRst)) / sPrevRst * 100) > sDelH Then
                        sDelFlag = "D"
                    End If
                Else
                    If Abs((Val(sCompRst) - Val(sPrevRst)) / sPrevRst * 100) > Abs(sDelL) Then
                        sDelFlag = "D"
                    End If
                End If
            '기간당 변화차
            Case "3"
                If Val(sCompRst) >= Val(sPrevRst) Then
                    If (Val(sCompRst) - Val(sPrevRst)) / Val(sDateDiff) > sDelH Then
                        sDelFlag = "D"
                    End If
                Else
                    If Abs(Val(sCompRst) - Val(sPrevRst)) / Val(sDateDiff) > Abs(sDelL) Then
                        sDelFlag = "D"
                    End If
                End If
            '기간당 변화비율
            Case "4"
                If Val(sCompRst) >= Val(sPrevRst) Then
                    If Abs((Val(sCompRst) - Val(sPrevRst)) / sPrevRst * 100 / Val(sDateDiff)) > sDelH Then
                        sDelFlag = "D"
                    End If
                Else
                    If Abs((Val(sCompRst) - Val(sPrevRst)) / sPrevRst * 100 / Val(sDateDiff)) > Abs(sDelL) Then
                        sDelFlag = "D"
                    End If
                End If
            '절대변화비율
            Case "5"
                If Abs(Val(sCompRst) - Val(sPrevRst)) / sPrevRst > sDelH Then
                    sDelFlag = "D"
                End If
                
                If Abs(Val(sCompRst) - Val(sPrevRst)) / sPrevRst > Abs(sDelL) Then
                    sDelFlag = "D"
                End If
        End Select
    End If
    
    Exit Function
    
ErrHandler:
    ViewMsg "JudgeResultBySex - Err(" & Err.Description & ")"
End Function

Public Function JudgeResultWithSex(ByVal sIFRstCd As String, ByVal sCompRst As String, sOneRst2 As String, ByVal sSex$, Optional ByVal sIFSeq$) As String
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sBuf$, sBuf1$, sBuf2$, sJGbn$, sRef1$, sRef2$, sLimit1Gbn$, sLimit2Gbn$, sLimit1$, sLimit2$
    
    For i = 1 To giOriginIFItemCnt
        If sIFSeq = "" Then
            If sIFRstCd = gIFItem(i).s04 Then
                sJGbn = gIFItem(i).s09
                    
                Select Case sSex
                    Case "M"
                        sBuf = gIFItem(i).s10 & ","
                        sBuf1 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sRef1 = sBuf1
                        
                        sBuf = gIFItem(i).s11 & ","
                        sBuf1 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sRef2 = sBuf1
                    Case "F"
                        sBuf = gIFItem(i).s10 & ","
                        sBuf1 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sBuf2 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        
                        If sBuf2 = "" Then
                            sRef1 = sBuf1
                        Else
                            sRef1 = sBuf2
                        End If
                        
                        sBuf = gIFItem(i).s11 & ","
                        sBuf1 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sBuf2 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        
                        If sBuf2 = "" Then
                            sRef2 = sBuf1
                        Else
                            sRef2 = sBuf2
                        End If
                End Select
                
                sLimit1Gbn = gIFItem(i).s12
                sLimit1 = gIFItem(i).s13
                sLimit2Gbn = gIFItem(i).s14
                sLimit2 = gIFItem(i).s15
                
                Exit For
            End If
        Else
            If sIFSeq = gIFItem(i).s01 Then
                sJGbn = gIFItem(i).s09
                    
                Select Case sSex
                    Case "M"
                        sBuf = gIFItem(i).s10 & ","
                        sBuf = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sRef1 = sBuf
                        
                        sBuf = gIFItem(i).s11 & ","
                        sBuf = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sRef2 = sBuf
                    Case "F"
                        sBuf = gIFItem(i).s10 & ","
                        sBuf1 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sBuf2 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        
                        If sBuf2 = "" Then
                            sRef1 = sBuf1
                        Else
                            sRef1 = sBuf2
                        End If
                        
                        sBuf = gIFItem(i).s11 & ","
                        sBuf1 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        sBuf2 = GetByOneUserSymbol(sBuf, sBuf, ",")
                        
                        If sBuf2 = "" Then
                            sRef2 = sBuf1
                        Else
                            sRef2 = sBuf2
                        End If
                End Select
                
                sLimit1Gbn = gIFItem(i).s12
                sLimit1 = gIFItem(i).s13
                sLimit2Gbn = gIFItem(i).s14
                sLimit2 = gIFItem(i).s15
                
                Exit For
            End If
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If sIFRstCd = gCalItem(i).s01 Then
            sJGbn = gCalItem(i).s07
            
            Select Case sSex
                Case "M"
                    sBuf = gCalItem(i).s08 & ","
                    sBuf = GetByOneUserSymbol(sBuf, sBuf, ",")
                    sRef1 = sBuf
                    
                    sBuf = gCalItem(i).s09 & ","
                    sBuf = GetByOneUserSymbol(sBuf, sBuf, ",")
                    sRef1 = sBuf
                Case "F"
                    sBuf = gCalItem(i).s08 & ","
                    Call GetByOneUserSymbol(sBuf, sBuf, ",")
                    sBuf = GetByOneUserSymbol(sBuf, sBuf, ",")
                    sRef1 = sBuf
                    
                    sBuf = gCalItem(i).s09 & ","
                    Call GetByOneUserSymbol(sBuf, sBuf, ",")
                    sBuf = GetByOneUserSymbol(sBuf, sBuf, ",")
                    sRef1 = sBuf
            End Select
            
                        
            Exit For
        End If
    Next
    
    If IsNumeric(sCompRst) = False Then
        If (sCompRst = "LOWER LIMIT" Or sCompRst = "UPPER LIMIT") Then
            JudgeResultWithSex = sCompRst
            
            If sCompRst = "LOWER LIMIT" Then
                If sLimit1 <> "" Then
                    Select Case sLimit1Gbn
                        Case "0"
                            JudgeResultWithSex = sLimit1
                        Case "1"
                            JudgeResultWithSex = "< " & sLimit1
                        Case "2"
                            JudgeResultWithSex = sLimit1 & " 이하"
                        Case Else
                    End Select
                End If
            ElseIf sCompRst = "UPPER LIMIT" Then
                If sLimit2 <> "" Then
                    Select Case sLimit2Gbn
                        Case "0"
                            JudgeResultWithSex = sLimit2
                        Case "1"
                            JudgeResultWithSex = "> " & sLimit2
                        Case "2"
                            JudgeResultWithSex = sLimit2 & " 이상"
                        Case Else
                    End Select
                End If
            End If
            
            sOneRst2 = ""
            
            Exit Function
        Else
            JudgeResultWithSex = sCompRst
            sOneRst2 = Chr$(124)
            Exit Function
        End If
    End If
        
    Select Case sJGbn
        Case "0"
            JudgeResultWithSex = sCompRst
            sOneRst2 = ""
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultWithSex = sCompRst
                sOneRst2 = "Low"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                JudgeResultWithSex = sCompRst
                sOneRst2 = ""
            Else
                JudgeResultWithSex = sCompRst
                sOneRst2 = "High"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultWithSex = "Negative"
                sOneRst2 = "Negative"
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResultWithSex = "Positive"
                sOneRst2 = "Positive"
            Else
                JudgeResultWithSex = "GrayZone(+/-)"
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultWithSex = sCompRst
                sOneRst2 = "Negative"
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResultWithSex = sCompRst
                sOneRst2 = "Positive"
            Else
                JudgeResultWithSex = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "4"
        '이하 / 이상
            If IsNumeric(sCompRst) = True Then
                If Val(sCompRst) <= Val(sRef1) Then
                    JudgeResultWithSex = "<" & sRef1
                    sOneRst2 = "이하"
                ElseIf Val(sCompRst) > Val(sRef1) And Val(sCompRst) < Val(sRef2) Then
                    JudgeResultWithSex = sCompRst
                    sOneRst2 = ""
                Else
                    JudgeResultWithSex = ">" & sRef2
                    sOneRst2 = "이상"
                End If
            Else
                If sCompRst = "LOWER LIMIT" Then
                    If sRef1 = "" Then
                    Else
                        JudgeResultWithSex = "<" & sRef1
                        sOneRst2 = "이하"
                    End If
                ElseIf sCompRst = "UPPER LIMIT" Then
                    If sRef2 = "" Then
                    Else
                        JudgeResultWithSex = ">" & sRef2
                        sOneRst2 = "이상"
                    End If
                End If
            End If
        Case "5"
        'QAL P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultWithSex = "Positive"
                sOneRst2 = "Positive"
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResultWithSex = "Negative"
                sOneRst2 = "Negative"
            Else
                JudgeResultWithSex = "GrayZone(+/-)"
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "6"
        'QAN P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultWithSex = sCompRst
                sOneRst2 = "Positive"
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResultWithSex = sCompRst
                sOneRst2 = "Negative"
            Else
                JudgeResultWithSex = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "7"
        'P/N 장비
        
        Case Else
        
    End Select
    
    'LIMIT구분에 따른 처리
    If sLimit1 <> "" And Val(sCompRst) < Val(sLimit1) Then
        Select Case sLimit1Gbn
            Case "0"
                JudgeResultWithSex = sLimit1
            Case "1"
                JudgeResultWithSex = "< " & sLimit1
            Case "2"
                JudgeResultWithSex = sLimit1 & " 이하"
        End Select
    End If
    
    If sLimit2 <> "" And Val(sCompRst) > Val(sLimit2) Then
        Select Case sLimit2Gbn
            Case "0"
                JudgeResultWithSex = sLimit2
            Case "1"
                JudgeResultWithSex = "> " & sLimit2
            Case "2"
                JudgeResultWithSex = sLimit2 & " 이상"
        End Select
    End If
    
    Exit Function
    
ErrHandler:
    ViewMsg "JudgeResultWithSex - Err(" & Err.Description & ")"
End Function

Public Function JudgeRstBySex$(ByVal sIFSeq$, ByVal sCompRst$, ByVal sSex$, sRefFlag$)
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sJGbn$, sRef1$, sRef2$, sLimit1$, sLimit1Gbn$, sLimit2$, sLimit2Gbn$
    
    If Len(sIFSeq) = 3 Then
        For i = 1 To giOriginIFItemCnt
            If sIFSeq = gIFItem(i).s01 Then
                sJGbn = gIFItem(i).s09
                
                If sSex = "M" Then
                    sRef1 = gIFItem(i).s10
                    sRef2 = gIFItem(i).s11
                ElseIf sSex = "F" Then
                    sRef1 = gIFItem(i).s12
                    sRef2 = gIFItem(i).s13
                End If
                
                sLimit1Gbn = gIFItem(i).s19
                sLimit1 = gIFItem(i).s20
                sLimit2Gbn = gIFItem(i).s21
                sLimit2 = gIFItem(i).s22
                
                Exit For
            End If
        Next
    ElseIf Len(sIFSeq) = 2 Then
        For i = 1 To giOriginCalItemCnt
            If sIFSeq = gCalItem(i).s01 Then
                sJGbn = gCalItem(i).s07
                
                If sSex = "M" Then
                    sRef1 = gCalItem(i).s08
                    sRef2 = gCalItem(i).s09
                ElseIf sSex = "F" Then
                    sRef1 = gCalItem(i).s10
                    sRef2 = gCalItem(i).s11
                End If
                
                Exit For
            End If
        Next
    End If
    
    sRefFlag = ""
    
    If Trim(sCompRst) = "" Then Exit Function
    
    If IsNumeric(sCompRst) = False Then
        If sJGbn <> "0" Then
            sRefFlag = "H"
        End If
        
        JudgeRstBySex = sCompRst
    
        Exit Function
    End If
        
    Select Case sJGbn
        Case "0"
            JudgeRstBySex = sCompRst
            sRefFlag = ""
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeRstBySex = sCompRst
                sRefFlag = "L"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                JudgeRstBySex = sCompRst
                sRefFlag = ""
            Else
                JudgeRstBySex = sCompRst
                sRefFlag = "H"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeRstBySex = "NEGATIVE"
                sRefFlag = ""
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeRstBySex = "POSITIVE"
                sRefFlag = ""
            Else
                JudgeRstBySex = "TRACE"
                sRefFlag = ""
            End If
            
            Exit Function
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeRstBySex = sCompRst
                sRefFlag = "N"
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeRstBySex = sCompRst
                sRefFlag = "P"
            Else
                JudgeRstBySex = sCompRst
                sRefFlag = "T"
            End If
            
        Case "4"
            
        Case "5"
        'QAL P/N
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeRstBySex = "POSITIVE"
                sRefFlag = ""
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeRstBySex = "NEGATIVE"
                sRefFlag = ""
            Else
                JudgeRstBySex = "TRACE"
                sRefFlag = ""
            End If
            
            Exit Function
        Case "6"
        'QAN P/N
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeRstBySex = sCompRst
                sRefFlag = "P"
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeRstBySex = sCompRst
                sRefFlag = "N"
            Else
                JudgeRstBySex = sCompRst
                sRefFlag = "T"
            End If
        
        Case Else
        
    End Select
    
    If IsNumeric(sCompRst) = False Then
        Exit Function
    End If
    
    'LIMIT구분에 따른 LIMIT 처리
    If IsNumeric(sLimit1) = True And sLimit1 <> "" And sLimit1Gbn <> "" Then
        If Val(sCompRst) <= Val(sLimit1) Then
            Select Case sLimit1Gbn
                Case "0"
                    '1.0
                    JudgeRstBySex = sLimit1
                Case "1"
                    '< 1.0
                    JudgeRstBySex = "< " & sLimit1
                Case "2"
                    '1.0 이하
                    JudgeRstBySex = sLimit1 & " 이하"
            End Select
        End If
    End If
    
    If IsNumeric(sLimit2) = True And sLimit2 <> "" And sLimit2Gbn <> "" Then
        If Val(sCompRst) >= Val(sLimit2) Then
            Select Case sLimit2Gbn
                Case "0"
                    '1.0
                    JudgeRstBySex = sLimit2
                Case "1"
                    '> 1.0
                    JudgeRstBySex = "> " & sLimit2
                Case "2"
                    '1.0 이상
                    JudgeRstBySex = sLimit2 & " 이상"
            End Select
        End If
    End If
    
    Exit Function
    
ErrHandler:
    ViewMsg "JudgeRstBySex - Err(" & Err.Description & ")"
End Function

Public Sub MakeIFItemStruct(ByVal sIFItem As String, ByVal iCnt As Integer)
    Dim i%
    Dim sDataRow() As String
    Dim sOneRow As String
    
    ReDim gIFItem(iCnt)
    ReDim sDataRow(iCnt) As String
    
    For i = 1 To iCnt
        sDataRow(i) = GetByOneUserSymbol(sIFItem, sIFItem, Chr(3))
    Next
    
    For i = 1 To iCnt
        sOneRow = sDataRow(i) & Chr(124)
        
        gIFItem(i).s01 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s02 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s03 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s04 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s05 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s06 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s07 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s08 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s09 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s10 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s11 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s12 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s13 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s14 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s15 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s16 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s17 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s18 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s19 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s20 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s21 = GetByOne(sOneRow, sOneRow)
        gIFItem(i).s22 = GetByOne(sOneRow, sOneRow)
    Next
End Sub

Public Sub MakeCalItemStruct(ByVal sCalItem As String, ByVal iCnt As Integer)
    Dim i%
    Dim sDataRow() As String
    Dim sOneRow As String
    
    ReDim gCalItem(iCnt)
    ReDim sDataRow(iCnt) As String
    
    For i = 1 To iCnt
        sDataRow(i) = GetByOneUserSymbol(sCalItem, sCalItem, Chr$(3))
    Next
    
    For i = 1 To iCnt
        sOneRow = sDataRow(i) & Chr(124)
        
        gCalItem(i).s01 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s02 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s03 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s04 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s05 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s06 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s07 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s08 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s09 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s10 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s11 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s12 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s13 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s14 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s15 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s16 = GetByOne(sOneRow, sOneRow)
    Next
End Sub

Public Function RackFormat(ByVal sRackDig As String) As String
    If sRackDig = "0" Then
    ElseIf sRackDig = "1" Then
        RackFormat = "0"
    ElseIf sRackDig = "2" Then
        RackFormat = "00"
    ElseIf sRackDig = "3" Then
        RackFormat = "000"
    ElseIf sRackDig = "4" Then
        RackFormat = "0000"
    ElseIf sRackDig = "5" Then
        RackFormat = "00000"
    ElseIf sRackDig = "6" Then
        RackFormat = "000000"
    ElseIf sRackDig = "7" Then
        RackFormat = "0000000"
    ElseIf sRackDig = "8" Then
        RackFormat = "00000000"
    ElseIf sRackDig = "9" Then
        RackFormat = "000000000"
    End If
End Function

Public Sub RegEditCurFrmTitle(ByVal sGbn As String, ByVal sBuf As String)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "WndTitle." & sGbn, sBuf)
    
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
End Sub

Public Sub RegUserInfo(ByVal sUID As String, ByVal sPWD As String, ByVal sUserNm As String, ByVal sUserOther As String)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "User.Id", sUID)
                    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "User.Pwd", sPWD)
                    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "User.Nm", sUserNm)
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "User.Other", sUserOther)
End Sub

Public Function RegServerOK(ByVal iCRow%, ByVal iRstCnt%, ByVal sIFRstCd$, ByVal sRst1$, ByVal sRst2$) As String
    
    Dim sBuf$, sCRst1$, sCRst2$, sCRstCd$, sCSvrCd$, sTIFSeq$, sTSvrCd$, sTRst1$, sTRst2$, sRetVal$, sTmp$, sIFSeq$
    Dim iTRstCnt%, i%, j%
    Dim vWSeq, vJDate, vJGbn, vJNo, vRack, vPos, vRegNo, vPtNm, vSex, vEmer, vRerun, vOther, vTmp, vIFItemCnt
    Dim objRst As Object
    
    '결과등록 DLL을 Call하여 서버쪽에 결과등록함
    sBuf = gRstcfg.sComponent

    If sBuf = "" Then
        ViewMsg "서버에 결과등록을 위한 DLL 파일이 존재하지 않습니다!!"
        Exit Function
    End If
    
    Set objRst = CreateObject(sBuf)
    
    With gfIFDisplayForm.spdIntList
        Call .GetText(1, iCRow, vWSeq)
        Call .GetText(3, iCRow, vJDate)
        Call .GetText(4, iCRow, vJGbn)
        Call .GetText(5, iCRow, vJNo)
        Call .GetText(6, iCRow, vRack)
        Call .GetText(7, iCRow, vPos)
        Call .GetText(8, iCRow, vRegNo)
        Call .GetText(9, iCRow, vPtNm)
        Call .GetText(10, iCRow, vSex)
        Call .GetText(11, iCRow, vEmer)
        Call .GetText(12, iCRow, vRerun)
        Call .GetText(13, iCRow, vOther)
    End With
    
    If vWSeq = "" Then
        Exit Function
    End If
    
    If Trim(vJDate) = "" Then
        vJDate = Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD")
    End If
    
    'QC인 경우는 제외
    If Len(vJNo) <> 13 Or Mid(vJNo, 7, 1) = "Q" Then
        Set objRst = Nothing
        Exit Function
    End If
    
    
    sTIFSeq = ""
    sTSvrCd = ""
    sTRst1 = ""
    iTRstCnt = 0
            
    'ServerCd로 변환 - 서버쪽코드가 존재하는 것만 등록
    For i = 1 To iRstCnt
        sCRstCd = GetByOne(sIFRstCd, sIFRstCd)
        sCSvrCd = ""

        sCRst1 = GetByOne(sRst1, sRst1)
        sCRst2 = GetByOne(sRst2, sRst2)
        
        With gfIFDisplayForm.spdIntList
            Call .GetText(16, iCRow, vIFItemCnt)
            
            For j = 1 To CInt(vIFItemCnt)
                Call .GetText(16 + j, iCRow, vTmp)
                
                sTmp = CStr(vTmp)
                
                sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
                
                'IFSeq를 IFRstCd로 Convert
                If Len(sIFSeq) = 2 And Left(sIFSeq, 1) = "C" Then
                '계산식일때는 원래가 IFSeq임
                    If sIFSeq = sCRstCd Then
                        'IFSeq를 서버쪽코드로 Convert
                        sCSvrCd = ConvertIFItemInfo(2, sIFSeq)
                        Exit For
                    End If
                Else
                '일반항목의 경우
                    If ConvertIFItemInfo(8, sIFSeq) = sCRstCd Then
                        'IFSeq를 서버쪽코드로 Convert
                        sCSvrCd = ConvertIFItemInfo(2, sIFSeq)
                        Exit For
                    End If
                End If
            Next
        End With
        
        If sCSvrCd = "" Then
        Else
            iTRstCnt = iTRstCnt + 1
            sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
            sTSvrCd = sTSvrCd & sCSvrCd & Chr(124)
            sTRst1 = sTRst1 & sCRst1 & Chr(124)
            sTRst2 = sTRst2 & sCRst2 & Chr(124)
        End If
    Next
    
    '서버등록 실행
    Call objRst.SetMachineInfo(gsMachineCd, gsMachineNm)

    sRetVal = objRst.RegServer(1, Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD"), CStr(vWSeq) & Chr(124), _
                CStr(vJDate) & Chr(124), CStr(vJGbn) & Chr(124), CStr(vJNo) & Chr(124), _
                CStr(vRack) & Chr(124), CStr(vPos) & Chr(124), _
                CStr(vRegNo) & Chr(124), CStr(vPtNm) & Chr(124), CStr(vSex) & Chr(124), _
                CStr(vEmer) & Chr(124), CStr(vRerun) & Chr(124), CStr(vOther) & Chr(3), _
                CStr(iTRstCnt) & Chr(124), sTIFSeq & Chr(3), sTSvrCd & Chr(3), sTRst1 & Chr(3), sTRst2 & Chr(3), _
                ADOCN1, ADOCN2, gSvrInfo.DBGbn)
                
    If sRetVal = "OK" Then
        ViewMsg CStr(vJNo) & "의 결과를 서버에 저장하였습니다!!"
    Else
        ViewMsgLog "서버 ERR : " & CStr(vJNo)
    End If
    
    Set objRst = Nothing
End Function

Public Sub RegOrder(ByVal iMode As Integer)
    On Error GoTo ErrHandler

    Dim objld As Object
    Dim i%
    Dim sTIFSeq$
    
    sTIFSeq = ""
    
    For i = 1 To gOrderTable.iOrdCnt
        sTIFSeq = sTIFSeq & gOrderTable.sIFSeq(i) & Chr(124)
    Next
    
    Set objld = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    Call objld.Add_IFResult(gsMachineCd, iMode, gOrderTable.sWDate, gOrderTable.sWSeq, sTIFSeq, _
                gOrderTable.sJDate, gOrderTable.sJGbn, gOrderTable.sJNo, _
                gOrderTable.sRack, gOrderTable.sPos, _
                gOrderTable.sRegNo, gOrderTable.sName, gOrderTable.sSex, _
                gOrderTable.sEmer, gOrderTable.sReRun, gOrderTable.sOther, "", "", "", "0", _
                frmInterface.optRegOpt(0).Value, gOrderTable.iOrdCnt)
                
    Set objld = Nothing
    
    'LastWSeq를 갱신
    gsLastWSeq = gOrderTable.sWSeq
    
    Exit Sub
    
ErrHandler:
    Set objld = Nothing
    ViewMsg "RegOrder 오류 - (" & Err.Description & ")"
End Sub

Private Sub SendResultSocket(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
                            ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                            ByVal sIFSpcCd As String, Optional ByVal iCnt As Integer)
    On Error GoTo ErrHandler
    
    Dim vIFItemCnt, vTmp, vChk
    Dim i%, j%, k%, iExist%
    Dim sTmp$, sCIFRstCd$, sCRst1$, sCRst2$, sCFlag$
    Dim sWDate$, sWSeq$, sJDate$, sJGbn$, sJNo$, sRack$, sPos$, sRegNo$, sName$, sSex$, sEmer$, sReRun$, sOther$
    Dim sTIFSeq$, sTRst1$, sTRst2$, sTFlag$
    Dim sIFSeq$, sRtnVal$
    
    With gfIFDisplayForm.spdIntList
        sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")

        Call .GetText(1, CInt(sCRow), vTmp)
        sWSeq = CStr(vTmp)

        Call .GetText(3, CInt(sCRow), vTmp)
        sJDate = CStr(vTmp)
        
        Call .GetText(4, CInt(sCRow), vTmp)
        sJGbn = CStr(vTmp)
        
        Call .GetText(5, CInt(sCRow), vTmp)
        sJNo = CStr(vTmp)
        
        Call .GetText(6, CInt(sCRow), vTmp)
        sRack = CStr(vTmp)
        
        Call .GetText(7, CInt(sCRow), vTmp)
        sPos = CStr(vTmp)
        
        Call .GetText(8, CInt(sCRow), vTmp)
        sRegNo = CStr(vTmp)
        
        Call .GetText(9, CInt(sCRow), vTmp)
        sName = CStr(vTmp)
        
        Call .GetText(10, CInt(sCRow), vTmp)
        sSex = CStr(vTmp)
        
        Call .GetText(11, CInt(sCRow), vTmp)
        sEmer = CStr(vTmp)
        
        Call .GetText(12, CInt(sCRow), vTmp)
        sReRun = CStr(vTmp)
        
        Call .GetText(13, CInt(sCRow), vTmp)
        sOther = CStr(vTmp)
        

    'iMode = 1 ---> 한 샘플씩 LOCAL 등록
        Call .GetText(16, CInt(sCRow), vIFItemCnt)
        
        For i = 1 To CInt(vIFItemCnt)
            Call .GetText(16 + i, CInt(sCRow), vTmp)
            
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
            sRst1 = GetByOne(sTmp, sTmp)
            sRst2 = GetByOne(sTmp, sTmp)
            
            sTIFSeq = sTIFSeq & sIFSeq & "|"
            sTRst1 = sTRst1 & sRst1 & "|"
            sTRst2 = sTRst2 & sRst2 & "|"
        Next i
    End With
    
'    If Mid(sJNo, 7, 1) <> "Q" Then
    If Trim(sJNo) <> "" And Mid(sJNo, 7, 1) <> "Q" Then     '2004/6/10 yk
        '--- 결과등록 프로그램에 메세지 전송(2003/3/17 yk)
        Dim sSendMsg    As String
        
        sSendMsg = "R" & Chr(3) & sWDate & Chr(3) & sWSeq & Chr(3) & sJNo & Chr(3) _
                & sTIFSeq & Chr(3) & sTRst1 & Chr(3) & sOther & Chr(4)
        
        With frmInterface.Winsock1
            If .State = sckConnected Then
                .SendData sSendMsg
            End If
        End With
        '-------------------------------------------------
    End If
    
    Exit Sub
ErrHandler:
    ViewMsg "SendResultSocket 오류 - (" & Err.Description & ")"
End Sub

Private Function RegResultTemp(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
                        ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                        ByVal sIFSpcCd As String, Optional ByVal iCnt As Integer) As String
    On Error GoTo ErrHandler
    
    'iMode = 1 ---> 한 샘플씩 자동 등록
    
    Dim vIFItemCnt, vTmp, vChk
    Dim i%, j%, k%, iExist%
    Dim sTmp$, sCIFRstCd$, sCRst1$, sCRst2$, sCFlag$
    Dim sWDate$, sWSeq$, sJDate$, sJGbn$, sJNo$, sRack$, sPos$, sRegNo$, sName$, sSex$, sEmer$, sReRun$, sOther$
    Dim sTIFSeq$, sTRst1$, sTRst2$, sTFlag$
    Dim sIFSeq$, sRtnVal$
    Dim objld   As Object
    
    Set objld = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    With gfIFDisplayForm.spdIntList
        sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")

        Call .GetText(1, CInt(sCRow), vTmp)
        sWSeq = CStr(vTmp)

        Call .GetText(3, CInt(sCRow), vTmp)
        sJDate = CStr(vTmp)
        
        Call .GetText(4, CInt(sCRow), vTmp)
        sJGbn = CStr(vTmp)
        
        Call .GetText(5, CInt(sCRow), vTmp)
        sJNo = CStr(vTmp)
        
        Call .GetText(6, CInt(sCRow), vTmp)
        sRack = CStr(vTmp)
        
        Call .GetText(7, CInt(sCRow), vTmp)
        sPos = CStr(vTmp)
        
        Call .GetText(8, CInt(sCRow), vTmp)
        sRegNo = CStr(vTmp)
        
        Call .GetText(9, CInt(sCRow), vTmp)
        sName = CStr(vTmp)
        
        Call .GetText(10, CInt(sCRow), vTmp)
        sSex = CStr(vTmp)
        
        Call .GetText(11, CInt(sCRow), vTmp)
        sEmer = CStr(vTmp)
        
        Call .GetText(12, CInt(sCRow), vTmp)
        sReRun = CStr(vTmp)
        
        Call .GetText(13, CInt(sCRow), vTmp)
        sOther = CStr(vTmp)
        

    'iMode = 1 ---> 한 샘플씩 LOCAL 등록
        Call .GetText(16, CInt(sCRow), vIFItemCnt)
        
        For i = 1 To CInt(vIFItemCnt)
            Call .GetText(16 + i, CInt(sCRow), vTmp)
            
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
            sRst1 = GetByOne(sTmp, sTmp)
            sRst2 = GetByOne(sTmp, sTmp)
            
            sTIFSeq = sTIFSeq & sIFSeq & "|"
            sTRst1 = sTRst1 & sRst1 & "|"
            sTRst2 = sTRst2 & sRst2 & "|"
        Next i
        
        sRtnVal = objld.Add_TEMPResult(gsMachineCd, 1, sWDate, sWSeq, _
                                        sTIFSeq, sJNo, sRegNo, sName, _
                                        sOther, sTRst1, sTRst2, CInt(Val(vIFItemCnt)))
        
        If IsNumeric(sRtnVal) = False Then
            If Mid(sJNo, 7, 1) <> "Q" Then
                '--- 결과등록 프로그램에 메세지 전송(2003/3/17 yk)
                Dim sSendMsg    As String
                
                sSendMsg = "R" & Chr(124) & sWDate & Chr(124) & sWSeq & Chr(124) & sJNo & Chr(124) & sOther
                
                With frmInterface.Winsock1
                    If .State = sckConnected Then
                        .SendData sSendMsg
                    End If
                End With
                '-------------------------------------------------
            End If
        Else
            If Len(sJNo) > 0 Then
                If Len(sJDate) = 0 Then
                    If Len(sJGbn) = 0 Then
                        ViewMsg sJNo & "의 저장에 실패하였습니다..."
                    Else
                        ViewMsg sJGbn & "-" & sJNo & "의 저장에 실패하였습니다..."
                    End If
                ElseIf Len(sJGbn) = 0 Then
                    If Len(sJDate) = 0 Then
                        ViewMsg sJNo & "의 저장에 실패하였습니다..."
                    Else
                        ViewMsg sJDate & "-" & sJNo & "의 저장에 실패하였습니다..."
                    End If
                Else
                    ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "의 저장에 실패하였습니다..."
                End If
            Else
                ViewMsg sWDate & "-" & sWSeq & "의 저장에 실패하였습니다..."
            End If
        End If
    End With
    
    Set objld = Nothing
    
    Exit Function
ErrHandler:
    Set objld = Nothing
    ViewMsg "RegResultTemp 오류 - (" & Err.Description & ")"
End Function


Public Function RegResult(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
                        ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                        ByVal sIFSpcCd As String, ByVal sFlag As String, Optional ByVal iCnt As Integer) As String
    
    On Error GoTo ErrHandler
    
    'iMode = 0 ---> 한 검사항목의 결과를 자동 등록
    'iMode = 1 ---> 한 샘플씩 자동 등록
    'iMode = 2 ---> Batch방식에 사용 여러 샘플 한 번에 등록
    
    Dim vIFItemCnt, vTmp, vChk
    Dim i%, j%, k%, iExist%
    Dim sTmp$, sCIFRstCd$, sCRst1$, sCRst2$, sCFlag$
    Dim sWDate$, sWSeq$, sJDate$, sJGbn$, sJNo$, sRack$, sPos$, sRegNo$, sName$, sSex$, sEmer$, sReRun$, sOther$
    Dim sTIFSeq$, sTRst1$, sTRst2$, sTFlag$
    Dim sIFSeq$, sRtnVal$
    Dim objld   As Object
    Dim bAutoFlag   As Boolean
    
    bAutoFlag = gfIFDisplayForm.optRegOpt(0).Value
    
    Set objld = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    With gfIFDisplayForm.spdIntList
        Select Case iMode
            Case 0, 1
                sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")
        
                Call .GetText(1, CInt(sCRow), vTmp)
                sWSeq = CStr(vTmp)
        
                Call .GetText(3, CInt(sCRow), vTmp)
                sJDate = CStr(vTmp)
                
                Call .GetText(4, CInt(sCRow), vTmp)
                sJGbn = CStr(vTmp)
                
                Call .GetText(5, CInt(sCRow), vTmp)
                sJNo = CStr(vTmp)
                
                Call .GetText(6, CInt(sCRow), vTmp)
                sRack = CStr(vTmp)
                
                Call .GetText(7, CInt(sCRow), vTmp)
                sPos = CStr(vTmp)
                
                Call .GetText(8, CInt(sCRow), vTmp)
                sRegNo = CStr(vTmp)
                
                Call .GetText(9, CInt(sCRow), vTmp)
                sName = CStr(vTmp)
                
                Call .GetText(10, CInt(sCRow), vTmp)
                sSex = CStr(vTmp)
                
                Call .GetText(11, CInt(sCRow), vTmp)
                sEmer = CStr(vTmp)
                
                Call .GetText(12, CInt(sCRow), vTmp)
                sReRun = CStr(vTmp)
                
                Call .GetText(13, CInt(sCRow), vTmp)
                sOther = CStr(vTmp)
                
            'iMode = 0 ---> 한 검사항목의 결과를 LOCAL 등록
                If iMode = 0 Then
                    sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                    sCRst1 = GetByOne(sRst1, sRst1)
                    sCRst2 = GetByOne(sRst2, sRst2)
                    sCFlag = GetByOne(sFlag, sFlag)
                    
                    Call .GetText(16, CInt(sCRow), vIFItemCnt)
                    
                    For i = 1 To CInt(vIFItemCnt)
                        Call .GetText(16 + i, CInt(sCRow), vTmp)
                        
                        sTmp = CStr(vTmp)
                        
                        iExist = 0
                        
                        sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
                        
                        If Len(sIFSeq) = 3 Then
                            If ConvertIFItemInfo(8, sIFSeq) = sCIFRstCd Then
                                iExist = 1
                                
                                Exit For
                            End If
                        ElseIf Len(sIFSeq) = 2 Then
                                iExist = 1
                                
                            Exit For
                        End If
                    Next
                    
                    If iExist = 1 Then
                        sRtnVal = objld.Add_IFResult(gsMachineCd, 0, sWDate, sWSeq, _
                                      sIFSeq, sJDate, sJGbn, sJNo, sRack, sPos, sRegNo, sName, sSex, sEmer, sReRun, _
                                      sOther, sCRst1, sCRst2, sCFlag, "0", bAutoFlag, iRstCnt)
                    End If
                    
                    If IsNumeric(sRtnVal) = False Then
                        If Len(sJNo) > 0 Then
                            If Len(sJDate) = 0 Then
                                If Len(sJGbn) = 0 Then
                                    ViewMsg sJNo & "의 결과를 저장하였습니다..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "의 결과를 저장하였습니다..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "의 결과를 저장하였습니다..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "의 결과를 저장하였습니다..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "의 결과를 저장하였습니다..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "의 결과를 저장하였습니다..."
                        End If
                        
                        Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, CInt(sCRow), CInt(sCRow), _
                                RGB(0, 0, 0), 연초록)
                        
                    Else
                        If Len(sJNo) > 0 Then
                            If Len(sJDate) = 0 Then
                                If Len(sJGbn) = 0 Then
                                    ViewMsg sJNo & "의 저장에 실패하였습니다..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "의 저장에 실패하였습니다..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "의 저장에 실패하였습니다..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "의 저장에 실패하였습니다..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "의 저장에 실패하였습니다..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "의 저장에 실패하였습니다..."
                        End If
                    End If
                End If

            'iMode = 1 ---> 한 샘플씩 LOCAL 등록
                If iMode = 1 Then
                    Call .GetText(16, CInt(sCRow), vIFItemCnt)
                    
                    For i = 1 To CInt(vIFItemCnt)
                        Call .GetText(16 + i, CInt(sCRow), vTmp)
                        
                        sTmp = CStr(vTmp)
                        
                        sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
                        sRst1 = GetByOne(sTmp, sTmp)
                        sRst2 = GetByOne(sTmp, sTmp)
                        sFlag = GetByOne(sTmp, sTmp)
                        
                        sTIFSeq = sTIFSeq & sIFSeq & "|"
                        sTRst1 = sTRst1 & sRst1 & "|"
                        sTRst2 = sTRst2 & sRst2 & "|"
                        sTFlag = sTFlag & sFlag & "|"
                    Next i
                    
                    sRtnVal = objld.Add_IFResult(gsMachineCd, 1, sWDate, sWSeq, _
                                  sTIFSeq, sJDate, sJGbn, sJNo, sRack, sPos, sRegNo, sName, sSex, sEmer, sReRun, _
                                  sOther, sTRst1, sTRst2, sTFlag, "0", bAutoFlag, CInt(Val(vIFItemCnt)))
                    
                    If IsNumeric(sRtnVal) = False Then
                        If Len(sJNo) > 0 Then
                            If Len(sJDate) = 0 Then
                                If Len(sJGbn) = 0 Then
                                    ViewMsg sJNo & "의 결과를 저장하였습니다..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "의 결과를 저장하였습니다..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "의 결과를 저장하였습니다..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "의 결과를 저장하였습니다..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "의 결과를 저장하였습니다..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "의 결과를 저장하였습니다..."
                        End If
                    
                        Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, CInt(sCRow), CInt(sCRow), _
                                RGB(0, 0, 0), 연초록)
                        
                    Else
                        If Len(sJNo) > 0 Then
                            If Len(sJDate) = 0 Then
                                If Len(sJGbn) = 0 Then
                                    ViewMsg sJNo & "의 저장에 실패하였습니다..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "의 저장에 실패하였습니다..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "의 저장에 실패하였습니다..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "의 저장에 실패하였습니다..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "의 저장에 실패하였습니다..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "의 저장에 실패하였습니다..."
                        End If
                    End If
                End If
                
        'iMode = 2 ---> 여러 샘플 Batch 등록
            Case 2
        
            Case Else
        End Select
    End With
    
    Set objld = Nothing
    
    Exit Function
    
ErrHandler:
    Set objld = Nothing
    ViewMsg "RegResult 오류 - (" & Err.Description & ")"
End Function


Public Sub RegViewMsgHwnd(ByVal lnHwnd As Long)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "ViewMsg.Hwnd", CStr(lnHwnd))
    
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
End Sub

Public Function ReOrder_IFSeq_And_RealOrdCnt(ByVal sBuf As String, iRealOrdCnt As Integer)
    Dim aBuf() As String
    Dim sTmp$, sMax$, sMin$
    Dim iCnt%, i%, j%
    
    Do
        sTmp = GetByOne(sBuf, sBuf)
        
        If sTmp = "" Then
            Exit Do
        End If
        
        iCnt = iCnt + 1
        
        ReDim Preserve aBuf(iCnt)
        
        aBuf(iCnt) = sTmp
    Loop
    
    iRealOrdCnt = iCnt
    
    For i = 1 To MAXIFITEM
        For j = 1 To iCnt
            If Format(i, "000") = aBuf(j) Then
                ReOrder_IFSeq_And_RealOrdCnt = ReOrder_IFSeq_And_RealOrdCnt & aBuf(j) & Chr$(124)
            Else
            End If
        Next
    Next
End Function

Public Function RemoveDuplicatedOrder$(ByVal sBuf As String, iRealOrdCnt As Integer)
    Dim sTmp$
    Dim iCnt%, i%, iExist%
    Dim aBuf() As String
    
    sBuf = Trim(sBuf)
    
    If Right(sBuf, 1) <> "," Then
        sBuf = sBuf & ","
    End If
    
    iCnt = 0
    
    Do
        sTmp = GetByOneUserSymbol(sBuf, sBuf, ",")
        
        If sTmp = "" Then
            sBuf = Trim(sBuf)
            
            If sBuf = "" Then
                Exit Do
            End If
        Else
            iExist = 0
            
            If iCnt = 0 Then
            
            Else
                For i = 1 To iCnt
                    If aBuf(i) = sTmp Then
                        iExist = 1
                        Exit For
                    Else
                        iExist = 0
                    End If
                Next
            End If
            
            If iExist = 0 Then
                iCnt = iCnt + 1
                
                ReDim Preserve aBuf(iCnt)
                aBuf(iCnt) = sTmp
            End If
        End If
    Loop
    
    iRealOrdCnt = iCnt
    
    sTmp = ""
    
    For i = 1 To iCnt
        sTmp = sTmp & aBuf(i) & ","
    Next
    
    RemoveDuplicatedOrder = sTmp
End Function

Public Sub ResultSpdClear2()
    With gfIFDisplayForm.spdRst
        .BlockMode = True
        .Row = -1
        .Row2 = -1
        .Col = -1
        .Col2 = -1
        .Action = SS_ACTION_CLEAR_TEXT
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(0, 0, 0)
        .BlockMode = False
    End With
    
    With gfIFDisplayForm.spdRst2
        .BlockMode = True
        .Row = -1
        .Row2 = -1
        .Col = -1
        .Col2 = -1
        .Action = SS_ACTION_CLEAR_TEXT
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(0, 0, 0)
        .BlockMode = False
    End With
End Sub

Public Function SpdForeBack(SpdName As Object, ByVal lnCol1 As Long, ByVal lnCol2 As Long, _
                ByVal lnRow1 As Long, ByVal lnRow2 As Long, ByVal sFcolor As String, ByVal sBcolor As String)
    
    With SpdName
        .BlockMode = True
        .Col = lnCol1
        .Col2 = lnCol2
        .Row = lnRow1
        .Row2 = lnRow2
        .ForeColor = sFcolor
        .BackColor = sBcolor
        .BlockMode = False
    End With

End Function

Public Sub spdReverse(spdReverse As Object, ByVal lnCol1 As Long, ByVal lnCol2 As Long, _
                       ByVal lnRow1, ByVal lnRow2, ByVal sColor As String, Optional vOption As Variant)
    Dim i%
    Dim iMatchRow%
    
    iMatchRow = 0
    
    With spdReverse
        For i = 1 To .MaxRows
            If lnCol1 = -1 Then
                .Row = i
                .Col = 1
                If .BackColor = sColor Then
                    iMatchRow = i
                    Exit For
                End If
            Else
                .Row = i
                .Col = lnCol1
                If .BackColor = sColor Then
                    iMatchRow = i
                    Exit For
                End If
            End If
        Next
    End With
    
    If iMatchRow = 0 Then
    Else
        If vOption = 1 Then     '흰 바탕
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = RGB(255, 255, 255)
                .BlockMode = False
            End With
        End If
        
        If vOption = 2 Then     '하늘 계열 바탕
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = &HDFFFDF
                .BlockMode = False
            End With
        End If
        
        If vOption = 3 Then     '노란 계열 바탕
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = &HE0FFFF
                .BlockMode = False
            End With
        End If
        
        If vOption = 1 Or vOption = 2 Or vOption = 3 Then
        Else
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = CStr(vOption)
                .BlockMode = False
            End With
        End If
    End If
    
    With spdReverse
        .BlockMode = True
        .Col = lnCol1
        .Col2 = lnCol2
        .Row = lnRow1
        .Row2 = lnRow2
        .BackColor = sColor
        .BlockMode = False
    End With
End Sub

Public Function SubCompute(ByVal Op As String, ByVal Op1 As Double, ByVal Op2 As Double) As Double
    Select Case Op
        Case "+":   SubCompute = Op1 + Op2
        Case "-":   SubCompute = Op1 - Op2
        Case "\":   SubCompute = Op1 \ Op2
        Case "%":   SubCompute = Op1 Mod Op2
        Case "*":   SubCompute = Op1 * Op2
        Case "/":   SubCompute = Op1 / Op2
        Case "^":   SubCompute = Op1 ^ Op2
    End Select
    
End Function

Public Sub Txt_Highlight(SomeTextBox As TextBox)
    SomeTextBox.SelStart = 0
    SomeTextBox.SelLength = Len(SomeTextBox)
End Sub

Public Sub TxtTypeOnlyNumeric(SomeTextBox1 As TextBox, iKeyAscii As Integer)
    If (iKeyAscii >= 48 And iKeyAscii <= 57) Or iKeyAscii = 8 Then
    Else
        iKeyAscii = 0
    End If
End Sub

Public Function ViewIFResult2(ByVal iCRow As Integer, ByVal iRstCnt As Integer, _
    ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, ByVal sIFSpcCd As String) As String
    
    Dim i%, j%
    Dim sCIFRstCd$, sCRst1$, sCRst2$, sCIFSeq$, sTmp$
    Dim vTmp, vIFCnt, vCRstCnt
    
    With gfIFDisplayForm
        With .spdIntList
            Call .GetText(15, iCRow, vCRstCnt)
            
            Call .GetText(16, iCRow, vIFCnt)
            
            For j = 1 To CInt(Val(vIFCnt))
                Call .GetText(16 + j, iCRow, vTmp)
                
                sTmp = CStr(vTmp)
                sCIFSeq = GetByOne(sTmp, sTmp)
                gResultTable(1).sTestCd(j) = sCIFSeq
            Next
            
            For i = 1 To iRstCnt
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                
                For j = 1 To CInt(Val(vIFCnt))
                    '계산항목일 때 : sIFRstCd = sIFSeq
                    If Left(gResultTable(1).sTestCd(j), 1) = "C" Then
                        If gResultTable(1).sTestCd(j) = sCIFRstCd Then
                            If sCRst2 = "Low" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(220, 220, 255))
                            ElseIf sCRst2 = "High" Or sCRst2 = "Positive" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(255, 220, 220))
                            ElseIf sCRst2 = "L" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(220, 220, 255))
                            ElseIf sCRst2 = "H" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(255, 220, 220))
                            ElseIf sCRst2 <> "" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(230, 230, 230))
                            Else
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                        RGB(0, 0, 0), RGB(255, 255, 255))
                            End If
                            
                            Exit For
                        End If
                    Else
                    '일반항목일 때
                        If ConvertIFItemInfo(8, gResultTable(1).sTestCd(j)) = sCIFRstCd Then
                            If sCRst2 = "Low" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(220, 220, 255))
                            ElseIf sCRst2 = "High" Or sCRst2 = "Positive" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(255, 220, 220))
                            ElseIf sCRst2 = "L" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(220, 220, 255))
                            ElseIf sCRst2 = "H" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(255, 220, 220))
                            ElseIf sCRst2 <> "" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(230, 230, 230))
                            Else
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                        RGB(0, 0, 0), RGB(255, 255, 255))
                            End If
                            
                            Exit For
                        End If
                    End If
                Next
            Next
        End With
        
'        'Result spdRst에 표시
'        Call DisplayResult2(iCRow)
        
        If Val(vCRstCnt) >= Val(vIFCnt) Then
            ViewIFResult2 = "DONE"
        Else
            ViewIFResult2 = "MORE"
        End If
    End With
End Function

Public Sub ViewMsg(ByVal sMsg As String)
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "ViewMsg.Hwnd")
    
    Call SetWindowText(Val(sBuf), sMsg)
End Sub

Public Sub ViewMsgLog(ByVal sMsg$)
    Dim i%, iExist%
    
    iExist = 0
    
    With gfIFDisplayForm
        For i = 1 To .listNoOrd.ListCount
            If sMsg = .listNoOrd.List(i - 1) Then
                iExist = 1
            End If
        Next
        
        If iExist = 0 Then
            .listNoOrd.AddItem sMsg
        End If
    End With
End Sub

Public Sub LogFileOpen()
    On Error GoTo ErrHandler
    
    Open App.Path & "\" & gsMachineNm & ".log" For Output Shared As #1
    Open App.Path & "\" & gsMachineNm & "Buf.log" For Output Shared As #2
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Number
    MsgBox Err.Description, vbCritical, "Log File Open Error"
End Sub

Public Sub LogFileClose()
    On Error GoTo ErrHandler
        
    Close #1
    Close #2
    
    Exit Sub
    
ErrHandler:
End Sub

Public Sub PortOpen()
    On Error GoTo ErrHandler

    With frmInterface.Comm1
        .CommPort = gCommInfo.sPort
        .Settings = gCommInfo.sBaudRate & "," & gCommInfo.sParity & "," & _
                      gCommInfo.sDataBit & "," & gCommInfo.sStopBIt
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Number
    MsgBox Err.Description, vbCritical, "Port Open Error"
End Sub

Public Sub PortClose()
    On Error GoTo ErrHandler
    
    If frmInterface.Comm1.PortOpen = True Then
        frmInterface.Comm1.PortOpen = False
    End If
    
    Exit Sub
    
ErrHandler:
End Sub

Public Sub DisplayInitItem()
    Dim i%, j%
    Dim iCurItemCnt%
    
    For i = 1 To MAXIFITEM
    'Interface 항목 Seq와 일치하는 검사명 뿌리기
        For j = 1 To giOriginIFItemCnt
            If Format$(i, "000") = gIFItem(j).s01 Then
                iCurItemCnt = iCurItemCnt + 1
                Call gfIFDisplayForm.spdIntList.SetText(16 + giTotIFItemCnt + iCurItemCnt, 0, gIFItem(j).s02 & "")
                
                Exit For
            End If
        Next
    Next
    
    For i = 1 To MAXCALITEM
    '계산항목과 일치하는 검사명 뿌리기
        For j = 1 To giOriginCalItemCnt
            If "C" & CStr(i - 1) = gCalItem(j).s01 Then
                iCurItemCnt = iCurItemCnt + 1
                Call gfIFDisplayForm.spdIntList.SetText(16 + giTotIFItemCnt + iCurItemCnt, 0, gCalItem(j).s02 & "")
            
                Exit For
            End If
        Next
    Next
End Sub

Public Sub DisplayRackFormat()
    With gfIFDisplayForm
        If Len(.txtRack) < .txtRack.MaxLength Then
            .txtRack = Format(.txtRack, RackFormat(.txtRack.MaxLength))
        End If
    End With
End Sub

Public Sub DisplayRackPos(ByVal iSRow As Integer)
    On Error GoTo ErrHandler
    
    Dim vRack, vPos
    Dim i%, j%
    Dim sEachRackPos$
    
    With gfIFDisplayForm.spdIntList
        If iSRow = .MaxRows Then Exit Sub
        
        For i = iSRow + 1 To .MaxRows
            Call .GetText(6, i - 1, vRack)
            Call .GetText(7, i - 1, vPos)
            
            For j = 1 To Val(gIFRack.sMaxRack)
                If vRack = "" Then
                    ViewMsgLog "위치 ERR : Rack(Tray) is empty!!"
                    Exit Sub
                Else
                    If gIFPosInfo(j).sRackNo = CStr(vRack) Then
                        sEachRackPos = gIFPosInfo(j).sPosMaxNo
                        Exit For
                    End If
                End If
            Next
            
            If vPos = "" Then
                ViewMsgLog "위치 ERR : Pos(Cup) is empty!!"
                Exit Sub
            End If
            
            If Val(vPos) > Val(sEachRackPos) Then
                ViewMsgLog "위치 ERR : Pos(Cup) is over!!"
                Exit Sub
            End If
            
            'Alphabet
            If gIFRack.sRackDigit = 1 And Val(gIFRack.sMaxRack) = 26 Then
                If Val(vPos) = Val(sEachRackPos) Then
                    'Pos 초기화
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format("1", RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack 증가
                    If IsNumeric(vRack) = False Then
                        Call .SetText(6, i, CVar(Chr(Asc(CStr(vRack)) + 1) & ""))
                    End If
                Else
                    'Pos 증가
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format(Val(vPos) + 1, RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack 그대로
                    If IsNumeric(vRack) = False Then
                        Call .SetText(6, i, CVar(Chr(Asc(CStr(vRack))) & ""))
                    End If
                End If
            '1, 01, 001, 0001, ....
            Else
                If Val(vPos) = Val(sEachRackPos) Then
                    'Pos 초기화
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format("1", RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack 증가
                    If IsNumeric(vRack) = True Then
                        Call .SetText(6, i, CVar(Format(Val(vRack) + 1, RackFormat(gIFRack.sRackDigit)) & ""))
                    End If
                Else
                    'Pos 증가
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format(Val(vPos) + 1, RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack 그대로
                    If IsNumeric(vRack) = True Then
                        Call .SetText(6, i, vRack)
                    End If
                End If
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayRackPos 오류 - (" & Err.Description & ")"
End Sub

Public Sub DisplayOrderOK(Optional ByVal sState$)
    On Error GoTo ErrHandler
    
    Dim i%
    Dim iRowCnt%
    Dim lngTwipHeight&
    
    With gfIFDisplayForm
        If .listTest.ListCount > 10 Then
            .listTest.RemoveItem (0)
        End If
        
        .listTest.AddItem gOrderTable.sJNo
    End With
    
    With gfIFDisplayForm.spdIntList
        '작업일자를 구함
        gOrderTable.sWDate = Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD")
        '작업일련번호를 구함
        gOrderTable.sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
        
        '해당바코드의 오더정보를 넘김
        .MaxRows = .MaxRows + 1
        gOrderTable.iCRow = .MaxRows
        
        Call .RowHeightToTwips(1, .RowHeight(1), lngTwipHeight)
        iRowCnt = Format((.Height / lngTwipHeight) - 2, "0")
        
'        If .MaxRows > iRowCnt Then
'            .TopRow = .MaxRows - iRowCnt + 1
'        End If
        
        Call .SetText(1, gOrderTable.iCRow, gOrderTable.sWSeq & "")
            If sState = "DISPLAY" Then
        Call .SetText(2, gOrderTable.iCRow, CVar("1"))
            Else
        Call .SetText(2, gOrderTable.iCRow, CVar("0"))
            End If
        Call .SetText(3, gOrderTable.iCRow, gOrderTable.sJDate & "")
        Call .SetText(4, gOrderTable.iCRow, gOrderTable.sJGbn & "")
        Call .SetText(5, gOrderTable.iCRow, gOrderTable.sJNo & "")
        Call .SetText(6, gOrderTable.iCRow, gOrderTable.sRack & "")
        Call .SetText(7, gOrderTable.iCRow, gOrderTable.sPos & "")
        Call .SetText(8, gOrderTable.iCRow, gOrderTable.sRegNo & "")
        Call .SetText(9, gOrderTable.iCRow, gOrderTable.sName & "")
        Call .SetText(10, gOrderTable.iCRow, gOrderTable.sSex & "")
        Call .SetText(11, gOrderTable.iCRow, gOrderTable.sEmer & "")
        Call .SetText(12, gOrderTable.iCRow, gOrderTable.sReRun & "")
        Call .SetText(13, gOrderTable.iCRow, gOrderTable.sOther & "")
        Call .SetText(14, gOrderTable.iCRow, CStr(gOrderTable.iOrdCnt) & "")
        Call .SetText(15, gOrderTable.iCRow, "N")
        Call .SetText(16, gOrderTable.iCRow, CStr(gOrderTable.iOrdCnt) & "")
        
        '검사항목 정보 숨기기
        For i = 1 To gOrderTable.iOrdCnt
            Call .SetText(16 + i, gOrderTable.iCRow, gOrderTable.sIFSeq(i) & "||||")
        Next i
        
        If sState <> "DISPLAY" Then
            Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, gOrderTable.iCRow, gOrderTable.iCRow, _
                    RGB(0, 0, 0), 연노랑)
        
            gfIFDisplayForm.lblOrder = gOrderTable.sJNo
        End If
    End With
    
    'Order 내역 Local MDB에 Insert
    Call RegOrder(1)
    
    'gOrderTable 초기화
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
    
    Exit Sub
    
ErrHandler:
    gfIFDisplayForm.listTest.AddItem "Error"
    
    'gOrderTable 초기화
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
End Sub

Public Sub DisplayPosFormat()
    With gfIFDisplayForm
        If Len(.txtPos) < .txtPos.MaxLength Then
            .txtPos = Format(.txtPos, RackFormat(.txtPos.MaxLength))
        End If
    End With
End Sub

Public Sub DisplayResultOK(ByVal iMode As Integer, ByVal sWDate As String, ByVal sWSeq As String, _
                ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                ByVal sRack As String, ByVal sPos As String, ByVal sRegNo As String, ByVal sName As String, _
                ByVal sSex As String, ByVal sEmer As String, ByVal sReRun As String, ByVal sOther As String, _
                ByVal iRstCnt As Integer, ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                ByVal sIFSpcCd As String, ByVal sCurRow As String, Optional ByVal sFlag As String)
    On Error GoTo ErrHandler
    
    Dim sRetVal$, sCWSeq$, sChkVal$
    Dim i%, iCRow%
    Dim vWSeq, vJDate, vJGbn, vJNo, vTmp
    Dim iRowCnt%
    Dim lngTwipHeight&
    
    giAddKey = 0
    
    ReDim gResultTable(1)
    
    With gfIFDisplayForm
        Select Case iMode
            Case 0  'JDate, JGbn, JNo를 넘기는 경우
                .lblResult = sJDate & "-" & sJGbn & "-" & sJNo
                
                iCRow = FindIFListWithJ(sJDate, sJGbn, sJNo)
                
                If iCRow > 0 Then
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag) = "NO" Then
                           
                        Exit Sub
                    End If
                    
                Else
                    If .chkOExist = "1" Then
                    '리스트에 없어도 결과받기를 체크한 경우
                        giAddKey = 1
                    
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '리스트에 없어도 결과받기를 체크하지 않은 경우
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case 1  'WSeq를 넘기는 경우
                .lblResult = sWDate & "-" & sWSeq
                
                iCRow = FindIFListWithW(sWSeq)
                
                If iCRow > 0 Then
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag) = "NO" Then
                           
                        Exit Sub
                    End If
                Else
                    If .chkOExist = "1" Then
                    '리스트에 없어도 결과받기를 체크한 경우
                        giAddKey = 1
                        
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '리스트에 없어도 결과받기를 체크하지 않은 경우
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                                                
            Case 2  'CurRow를 넘기는 경우 - 예) 소변기기같은 단방향 장비
                If .spdIntList.MaxRows >= CInt(sCurRow) Then
                    With .spdIntList
                        Call .GetText(1, CInt(sCurRow), vWSeq)
                        Call .GetText(3, CInt(sCurRow), vJDate)
                        Call .GetText(4, CInt(sCurRow), vJGbn)
                        Call .GetText(5, CInt(sCurRow), vJNo)
                    End With
                    
                    If Len(vJNo) > 0 Then
                        If Trim(vJDate) & Trim(vJGbn) <> "" Then
                            .lblResult = CStr(vJDate) & "-" & CStr(vJGbn) & "-" & CStr(vJNo)
                        Else
                            .lblResult = CStr(vJNo)
                        End If
                    Else
                        .lblResult = Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD") & "-" & CStr(vWSeq)
                    End If
                    
                    iCRow = CInt(sCurRow)
                    
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag) = "NO" Then
                           
                        Exit Sub
                    End If
            
            '리스트 없이 진행하는 완전 단방향의 경우
                Else
                    If .chkOExist = "1" Then
                    '리스트에 없어도 결과받기를 체크한 경우
                        giAddKey = 1
                        
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '리스트에 없어도 결과받기를 체크하지 않은 경우
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
            
            Case 3  'JNo를 넘기는 경우
                .lblResult = sJNo
                
                iCRow = FindIFListWithJNo(sJNo)
                
                If iCRow > 0 Then
                    Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag)
                    
                Else
                    If .chkOExist = "1" Then
                    '리스트에 없어도 결과받기를 체크한 경우
                        '일단 Order를 가져와 뿌린후 결과를 나타냄
                        gOrderTable.sSampID = sJNo
                        
                        If Mid(sJNo, 7, 1) <> "Q" Then
                            Call gfIFDisplayForm.Order_Input("N")
                        End If
                        
                        iCRow = FindIFListWithJNo(sJNo)
                        
                        If iCRow > 0 Then
                            Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag)
                        Else
                            giAddKey = 1
                    
                            sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                        sRack, sPos, sRegNo, sName, _
                                        sSex, sEmer, sReRun, sOther, _
                                        iRstCnt, sIFRstCd, sRst1, sRst2, _
                                        sIFSpcCd, sCurRow, sFlag)
                        End If
                    Else
                    '리스트에 없어도 결과받기를 체크하지 않은 경우
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case 4  '작업리스트와 순서매칭
                With .spdIntList
                    If .MaxRows = 0 Then
                        iCRow = 0
                    Else
                        For i = 1 To .MaxRows
                            Call .GetText(15, i, vTmp)
                            
                            If vTmp = "N" Then
                                iCRow = i
                                Exit For
                            Else
                                iCRow = 0
                            End If
                        Next
                    End If
                End With
                
                If iCRow > 0 Then
                    With .spdIntList
                        Call .GetText(5, iCRow, vJNo)
                    End With
                    
                    .lblResult = CStr(vJNo)
                    
                    Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag)
                
                Else
                    If .chkOExist = "1" Then
                    '리스트에 없어도 결과받기를 체크한 경우
                        giAddKey = 1
                    
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '리스트에 없어도 결과받기를 체크하지 않은 경우
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case Else
            
        End Select
        
        '계산식이 포함된 항목을 조사하여 나타냄
        sChkVal = ChkCalResult(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        
        'Low, High 등을 판정하여 색상을 나타냄
        sRetVal = ViewIFResult2(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        
        With .spdIntList
            Call .RowHeightToTwips(1, .RowHeight(1), lngTwipHeight)
            iRowCnt = Format((.Height / lngTwipHeight) - 2, "0")
            
'            If .MaxRows > iRowCnt Then
'                .TopRow = .MaxRows - iRowCnt + 1
'            End If
        End With
        
    'gsTxMode="0" => Batch, gsTxMode="1" => RealTime(한 항목씩), gsTxMode="2" => RealTime(한 환자씩)
        If gsTXMode = "0" Then
        ElseIf gsTXMode = "1" Then
            If sRetVal = "NONE" Then
            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
                If sChkVal = "1" Then
                    Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sFlag)
                Else
                    Call RegResult(0, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sFlag)
                End If
                
                If giAddKey = 1 Then
                    If sCWSeq = "" Then
                    Else
                        gsLastWSeq = sCWSeq
                    End If
                End If
                
                If sRetVal = "DONE" Then
                    Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
                         gResultTable(1).iCRow, RGB(0, 0, 0), 연초록)
                End If
            End If
        ElseIf gsTXMode = "2" Then
        '원하는 결과등록방식대로 수정 가능함.
            If sRetVal = "NONE" Then
            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
                '환자단위로 결과 등록 시 사용
                If giAddKey = 1 Then
                    If sCWSeq = "" Then
                    Else
                        gsLastWSeq = sCWSeq
                    End If
                End If
                
                Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sFlag)
                
                Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
                                gResultTable(1).iCRow, RGB(0, 0, 0), 연초록)
            End If
        End If
        
        If gfIFDisplayForm.optRegOpt(0).Value = True Then
'            '서버 자동등록用 임시테이블에 등록
'            Call RegResultTemp(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
            'WinSock으로 결과등록화면에 결과내역 전달
            Call SendResultSocket(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        End If
        
        Erase gResultTable
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayResultOK 에러발생 - ( " & Err.Description & " )"
End Sub

Public Sub DisplayResultOkBySex(ByVal iMode As Integer, ByVal sWDate As String, ByVal sWSeq As String, _
                ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                ByVal sRack As String, ByVal sPos As String, ByVal sRegNo As String, ByVal sName As String, _
                ByVal sSex As String, ByVal sEmer As String, ByVal sReRun As String, ByVal sOther As String, _
                ByVal iRstCnt As Integer, ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                ByVal sIFSpcCd As String, ByVal sCurRow As String)
'    On Error GoTo ErrHandler
'
'    Dim sRetVal$, sCWSeq$, sChkVal$
'    Dim i%, iCRow%
'    Dim vWSeq, vJDate, vJGbn, vJNo, vTmp
'    Dim iRowCnt%
'    Dim lngTwipHeight&
'
'    giAddKey = 0
'
'    ReDim gResultTable(1)
'
'    With gfIFDisplayForm
'        Select Case iMode
'            Case 0  'JDate, JGbn, JNo를 넘기는 경우
'                .lblResult = sJDate & "-" & sJGbn & "-" & sJNo
'
'                iCRow = FindIFListWithJ(sJDate, sJGbn, sJNo)
'
'                If iCRow > 0 Then
'                    If OldIFListBySex(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
'                            sRegNo, sName, sSex, sEmer, sReRun, sOther) = "NO" Then
'
'                        Exit Sub
'                    End If
'
'                Else
'                    If .chkOExist = "1" Then
'                    '리스트에 없어도 결과받기를 체크한 경우
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '리스트에 없어도 결과받기를 체크하지 않은 경우
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 1  'WSeq를 넘기는 경우
'                .lblResult = sWDate & "-" & sWSeq
'
'                iCRow = FindIFListWithW(sWSeq)
'
'                If iCRow > 0 Then
'                    If OldIFListBySex(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
'                            sRegNo, sName, sSex, sEmer, sReRun, sOther) = "NO" Then
'
'                        Exit Sub
'                    End If
'                Else
'                    If .chkOExist = "1" Then
'                    '리스트에 없어도 결과받기를 체크한 경우
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '리스트에 없어도 결과받기를 체크하지 않은 경우
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 2  'CurRow를 넘기는 경우 - 예) 소변기기같은 단방향 장비
'                If .spdIntList.MaxRows >= CInt(sCurRow) Then
'                    With .spdIntList
'                        Call .GetText(1, CInt(sCurRow), vWSeq)
'                        Call .GetText(3, CInt(sCurRow), vJDate)
'                        Call .GetText(4, CInt(sCurRow), vJGbn)
'                        Call .GetText(5, CInt(sCurRow), vJNo)
'                    End With
'
'                    If Len(vJNo) > 0 Then
'                        .lblResult = CStr(vJDate) & "-" & CStr(vJGbn) & "-" & CStr(vJNo)
'                    Else
'                        .lblResult = Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD") & "-" & CStr(vWSeq)
'                    End If
'
'                    iCRow = CInt(sCurRow)
'
'                    If OldIFListBySex(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
'                            sRegNo, sName, sSex, sEmer, sReRun, sOther) = "NO" Then
'
'                        Exit Sub
'                    End If
'
'            '리스트 없이 진행하는 완전 단방향의 경우
'                Else
'                    If .chkOExist = "1" Then
'                    '리스트에 없어도 결과받기를 체크한 경우
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '리스트에 없어도 결과받기를 체크하지 않은 경우
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 3  'JNo를 넘기는 경우
'                .lblResult = sJNo
'
'                iCRow = FindIFListWithJNo(sJNo)
'
'                If iCRow > 0 Then
'                    Call OldIFListBySex(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
'                            sRegNo, sName, sSex, sEmer, sReRun, sOther)
'
'                Else
'                    If .chkOExist = "1" Then
'                    '리스트에 없어도 결과받기를 체크한 경우
'                        '일단 Order를 가져와 뿌린후 결과를 나타냄
'                        gOrderTable.sSampID = sJNo
'
'                        Call gfIFDisplayForm.Order_Input("N")
'
'                        iCRow = FindIFListWithJNo(sJNo)
'
'                        If iCRow > 0 Then
'                            Call OldIFListBySex(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
'                            sRegNo, sName, sSex, sEmer, sReRun, sOther)
'                        Else
'                            giAddKey = 1
'
'                            sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                        sRack, sPos, sRegNo, sName, _
'                                        sSex, sEmer, sReRun, sOther, _
'                                        iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                        sIFSpcCd, sCurRow)
'                        End If
'                    Else
'                    '리스트에 없어도 결과받기를 체크하지 않은 경우
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 4  '작업리스트와 순서매칭
'                With .spdIntList
'                    If .MaxRows = 0 Then
'                        iCRow = 0
'                    Else
'                        For i = 1 To .MaxRows
'                            Call .GetText(15, i, vTmp)
'
'                            If vTmp = "N" Then
'                                iCRow = i
'                                Exit For
'                            Else
'                                iCRow = 0
'                            End If
'                        Next
'                    End If
'                End With
'
'                If iCRow > 0 Then
'                    With .spdIntList
'                        Call .GetText(5, iCRow, vJNo)
'                    End With
'
'                    .lblResult = CStr(vJNo)
'
'                    Call OldIFListBySex(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
'                            sRegNo, sName, sSex, sEmer, sReRun, sOther)
'
'                Else
'                    If .chkOExist = "1" Then
'                    '리스트에 없어도 결과받기를 체크한 경우
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '리스트에 없어도 결과받기를 체크하지 않은 경우
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case Else
'
'        End Select
'
'        '계산식이 포함된 항목을 조사하여 나타냄
'        sChkVal = ChkCalResult(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
'
'        'Low, High 등을 판정하여 색상을 나타냄
'        sRetVal = ViewIFResult2(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
'
'        With .spdIntList
'            Call .RowHeightToTwips(1, .RowHeight(1), lngTwipHeight)
'            iRowCnt = Format((.Height / lngTwipHeight) - 2, "0")
'
'            If .MaxRows > iRowCnt Then
'                .TopRow = .MaxRows - iRowCnt + 1
'            End If
'        End With
'
'    'gsTxMode="0" => Batch, gsTxMode="1" => RealTime(한 항목씩), gsTxMode="2" => RealTime(한 환자씩)
'        If gsTXMode = "0" Then
'        ElseIf gsTXMode = "1" Then
'            If sRetVal = "NONE" Then
'            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
'                If sChkVal = "1" Then
'                    Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
'                Else
'                    Call RegResult(0, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
'                End If
'
'                If giAddKey = 1 Then
'                    If sCWSeq = "" Then
'                    Else
'                        gsLastWSeq = sCWSeq
'                    End If
'                End If
'
'                If sRetVal = "DONE" Then
'                    Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
'                         gResultTable(1).iCRow, RGB(0, 0, 0), 연초록)
'                End If
'            End If
'        ElseIf gsTXMode = "2" Then
'        '원하는 결과등록방식대로 수정 가능함.
'            If sRetVal = "NONE" Then
'            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
'                '환자단위로 결과 등록 시 사용
'                If giAddKey = 1 Then
'                    If sCWSeq = "" Then
'                    Else
'                        gsLastWSeq = sCWSeq
'                    End If
'                End If
'
'                Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
'
'                Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
'                         gResultTable(1).iCRow, RGB(0, 0, 0), 연초록)
'            End If
'        End If
'
'        If gfIFDisplayForm.optRegOpt(0).Value = True Then
'            Call RegServerOK(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2)
'        End If
'
'        Erase gResultTable
'    End With
'
'    Exit Sub
'
'ErrHandler:
'    ViewMsg "DisplayResultOkBySex 에러발생 - ( " & Err.Description & " )"
End Sub

Public Sub SetSpdIntLIstColHidden()
    Dim i%
    
    With gfIFDisplayForm.spdIntList
        If gRstcfg.sUse = "1" Then
            For i = 1 To MAXRESULTFIELD - 3
                If i < 4 Then
                    .Col = i + 2
                Else
                    .Col = i + 4
                End If
                
                If gRstcfg.sFUse(i) And Val(gRstcfg.sFSize(i)) > 0 Then
                    .ColHidden = False
                    
                    'HEADER
                    .Row = 0
                    .Text = gRstcfg.sFName(i)
                Else
                    .ColHidden = True
                End If
            Next
            
            gfIFDisplayForm.pnlJNo(0) = gRstcfg.sFName(3)
            gfIFDisplayForm.pnlJNo(1) = gRstcfg.sFName(3)
        Else
            If gOrdCfg.sUse = "1" Then
                For i = 1 To MAXORDERFIELD - 1
                    If i < 4 Then
                        .Col = i + 2
                    Else
                        .Col = i + 4
                    End If
                    
                    If gOrdCfg.sFUse(i) And Val(gOrdCfg.sFSize(i)) > 0 Then
                        .ColHidden = False
                        
                        .Row = 0
                        .Text = gRstcfg.sFName(i)
                    Else
                        .ColHidden = True
                    End If
                Next
            Else
                MsgBox "환경설정에서 Result Setting을 설정하십시요!!", vbCritical
                Exit Sub
            End If
            
            gfIFDisplayForm.pnlJNo(0) = gOrdCfg.sFName(3)
            gfIFDisplayForm.pnlJNo(1) = gOrdCfg.sFName(3)
        End If
        
        If giTotIFItemCnt = 0 Then
            .MaxCols = 16
        Else
            .MaxCols = 16 + 2 * giTotIFItemCnt
            
            For i = 17 To 17 + giTotIFItemCnt - 1
                .Col = i
                .ColHidden = True
            Next
            
            For i = 17 + giTotIFItemCnt To .MaxCols
                .Col = i
                .ColWidth(i) = GetItemColWidth
            Next
        End If
    End With
End Sub

Public Sub SetDefaultRegOption()
    With gfIFDisplayForm
        If gsClientServerMode = "C" Then
            .optRegOpt(1).Value = True
            .optRegOpt(0).Enabled = False
        ElseIf gsClientServerMode = "S" Then
            .optRegOpt(0).Value = True
            .optRegOpt(0).Enabled = True
        Else
            .optRegOpt(1).Value = True
        End If
    End With
End Sub
