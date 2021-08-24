Attribute VB_Name = "modVariable"
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
'1=양방향(OrderMode : No,  QueryMode : No),
'2=양방향(OrderMode : Yes, QueryMode : No),
'3=양방향(OrderMode : No,  QueryMode : Yes),
'4=양방향(OrderMode : Yes, QueryMode : Yes),
Public gsIFMode$
Public gsINITMode$  'Initialize 버튼 사용 모드 - 0=사용안함, 1=사용함
Public gsTXMode$    '결과전송방식 모드 - 0=배치, 1=리얼타임(항목별 전송), 2=리얼타임(환자별 전송)
Public gsAPMode$    '자동출력 모드

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
    sFOrd(MAXORDERFIELD) As String
    sFSize(MAXORDERFIELD) As String
End Type

Type RSTFIELDCFG
    sComponent As String
    sUse As String
    sStorageType As String
    sPath As String
    sFUse(MAXRESULTFIELD) As String
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
            sWDate = Format(Now, "YYYYMMDD")
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

    MsgBox "계산식에 오류가 있습니다.", vbExclamation, "오류"
    If Not IsMissing(nState) Then nState = False
    
    Exit Function
    
End Function

Public Function ChkCalResult(ByVal iCRow As Integer, iRstCnt As Integer, sIFRstCd As String, sRst1 As String, sRst2 As String, sIFSpcCd As String) As String
    Dim i%, j%, k%, iPos%, iSPos%, iCnt%, iExist%
    Dim sCIFRstCd$, sCRst$, sCIFSeq$, sTmp$, sCF$, sCRst2$
    Dim vTmp, vIFCnt, vCRstCnt
    Dim sCompIFSeq As COMPIFSEQ
        
    With gfIFDisplayForm
        With .spdIntList
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
                            
                            If sCRst = "" Then
                            Else
                                iExist = iExist + 1
                            End If
                        End If
                    Next
                Next
                
                '계산을 위한 결과가 모두 전송되었다면
                '계산 결과를 스프레드에 나타냄
                If iCnt = iExist Then
                    For j = 1 To giTotIFItemCnt
                        Call .GetText(16 + j, iCRow, vTmp)
                        
                        sTmp = CStr(vTmp)
                        sCIFSeq = GetByOne(sTmp, sTmp)
                        
                        If sCIFSeq = "" Then
                            sCompIFSeq.iSpdCol = j
                            Exit For
                        End If
                    Next
                    
                    For k = 1 To iCnt
                        sCF = Replace(sCF, "I" & sCompIFSeq.sIFSeq(k), sCompIFSeq.sResult(k))
                    Next
                    
                    sCRst = CFCompute(sCF)
                    
                    If Left(sCRst, 1) = "-" Then
                        sCRst = ConvertResult("-", "0", sCRst, gCalItem(i).s01)
                    Else
                        sCRst = ConvertResult("+", "0", sCRst, gCalItem(i).s01)
                    End If
                    
                    '계산 항목은 IFSEQ("C1"과 같은) 사용하여 IFRstCd를 대신
                    sCRst = JudgeResult(gCalItem(i).s01, sCRst, sRst2)
                    
                    sTmp = sRst2
                    
                    For k = 1 To iRstCnt
                        Call GetByOne(sTmp, sTmp)
                    Next
                    
                    sCRst2 = GetByOne(sTmp, sTmp)
                    
                    Call .GetText(15, iCRow, vCRstCnt)
                    Call .GetText(16, iCRow, vIFCnt)
                    
                    Call .SetText(15, iCRow, CInt(Val(vCRstCnt) + 1) & "")
                    Call .SetText(16, iCRow, CInt(Val(vIFCnt) + 1) & "")
                    
                    Call .SetText(16 + sCompIFSeq.iSpdCol, iCRow, gCalItem(i).s01 & "|" & sCRst & "|" & sCRst2 & "|")
                    
                    '결과등록을 위해 결과등록 파라미터 변환
                    iRstCnt = iRstCnt + 1
                    sIFRstCd = sIFRstCd & gCalItem(i).s01 & "|"
                    sRst1 = sRst1 & sCRst & "|"
                    'sRst2는 JudgeResult에서 받아 옴
                    'sRst2 = sRst2 & "|"
                    
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

Public Function ChkCalResult1(ByVal iCRow As Integer, iRstCnt As Integer, sIFRstCd As String, sRst1 As String, sRst2 As String, sIFSpcCd As String) As String
    Dim i%, j%, k%, iPos%, iSPos%, iCnt%, iExist%, iAlready%
    Dim sCIFRstCd$, sCRst$, sCIFSeq$, sTmp$, sCF$, sCRst2$
    Dim vTmp, vIFCnt, vCRstCnt
    Dim sCompIFSeq As COMPIFSEQ
        
    With gfIFDisplayForm
        With .spdIntList
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
                            
                            If sCRst = "" Then
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
                    sCRst = JudgeResult2(gCalItem(i).s01, sCRst, sCRst2)
                    
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
                    'sCRst2는 JudgeResult2에서 받아 옴
                    sRst2 = sRst2 & sCRst2 & "|"
                    
                    If Len(sIFSpcCd) = 0 Then
                    Else
                        sIFSpcCd = sIFSpcCd & "|"
                    End If
                    
                    ChkCalResult1 = "1"
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

Public Function ConvertIFItemInfoExp(ByVal iMode As Integer, ByVal sComp1 As String, ByVal sComp2 As String) As String
    Dim i%, k%
    Dim aTmp()  As String
    Dim tmpCd1$, tmpCd2$
    Dim tmpChk  As Boolean
    
    Select Case iMode
        '서버쪽코드를 IFSEQ로
        Case 1
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s06 = sComp1 Then
                    tmpChk = False
                
                    Erase aTmp()
                    aTmp() = Split(gIFItem(i).s05 & ",", ",")
                    
                    For k = 0 To UBound(aTmp()) - 1
                        If Trim(aTmp(k)) = "" Then Exit For
                        
                        If Trim(aTmp(k)) = sComp2 Then
                            tmpChk = True
                            Exit For
                        End If
                    Next k
                
                    If tmpChk = True Then
                        ConvertIFItemInfoExp = gIFItem(i).s01
                        Exit For
                    End If
                End If
            Next i
            
            For i = 1 To giOriginCalItemCnt
                If gCalItem(i).s03 = sComp1 Then
                    ConvertIFItemInfoExp = gCalItem(i).s01
                    Exit For
                End If
            Next
            
'        'IFSEQ를 서버쪽코드로
'        Case 2
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s01 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s06
'                    Exit For
'                End If
'            Next
'
'            For i = 1 To giOriginCalItemCnt
'                If gCalItem(i).s01 = sComp Then
'                    ConvertIFItemInfo = gCalItem(i).s03
'                    Exit For
'                End If
'            Next
'
'        '검사항목명을 IFSEQ로
'        Case 3
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s02 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s01
'                    Exit For
'                End If
'            Next
'
'            For i = 1 To giOriginCalItemCnt
'                If gCalItem(i).s02 = sComp Then
'                    ConvertIFItemInfo = gCalItem(i).s01
'                    Exit For
'                End If
'            Next
'
'        'IFSEQ를 검사항목명으로
'        Case 4
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s01 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s02
'                    Exit For
'                End If
'            Next
'
'            For i = 1 To giOriginCalItemCnt
'                If gCalItem(i).s01 = sComp Then
'                    ConvertIFItemInfo = gCalItem(i).s02
'                    Exit For
'                End If
'            Next
'
'        'IFORDCD를 IFSEQ로
'        Case 5
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s03 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s01
'                    Exit For
'                End If
'            Next
'
'        'IFSEQ를 IFORDCD로
'        Case 6
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s01 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s03
'                    Exit For
'                End If
'            Next
'
'        'IFRSTCD를 IFSEQ로
'        Case 7
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s04 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s01
'                    Exit For
'                End If
'            Next
'
'        'IFSEQ를 IFRSTCD로
'        Case 8
'            For i = 1 To giOriginIFItemCnt
'                If gIFItem(i).s01 = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s04
'                    Exit For
'                End If
'            Next
'
'        '서버쪽코드를 IFSEQ로(Sub 코드 제외)
'        Case 9
'            For i = 1 To giOriginIFItemCnt
'                If Trim(Left(gIFItem(i).s06, 8)) = sComp Then
'                    ConvertIFItemInfo = gIFItem(i).s01
'                    Exit For
'                End If
'            Next
'
'            For i = 1 To giOriginCalItemCnt
'                If Trim(Left(gCalItem(i).s03, 8)) = sComp Then
'                    ConvertIFItemInfo = gCalItem(i).s01
'                    Exit For
'                End If
'            Next
             
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
                    
        '서버쪽코드를 IFSEQ로(Sub 코드 제외)
        Case 9
            For i = 1 To giOriginIFItemCnt
                If Trim(Left(gIFItem(i).s06, 8)) = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
            
            For i = 1 To giOriginCalItemCnt
                If Trim(Left(gCalItem(i).s03, 8)) = sComp Then
                    ConvertIFItemInfo = gCalItem(i).s01
                    Exit For
                End If
            Next
            
        'IFSEQ를 IFSPCECIMEN로
        Case 10
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
                                ByVal sIFSpcCd As String, ByVal sCurRow As String) As String
    Dim i%
    Dim vWSeq, vCRstCnt
    Dim sCIFSeq$, sCIFRstCd$, sCRst1$, sCRst2$
    
    NewIFList = ""
    
    With gfIFDisplayForm
        If Len(sWDate) = 0 Then
            sWDate = Format(Now, "YYYYMMDD")
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
                    Call .GetText(15, .MaxRows, vCRstCnt)      'Result
                    Call .SetText(15, .MaxRows, CStr(Val(vCRstCnt) + 1) & "")
                    
                    Call .SetText(16 + Val(vCRstCnt) + 1, .MaxRows, sCIFSeq & "|" & sCRst1 & "|" & sCRst2 & "|")
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

Public Function OldIFList(ByVal iCRow%, ByVal iRstCnt%, _
            ByVal sIFRstCd$, ByVal sRst1$, ByVal sRst2$, ByVal sIFSpcCd$, _
            ByVal sRack$, ByVal sPos$, ByVal sRegNo$, ByVal sName$, _
            ByVal sSex$, ByVal sEmer$, ByVal sReRun$, ByVal sOther$) As String
    
    On Error GoTo ErrHandler
    
    Dim i%, j%, k%, iAdd%, iCCol%, iCompCnt%, iAllCnt%, iExist%
    Dim aIFSeq()    As String
    Dim sCIFRstCd$, sCRst1$, sCRst2$, sCIFSeq$, sTmp$, sPIFSeq$, sPRst1$, sPRst2$, sTIFRstCd$, sTRst1$, sTRst2$
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

Public Function ConvertResult(ByVal 부호 As String, ByVal 지수 As String, ByVal 결과 As String, ByVal IFCD As String) As String
    Dim i%, Pos%
    Dim sDot$, sDotGbn$
    Dim s정수부$, s소수부$, s1$, s2$, s3$, s4$
    
    For i = 1 To giOriginIFItemCnt
        If IFCD = gIFItem(i).s04 Then
            sDot = gIFItem(i).s07
            sDotGbn = gIFItem(i).s08
                        
            Exit For
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If IFCD = gCalItem(i).s01 Then
            sDot = gCalItem(i).s05
            sDotGbn = gCalItem(i).s06
                        
            Exit For
        End If
    Next
    
    If 부호 = "+" Then
        ConvertResult = CStr(결과 * (10 ^ Val(지수)))
    ElseIf 부호 = "-" Then
        ConvertResult = CStr(결과 / (10 ^ Val(지수)))
    End If
    
    If Left(ConvertResult, 1) = "." Then
        ConvertResult = "0" & ConvertResult
    End If
    
'실제 결과값을 소수점 설정에 따라 바꿈
    If sDot = "" Or IsNumeric(sDot) = False Or sDotGbn = "" Or IsNumeric(sDotGbn) = True Then
    Else
    '소수 한자리 더 아래까지 구함
        s소수부 = 소수부구하기(ConvertResult, CInt(sDot) + 1)
        
        If s소수부 = "" Then
            If sDot = "0" Then
                ConvertResult = ConvertResult
            Else
                ConvertResult = ConvertResult & "." & String(CInt(sDot), "0")
            End If
                        
            Exit Function
        End If
        
        s정수부 = 정수부구하기(ConvertResult)
    '소수 한자리 더 아래가 없을 때
        If Mid$(s소수부, CInt(sDot) + 1, 1) = "" Then
            If sDot = "0" Then
                ConvertResult = s정수부
                Exit Function
            Else
                ConvertResult = s정수부 & "." & s소수부
                Exit Function
            End If
    '소수 한자리 더 아래가 있을 때
        Else
        '내림의 경우
            If sDotGbn = "L" Then
                If sDot = "0" Then
                    ConvertResult = s정수부
                    Exit Function
                Else
                    ConvertResult = s정수부 & "." & Mid$(s소수부, 1, CInt(sDot))
                    Exit Function
                End If
        '반올림의 경우
            ElseIf sDotGbn = "H" Then
                If CInt(Mid$(s소수부, CInt(sDot) + 1, 1)) >= 5 Then
                    '올림의 경우와 같음
                    '소수 한 자리 더 아래까지의 숫자
                    s1 = s정수부 & "." & Mid$(s소수부, 1, CInt(sDot) + 1)
                    
                    '(하나 위의 정수) - (소수 한 자리 더 아래까지의 숫자)
                    If s정수부 = "" Then
                        s2 = CStr(CSng(1 - Val(s1)))
                    Else
                        s2 = CStr(CSng(CInt(s정수부) + 1 - Val(s1)))
                    End If
                                        
                    If Left$(s2, 1) = "." Then
                        s2 = "0" & s2
                    End If
                    
                    If s2 = "1" Then
                    '소수점이 모두 0인 경우
                        If sDot = "0" Then
                            ConvertResult = s정수부
                            Exit Function
                        Else
                            ConvertResult = s정수부 & "." & Mid$(s소수부, CInt(sDot))
                            Exit Function
                        End If
                    Else
                    '소수점이 모두 0이 아닌 경우
                        For i = 1 To Len(s2) - 1
                            If IsNumeric(Mid$(s2, i, 1)) = True Then
                                s3 = s3 & "0"
                            Else
                                s3 = s3 & "."
                            End If
                        Next
                        
                        's1 = 17.397 일 때 s3 = 0.003 을 얻기
                        s3 = s3 & CStr(10 - CInt(Mid$(s소수부, CInt(sDot) + 1, 1)))
                        
                        s4 = CStr(CDbl(s1) + CDbl(s3))
                        
                        If sDot = "0" Then
                            ConvertResult = s4
                            Exit Function
                        Else
                            ConvertResult = s4
                            
                            If Len(소수부구하기2(s4)) < CInt(sDot) Then
                                ConvertResult = 정수부구하기(s4) & "." & 소수부구하기2(s4) & _
                                    String(CInt(sDot) - Len(소수부구하기2(s4)), "0")
                            End If
                        End If
                    End If
                Else
                '내림의 경우와 같음
                    If sDot = "0" Then
                        ConvertResult = s정수부
                        Exit Function
                    Else
                        ConvertResult = s정수부 & "." & Mid$(s소수부, 1, CInt(sDot))
                        Exit Function
                    End If
                End If
        '올림의 경우
            ElseIf sDotGbn = "U" Then
            '소수 한 자리 더 아래까지의 숫자
                s1 = s정수부 & "." & Mid$(s소수부, 1, CInt(sDot) + 1)
            
            '(하나 위의 정수) - (소수 한 자리 더 아래까지의 숫자)
                If s정수부 = "" Then
                    s2 = CStr(CSng(1 - Val(s1)))
                Else
                    s2 = CStr(CSng(CInt(s정수부) + 1 - Val(s1)))
                End If
                
                If Left$(s2, 1) = "." Then
                    s2 = "0" & s2
                End If
                
                If s2 = "1" Then
                '소수점이 모두 0인 경우
                    If sDot = "0" Then
                        ConvertResult = s정수부
                        Exit Function
                    Else
                        ConvertResult = s정수부 & "." & Mid$(s소수부, CInt(sDot))
                        Exit Function
                    End If
                Else
                '소수점이 모두 0이 아닌 경우
                    For i = 1 To Len(s2) - 1
                        If IsNumeric(Mid$(s2, i, 1)) = True Then
                            s3 = s3 & "0"
                        Else
                            s3 = s3 & "."
                        End If
                    Next
                    
                    's1 = 17.397 일 때 s3 = 0.003 을 얻기
                    s3 = s3 & CStr(10 - CInt(Mid$(s소수부, CInt(sDot) + 1, 1)))
                    
                    s4 = CStr(CDbl(s1) + CDbl(s3))
                    
                    If sDot = "0" Then
                        ConvertResult = s4
                        Exit Function
                    Else
                        ConvertResult = s4
                        
                        If Len(소수부구하기2(s4)) < CInt(sDot) Then
                            ConvertResult = 정수부구하기(s4) & "." & 소수부구하기2(s4) & _
                                String(CInt(sDot) - Len(소수부구하기2(s4)), "0")
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function ConvertResult1(ByVal sSign As String, ByVal sExp As String, ByVal sRst As String, ByVal sIFRstCd As String) As String
    Dim i%
    Dim sDot$, sDotGbn$
    Dim sValue$, sTmpVal$
    
    For i = 1 To giOriginIFItemCnt
        If sIFRstCd = gIFItem(i).s04 Then
            sDot = gIFItem(i).s07
            sDotGbn = gIFItem(i).s08
                        
            Exit For
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
        sValue = CStr(Val(sRst) * (10 ^ Val(sExp)))
    ElseIf sSign = "-" Then
        sValue = CStr(Val(sRst) / (10 ^ Val(sExp)))
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

Public Sub DisplayResult(ByVal iRow As Integer)
    On Error GoTo ErrHandler
    
    Dim vIFCnt, vTmp, vJDate, vJGbn, vJNo
    Dim i%, j%, k%
    Dim sTmp$, sTestCd$, sCRst1$, sTestNm$, sCRst2$
    
    Call ResultSpdClear
    
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
            
            sTestCd = GetByOne(sTmp, sTmp)
            sCRst1 = GetByOne(sTmp, sTmp)
            sCRst2 = GetByOne(sTmp, sTmp)
            
            sTestNm = ConvertIFItemInfo(4, sTestCd)
            
            If i <= 15 Then
                
                Call gfIFDisplayForm.spdRst.SetText(1, i, sTestNm & "")
                Call gfIFDisplayForm.spdRst.SetText(2, i, sCRst1 & "")
                Call gfIFDisplayForm.spdRst.SetText(3, i, sCRst2 & "")
            Else

                Call gfIFDisplayForm.spdRst2.SetText(1, i - 15, sTestNm & "")
                Call gfIFDisplayForm.spdRst2.SetText(2, i - 15, sCRst1 & "")
                Call gfIFDisplayForm.spdRst2.SetText(3, i, sCRst2 & "")
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayResult 에러발생" & "(" & CStr(Err.Description) & ")"
End Sub

Public Sub DisplayResult1(ByVal iRow As Integer)
    On Error GoTo ErrHandler
    
    Dim vIFCnt, vTmp, vJDate, vJGbn, vJNo
    Dim i%, j%, k%
    Dim sTmp$, sTestCd$, sCRst1$, sTestNm$, sCRst2$
    
    Call ResultSpdClear1
    
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
            
            sTestCd = GetByOne(sTmp, sTmp)
            sCRst1 = GetByOne(sTmp, sTmp)
            sCRst2 = GetByOne(sTmp, sTmp)
            
            sTestNm = ConvertIFItemInfo(4, sTestCd)
                            
            Call gfIFDisplayForm.spdRst.SetText(1, i, sTestNm & "")
            Call gfIFDisplayForm.spdRst.SetText(2, i, sCRst1 & "")
            Call gfIFDisplayForm.spdRst.SetText(3, i, sCRst2 & "")
            Call gfIFDisplayForm.spdRst.SetText(4, i, sCRst2 & "")
            
            If sCRst2 = "High" Or sCRst2 = "Positive" Then
                Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(255, 220, 220))
            ElseIf sCRst2 = "Low" Then
                Call SpdForeBack(gfIFDisplayForm.spdRst, 1, 4, i, i, RGB(0, 0, 0), RGB(220, 220, 255))
            End If

        Next
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayResult1 에러발생" & "(" & CStr(Err.Description) & ")"
End Sub

Public Sub EditIFList(ByVal iCRow As Integer, ByVal iRstCnt As Integer, _
    ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, ByVal sIFSpcCd As String)
    
    On Error GoTo ErrHandler
    
    Dim i%, j%, k%, iAdd%, iCCol%
    Dim sCIFRstCd$, sCRst1$, sCRst2$, sCIFSeq$, sTmp$, sPIFSeq$, sPRst1$, sPRst2$, sTIFRstCd$, sTRst1$, sTRst2$
    Dim vTmp, vIFCnt, vCRstCnt, vRack, vPos, vTTestInfo
    
    With gfIFDisplayForm
        With .spdIntList
            For i = 1 To iRstCnt
                iAdd = 0
                                
                Call .GetText(16, iCRow, vIFCnt)
                Call .GetText(6, iCRow, vRack)
                Call .GetText(7, iCRow, vPos)
                Call .GetText(15, iCRow, vCRstCnt)      'Result
                
                If vRack = "ADD" And vPos = "NO" Then
                    iAdd = 1
                End If
                
                ReDim gResultTable(1)
                          
                If iAdd = 0 Then
                    For j = 1 To CInt(Val(vIFCnt))
                        Call .GetText(16 + j, iCRow, vTmp)
                        
                        sTmp = CStr(vTmp)
                        sCIFSeq = GetByOne(sTmp, sTmp)
                        gResultTable(1).sTestCd(j) = sCIFSeq
                    Next
                ElseIf iAdd = 1 Then
                    For j = 1 To .MaxCols
                        sTIFRstCd = sIFRstCd
                        sTRst1 = sRst1
                        sTRst2 = sRst2
                        Call .GetText(16 + j, iCRow, vTmp)
                        
                        sTmp = CStr(vTmp)
                        sCIFSeq = GetByOne(sTmp, sTmp)
                        
                        '전송받은 항목이 재전송되면 iAdd = 0
                        If sCIFSeq = "" Then
                            sTmp = GetByOne(sTmp, sTmp)
                            
                            If sTmp = "" Then
                                iCCol = j
                            Else
                                iAdd = 0
                            End If
                            
                            Exit For
                        Else
                            For k = 1 To iRstCnt
                                sCIFRstCd = GetByOne(sTIFRstCd, sTIFRstCd)
                                sCRst1 = GetByOne(sTRst1, sTRst1)
                                sCRst2 = GetByOne(sTRst2, sTRst2)
                                
                                '이전 검사항목과 현재 전송받은 항목이 같으면
                                'For 문 탈출
                                If ConvertIFItemInfo(8, sCIFSeq) = sCIFRstCd Then
                                    iAdd = 0
                                    Exit For
                                End If
                            Next
                        End If
                        
                        If iAdd = 0 Then
                        '이전 검사항목과 현재 전송받은 항목이 같으면
                        '현재까지 전송받은 항목의 정보를 넘기고
                        '전체 For 문 탈출
                            For k = 1 To CInt(Val(vIFCnt))
                                Call .GetText(16 + j, iCRow, vTmp)
                                
                                sTmp = CStr(vTmp)
                                sCIFSeq = GetByOne(sTmp, sTmp)
                                gResultTable(1).sTestCd(j) = sCIFSeq
                            Next
                            
                            Exit For
                        End If
                    Next
                End If
                
                'iAdd = 0 (Edit), iAdd =1 (Add)에 따라 화면에 나타냄
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                
                If iAdd = 0 Then
                    For j = 1 To CInt(Val(vIFCnt))
                        If Len(sCIFRstCd) = 0 Then
                        Else
                            If ConvertIFItemInfo(8, gResultTable(1).sTestCd(j)) = sCIFRstCd Then
                                Call .GetText(16 + j, iCRow, vTTestInfo)
                                
                                sTmp = CStr(vTTestInfo)
                                sPIFSeq = GetByOne(sTmp, sTmp)
                                sPRst1 = GetByOne(sTmp, sTmp)
                                sPRst2 = GetByOne(sTmp, sTmp)
                                
                                If vCRstCnt = "N" Then
                                    Call .SetText(15, iCRow, "1")
                                Else
                                    '전송받은 항목이 재전송되면 결과 카운트 그대로
                                    If sPRst1 = "" Then
                                        Call .SetText(15, iCRow, CStr(Val(vCRstCnt) + 1) & "")
                                    Else
                                    End If
                                End If
                                
                                Call .SetText(16 + j, iCRow, gResultTable(1).sTestCd(j) & "|" & sCRst1 & "|" & sCRst2 & "|")
                                Exit For
                            End If
                        End If
                    Next
                'List에 없는 경우 추가시
                ElseIf iAdd = 1 Then
                    Call .GetText(16, iCRow, vIFCnt)
                    
                    sTmp = ConvertIFItemInfo(7, sCIFRstCd)
                    
                    If sTmp = "" Then
                    Else
                        Call .SetText(15, iCRow, CStr(Val(vCRstCnt) + 1) & "")
                        Call .SetText(16, iCRow, CStr(Val(vIFCnt) + 1) & "")
                        
                        Call .SetText(16 + iCCol + i - 1, iCRow, sTmp & "|" & sCRst1 & "|" & sCRst2 & "|")
                    End If
                End If
            Next
        End With
        
        '현재 Row 기록
        gResultTable(1).iCRow = iCRow
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "EditIFList - Err(" & Err.Description & ")"
End Sub



Public Sub EditRegState(ByVal iPersonCnt As Integer, ByVal sWDate As String, ByVal sTWSeq As String)
    On Error GoTo ErrHandler
    
    Dim objLD As Object
    
    Set objLD = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    Call objLD.Edit_IFResult(gsMachineCd, 2, sWDate, sTWSeq, "", "", "", "", _
                            "", "", "", "", "", "", "", "", "", "", iPersonCnt)
                            
    Set objLD = Nothing
      
    Exit Sub
ErrHandler:
    Set objLD = Nothing
    MsgBox "EditRegState - Err(" & Err.Description & ")"
End Sub
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

Public Sub GetIFComment()
    Dim sBuf$
    Dim iComCnt%
    Dim i%
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "IFComment")
        
    If sBuf = "" Then
    Else
        iComCnt = CInt(GetByOne(sBuf, sBuf))
        
        ReDim gCommentCd(iComCnt)
        
        For i = 1 To iComCnt
            gCommentCd(i) = GetByOne(sBuf, sBuf)
        Next
    End If
    
End Sub

Public Function GetIFStateFlag(ByVal sGbn As String) As Integer
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Flag." & sGbn)
        
    GetIFStateFlag = CInt(Val(sBuf))
End Function

Public Function GetLastWorkSeq(ByVal sWDate As String) As String
    Dim objDB As Object
    Dim sRtnVal$
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    sRtnVal = objDB.Get_LastIFResult(gsMachineCd, sWDate)
    
    gsLastWSeq = Format$(GetByOneUserSymbol(sRtnVal, sRtnVal, Chr$(3)), "0000")
    
    Set objDB = Nothing
    
End Function

Public Function GetCurLastWSeq() As String
    GetCurLastWSeq = ""
    
    With gfIFDisplayForm.spdIntList
        Call GetLastWorkSeq(Format(Now, "YYYYMMDD"))
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
    
    For i = 1 To CInt(gIFRack.sMaxRack)
        gIFPosInfo(i).sRackNo = Format$(i, RackFormat(gIFRack.sRackDigit))
        gIFPosInfo(i).sPosMaxNo = GetByOne(gIFRack.sPosSetting, gIFRack.sPosSetting)
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

Public Function GetUserId()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "User.Id")
        
    GetUserId = sBuf
    
    Exit Function
    
ErrHandler:
    ViewMsg "GetUserId - Err(" & Err.Description & ")"
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

Public Function JudgeResult(ByVal sIFRstCd As String, ByVal sCompRst As String, sTRst2 As String) As String
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sJGbn$, sRef1$, sRef2$
    
    For i = 1 To giOriginIFItemCnt
        If sIFRstCd = gIFItem(i).s04 Then
            sJGbn = gIFItem(i).s09
            sRef1 = gIFItem(i).s10
            sRef2 = gIFItem(i).s11
            
            Exit For
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If sIFRstCd = gCalItem(i).s01 Then
            sJGbn = gCalItem(i).s07
            sRef1 = gCalItem(i).s08
            sRef2 = gCalItem(i).s09
            
            Exit For
        End If
    Next
    
    If IsNumeric(sCompRst) = False Then
        If (sCompRst = "LOWER LIMIT" Or sCompRst = "UPPER LIMIT") And sJGbn = "4" Then
        Else
            JudgeResult = sCompRst
            sTRst2 = sTRst2 & "|"
            Exit Function
        End If
    End If
        
    Select Case sJGbn
        Case "0"
            JudgeResult = sCompRst
            sTRst2 = sTRst2 & "|"
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "Low|"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                sTRst2 = sTRst2 & "|"
            Else
                sTRst2 = sTRst2 & "High|"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResult = "Negative"
                sTRst2 = sTRst2 & "|"
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResult = "Positive"
                sTRst2 = sTRst2 & "|"
            Else
                JudgeResult = "GrayZone(+/-)"
                sTRst2 = sTRst2 & "|"
            End If
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "Negative|"
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "Positive|"
            Else
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "GrayZone(+/-)|"
            End If
        Case "4"
        '이하 / 이상
            If IsNumeric(sCompRst) = True Then
                If Val(sCompRst) <= Val(sRef1) Then
                    JudgeResult = sRef1
                    sTRst2 = sTRst2 & "이하|"
                ElseIf Val(sCompRst) > Val(sRef1) And Val(sCompRst) < Val(sRef2) Then
                    JudgeResult = sCompRst
                    sTRst2 = sTRst2 & "|"
                Else
                    JudgeResult = sRef2
                    sTRst2 = sTRst2 & "이상|"
                End If
            Else
                If sCompRst = "LOWER LIMIT" Then
                    If sRef1 = "" Then
                    Else
                        JudgeResult = sRef1
                        sTRst2 = sTRst2 & "이하|"
                    End If
                ElseIf sCompRst = "UPPER LIMIT" Then
                    If sRef2 = "" Then
                    Else
                        JudgeResult = sRef2
                        sTRst2 = sTRst2 & "이상|"
                    End If
                End If
            End If
        Case "5"
        'QAL P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult = "Positive"
                sTRst2 = sTRst2 & "|"
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResult = "Negative"
                sTRst2 = sTRst2 & "|"
            Else
                JudgeResult = "GrayZone(+/-)"
                sTRst2 = sTRst2 & "|"
            End If
        Case "6"
        'QAN P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "Positive|"
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "Negative|"
            Else
                JudgeResult = sCompRst
                sTRst2 = sTRst2 & "GrayZone(+/-)|"
            End If
        Case "7"
        'P/N 장비
        
        Case Else
        
    End Select
    
    Exit Function
    
ErrHandler:
    MsgBox "JudgeResult - Err(" & Err.Description & ")"
End Function

Public Function JudgeResult1(ByVal sIFRstCd As String, ByVal sCompRst As String, sOneRst2 As String) As String
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sJGbn$, sRef1$, sRef2$
    
    For i = 1 To giOriginIFItemCnt
        If sIFRstCd = gIFItem(i).s04 Then
            sJGbn = gIFItem(i).s09
            sRef1 = gIFItem(i).s10
            sRef2 = gIFItem(i).s11
            
            Exit For
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If sIFRstCd = gCalItem(i).s01 Then
            sJGbn = gCalItem(i).s07
            sRef1 = gCalItem(i).s08
            sRef2 = gCalItem(i).s09
            
            Exit For
        End If
    Next
    
    If IsNumeric(sCompRst) = False Then
        If (sCompRst = "LOWER LIMIT" Or sCompRst = "UPPER LIMIT") And sJGbn = "4" Then
        Else
            JudgeResult1 = sCompRst
            sOneRst2 = Chr$(124)
            Exit Function
        End If
    End If
        
    Select Case sJGbn
        Case "0"
            JudgeResult1 = sCompRst
            sOneRst2 = ""
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult1 = sCompRst
                sOneRst2 = "Low"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                sOneRst2 = ""
            Else
                sOneRst2 = "High"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResult1 = "Negative"
                sOneRst2 = ""
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResult1 = "Positive"
                sOneRst2 = ""
            Else
                JudgeResult1 = "GrayZone(+/-)"
                sOneRst2 = ""
            End If
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResult1 = sCompRst
                sOneRst2 = "Negative"
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResult1 = sCompRst
                sOneRst2 = "Positive"
            Else
                JudgeResult1 = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "4"
        '이하 / 이상
            If IsNumeric(sCompRst) = True Then
                If Val(sCompRst) <= Val(sRef1) Then
                    JudgeResult1 = "<" & sRef1
                    sOneRst2 = "이하"
                ElseIf Val(sCompRst) > Val(sRef1) And Val(sCompRst) < Val(sRef2) Then
                    JudgeResult1 = sCompRst
                    sOneRst2 = ""
                Else
                    JudgeResult1 = ">" & sRef2
                    sOneRst2 = "이상"
                End If
            Else
                If sCompRst = "LOWER LIMIT" Then
                    If sRef1 = "" Then
                    Else
                        JudgeResult1 = "<" & sRef1
                        sOneRst2 = "이하"
                    End If
                ElseIf sCompRst = "UPPER LIMIT" Then
                    If sRef2 = "" Then
                    Else
                        JudgeResult1 = ">" & sRef2
                        sOneRst2 = "이상"
                    End If
                End If
            End If
        Case "5"
        'QAL P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult1 = "Positive"
                sOneRst2 = ""
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResult1 = "Negative"
                sOneRst2 = ""
            Else
                JudgeResult1 = "GrayZone(+/-)"
                sOneRst2 = ""
            End If
        Case "6"
        'QAN P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult1 = sCompRst
                sOneRst2 = "Positive"
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResult1 = sCompRst
                sOneRst2 = "Negative"
            Else
                JudgeResult1 = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "7"
        'P/N 장비
        
        Case Else
        
    End Select
    
    Exit Function
    
ErrHandler:
    ViewMsg "JudgeResult1 - Err(" & Err.Description & ")"
End Function

Public Function JudgeResult2(ByVal sIFSeq As String, ByVal sCompRst As String, sOneRst2 As String) As String
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sJGbn$, sRef1$, sRef2$
    
    For i = 1 To giOriginIFItemCnt
        If sIFSeq = gIFItem(i).s01 Then
            sJGbn = gIFItem(i).s09
            sRef1 = gIFItem(i).s10
            sRef2 = gIFItem(i).s11
            
            Exit For
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If sIFSeq = gCalItem(i).s01 Then
            sJGbn = gCalItem(i).s07
            sRef1 = gCalItem(i).s08
            sRef2 = gCalItem(i).s09
            
            Exit For
        End If
    Next
    
    If IsNumeric(sCompRst) = False Then
        If (sCompRst = "LOWER LIMIT" Or sCompRst = "UPPER LIMIT") And sJGbn = "4" Then
        Else
            JudgeResult2 = sCompRst
            sOneRst2 = ""
            Exit Function
        End If
    End If
        
    Select Case sJGbn
        Case "0"
            JudgeResult2 = sCompRst
            sOneRst2 = ""
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult2 = sCompRst
                sOneRst2 = "Low"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                JudgeResult2 = sCompRst
                sOneRst2 = ""
            Else
                JudgeResult2 = sCompRst
                sOneRst2 = "High"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResult2 = "Negative"
                sOneRst2 = ""
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResult2 = "Positive"
                sOneRst2 = ""
            Else
                JudgeResult2 = "GrayZone(+/-)"
                sOneRst2 = ""
            End If
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResult2 = sCompRst
                sOneRst2 = "Negative"
            ElseIf Val(sCompRst) > Val(sRef1) + Val(sRef2) Then
                JudgeResult2 = sCompRst
                sOneRst2 = "Positive"
            Else
                JudgeResult2 = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "4"
        '이하 / 이상
            If IsNumeric(sCompRst) = True Then
                If Val(sCompRst) <= Val(sRef1) Then
                    JudgeResult2 = sRef1
                    sOneRst2 = "이하"
                ElseIf Val(sCompRst) > Val(sRef1) And Val(sCompRst) < Val(sRef2) Then
                    JudgeResult2 = sCompRst
                    sOneRst2 = ""
                Else
                    JudgeResult2 = sRef2
                    sOneRst2 = "이상"
                End If
            Else
                If sCompRst = "LOWER LIMIT" Then
                    If sRef1 = "" Then
                    Else
                        JudgeResult2 = sRef1
                        sOneRst2 = "이하"
                    End If
                ElseIf sCompRst = "UPPER LIMIT" Then
                    If sRef2 = "" Then
                    Else
                        JudgeResult2 = sRef2
                        sOneRst2 = "이상"
                    End If
                End If
            End If
        Case "5"
        'QAL P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult2 = "Positive"
                sOneRst2 = ""
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResult2 = "Negative"
                sOneRst2 = ""
            Else
                JudgeResult2 = "GrayZone(+/-)"
                sOneRst2 = ""
            End If
        Case "6"
        'QAN P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResult2 = sCompRst
                sOneRst2 = "Positive"
            ElseIf Val(sCompRst) >= Val(sRef1) + Val(sRef2) Then
                JudgeResult2 = sCompRst
                sOneRst2 = "Negative"
            Else
                JudgeResult2 = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "7"
        'P/N 장비
        
        Case Else
        
    End Select
    
    Exit Function
    
ErrHandler:
    ViewMsg "JudgeResult2 - Err(" & Err.Description & ")"
End Function

Public Function JudgeResultNewByIFSeq(ByVal sIFSeq As String, ByVal sCompRst As String, _
                                   sOneRst2 As String, Optional sConvRst As String) As String
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sJGbn$, sRef1$, sRef2$, sLimit1Gbn$, sLimit2Gbn$, sLimit1$, sLimit2$
    
    For i = 1 To giOriginIFItemCnt
        If sIFSeq = gIFItem(i).s01 Then
            sJGbn = gIFItem(i).s09
            sRef1 = gIFItem(i).s10
            sRef2 = gIFItem(i).s11
            sLimit1Gbn = gIFItem(i).s12
            sLimit1 = gIFItem(i).s13
            sLimit2Gbn = gIFItem(i).s14
            sLimit2 = gIFItem(i).s15
            
            Exit For
        End If
    Next
    
    For i = 1 To giOriginCalItemCnt
        If sIFSeq = gCalItem(i).s01 Then
            sJGbn = gCalItem(i).s07
            sRef1 = gCalItem(i).s08
            sRef2 = gCalItem(i).s09
            
            Exit For
        End If
    Next
    
    sOneRst2 = ""
    sConvRst = ""
    
    If IsNumeric(sCompRst) = False Then
        If (sCompRst = "LOWER LIMIT" Or sCompRst = "UPPER LIMIT") Then
            JudgeResultNewByIFSeq = sCompRst
            
            If sCompRst = "LOWER LIMIT" Then
                If sLimit1 <> "" Then
                    Select Case sLimit1Gbn
                        Case "0"
                            sConvRst = sLimit1
                            JudgeResultNewByIFSeq = sCompRst
                        Case "1"
                            sConvRst = "< " & sLimit1
                            JudgeResultNewByIFSeq = sCompRst
                        Case "2"
                            sConvRst = sLimit1 & " 이하"
                            JudgeResultNewByIFSeq = sCompRst
                        Case Else
                    End Select
                End If
            ElseIf sCompRst = "UPPER LIMIT" Then
                If sLimit2 <> "" Then
                    Select Case sLimit2Gbn
                        Case "0"
                            sConvRst = sLimit2
                            JudgeResultNewByIFSeq = sCompRst
                        Case "1"
                            sConvRst = "> " & sLimit2
                            JudgeResultNewByIFSeq = sCompRst
                        Case "2"
                            sConvRst = sLimit2 & " 이상"
                            JudgeResultNewByIFSeq = sCompRst
                        Case Else
                    End Select
                End If
            End If
            
            Exit Function
        Else
            JudgeResultNewByIFSeq = sCompRst
            
            Exit Function
        End If
    End If
    
    sOneRst2 = ""
    sConvRst = ""
            
    Select Case sJGbn
        Case "0"
            JudgeResultNewByIFSeq = sCompRst
            sOneRst2 = ""
            sConvRst = ""
        Case "1"
        'L/H
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "Low"
            ElseIf Val(sRef1) <= Val(sCompRst) And Val(sCompRst) <= Val(sRef2) Then
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = ""
            Else
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "High"
            End If
        Case "2"
        'QAL N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultNewByIFSeq = "Negative"
                sOneRst2 = "Negative"
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeResultNewByIFSeq = "Positive"
                sOneRst2 = "Positive"
            Else
                JudgeResultNewByIFSeq = "GrayZone(+/-)"
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "3"
        'QAN N/P
            If Val(sCompRst) <= Val(sRef1) Then
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "Negative"
            ElseIf Val(sCompRst) > Val(sRef2) Then
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "Positive"
            Else
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "4"
        '이하 / 이상
            If IsNumeric(sCompRst) = True Then
                If Val(sCompRst) <= Val(sRef1) Then
                    JudgeResultNewByIFSeq = "<" & sRef1
                    sOneRst2 = "이하"
                ElseIf Val(sCompRst) > Val(sRef1) And Val(sCompRst) < Val(sRef2) Then
                    JudgeResultNewByIFSeq = sCompRst
                    sOneRst2 = ""
                Else
                    JudgeResultNewByIFSeq = ">" & sRef2
                    sOneRst2 = "이상"
                End If
            Else
                If sCompRst = "LOWER LIMIT" Then
                    If sRef1 = "" Then
                    Else
                        JudgeResultNewByIFSeq = "<" & sRef1
                        sOneRst2 = "이하"
                    End If
                ElseIf sCompRst = "UPPER LIMIT" Then
                    If sRef2 = "" Then
                    Else
                        JudgeResultNewByIFSeq = ">" & sRef2
                        sOneRst2 = "이상"
                    End If
                End If
            End If
        Case "5"
        'QAL P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultNewByIFSeq = "Positive"
                sOneRst2 = "Positive"
            ElseIf Val(sCompRst) >= Val(sRef2) Then
                JudgeResultNewByIFSeq = "Negative"
                sOneRst2 = "Negative"
            Else
                JudgeResultNewByIFSeq = "GrayZone(+/-)"
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "6"
        'QAN P/N
            If Val(sCompRst) < Val(sRef1) Then
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "Positive"
            ElseIf Val(sCompRst) >= Val(sRef2) Then
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "Negative"
            Else
                JudgeResultNewByIFSeq = sCompRst
                sOneRst2 = "GrayZone(+/-)"
            End If
        Case "7"
        'P/N 장비
        
        Case Else
        
    End Select
    
    'LIMIT구분에 따른 처리
    If sLimit1 <> "" And Val(sCompRst) <= Val(sLimit1) Then
        Select Case sLimit1Gbn
            Case "0"
                sConvRst = sLimit1
            Case "1"
                sConvRst = "< " & sLimit1
            Case "2"
                sConvRst = sLimit1 & " 이하"
        End Select
    End If
    
    If sLimit2 <> "" And Val(sCompRst) >= Val(sLimit2) Then
        Select Case sLimit2Gbn
            Case "0"
                sConvRst = sLimit2
            Case "1"
                sConvRst = "> " & sLimit2
            Case "2"
                sConvRst = sLimit2 & " 이상"
        End Select
    End If
    
    Exit Function
    
ErrHandler:
    ViewMsg "JudgeResultNewByIFSeq - Err(" & Err.Description & ")"
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
        sOneRow = sDataRow(i) & "|"
        
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
    Next
End Sub

Public Sub MakeIFOrder()
    Dim sBuf$, sSampInfo$, sItemInfo$, sTmp$, sOneRow$
    Dim i%, j%
    
'''    Type ORDTBL
'''        iCRow As Integer
'''        sSampID As String
'''        sIFSpcCd As String
'''        sOrdOpt As String
'''        iOrdCnt As Integer
'''        sIFOrdCd() As String
'''        sServerCd() As String
'''        sIFRstCd() As String
'''        'IFRESULT
'''        sWDate As String
'''        sWSeq As String
'''        sJDate As String
'''        sJGbn As String
'''        sJNo As String
'''        sRack As String
'''        sPos As String
'''        sRegNo As String
'''        sName As String
'''        sSex As String
'''        sEmer As String
'''        sReRun As String
'''        sOther As String
'''    End Type
    
    With gOrderTable
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.SampInfo")
        sSampInfo = sBuf
        
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.ItemInfo")
        sItemInfo = sBuf
        
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.ItemCnt")
        .iOrdCnt = CInt(Val(sBuf))
        
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.OrdOpt")
        .sOrdOpt = sBuf
        
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.WDate")
        .sWDate = sBuf
        
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.WSeq")
        .sWSeq = sBuf
        
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.Specimen")
        .sIFSpcCd = sBuf
        
        For i = 1 To MAXORDERFIELD + 2
            sTmp = GetByOne(sSampInfo, sSampInfo)
            
            If i = 1 Then
            '접수일자
                .sJDate = sTmp
            ElseIf i = 2 Then
            '접수구분
                .sJGbn = sTmp
            ElseIf i = 3 Then
            '접수번호
                .sJNo = sTmp
            ElseIf i = 4 Then
            'Rack
                .sRack = sTmp
            ElseIf i = 5 Then
            'Pos
                .sPos = sTmp
            ElseIf i = 6 Then
            'RegNo
                .sRegNo = sTmp
            ElseIf i = 7 Then
            'Name
                .sName = sTmp
            ElseIf i = 8 Then
            'Sex
                .sSex = sTmp
            ElseIf i = 9 Then
            '응급
                .sEmer = sTmp
            ElseIf i = 10 Then
            '재검
                .sReRun = sTmp
            ElseIf i = 11 Then
            '기타
                .sOther = sTmp
            End If
        Next
        
        ReDim .sIFOrdCd(CInt(Val(.iOrdCnt)))
        ReDim .sIFRstCd(CInt(Val(.iOrdCnt)))
        ReDim .sServerCd(CInt(Val(.iOrdCnt)))
                
        For i = 1 To Val(.iOrdCnt)
            sOneRow = GetByOneUserSymbol(sItemInfo, sItemInfo, Chr(3))
            
            'IFTEST 설정 관련 수 = IFTESTFIELD
            For j = 1 To IFTESTFIELD
                sTmp = GetByOne(sOneRow, sOneRow)
            
                If j = 1 Then
                'IFTESTSEQ
                ElseIf j = 2 Then
                'IFTESTNM
                ElseIf j = 3 Then
                'IFORDCD
                    .sIFOrdCd(i) = sTmp
                ElseIf j = 4 Then
                'IFRSTCD
                    .sIFRstCd(i) = sTmp
                ElseIf j = 5 Then
                'IFSPCCD
                    '위에서 구했으므로 생략
                    'gOrderTable.sIFSpcCd = sTmp
                ElseIf j = 6 Then
                'IFSVRCD
                    .sServerCd(i) = sTmp
                    
                    Exit For
                ElseIf j = 7 Then
                'DOTDIGIT
                ElseIf j = 8 Then
                'LHU
                ElseIf j = 9 Then
                'JUDGEGBN
                ElseIf j = 10 Then
                'REF1
                ElseIf j = 11 Then
                'REF2
                
                End If
            Next
        Next
    End With
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
        sOneRow = sDataRow(i) & "|"
        
        gCalItem(i).s01 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s02 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s03 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s04 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s05 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s06 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s07 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s08 = GetByOne(sOneRow, sOneRow)
        gCalItem(i).s09 = GetByOne(sOneRow, sOneRow)
'''        gcalItem(i).s10 = GetByOne(sOneRow, sOneRow)
'''        gcalItem(i).s11 = GetByOne(sOneRow, sOneRow)
'''        gcalItem(i).s12 = GetByOne(sOneRow, sOneRow)
'''        gcalItem(i).s13 = GetByOne(sOneRow, sOneRow)
'''        gcalItem(i).s14 = GetByOne(sOneRow, sOneRow)
'''        gcalItem(i).s15 = GetByOne(sOneRow, sOneRow)
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

Public Sub RegEditOrdInfo(ByVal i1stRow As Integer, ByVal sSampInfo As String, ByVal iItemCnt As Integer, ByVal sItemInfo As String, _
                ByVal sWDate As String, ByVal sWSeq As String, ByVal sSpecimen As String, ByVal sOrdOpt As String)
    Dim bRetVal As Boolean
        
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.CurRow", CStr(i1stRow))
                
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.SampInfo", sSampInfo)
                
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.ItemCnt", CStr(iItemCnt))
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                       "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.ItemInfo", sItemInfo)
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.WDate", sWDate)
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.WSeq", sWSeq)
            
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.Specimen", sSpecimen)
                        
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\Ack_if\Interface Config\" & gsMachineCd, "Ord.OrdOpt", sOrdOpt)
            
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

Public Sub RegIFStateFlag(ByVal sGbn As String, ByVal sVal As String)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Flag." & sGbn, sVal)
End Sub

Public Sub RegOrder(ByVal iMode As Integer)
    On Error GoTo ErrHandler
'    Public Function Add_IFResult(ByVal sMachineCd As String, ByVal iMode As Integer, _
'            ByVal sWDate As String, ByVal sWSeq As String, ByVal sIFSeq As String, _
'            ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
'            ByVal sRack As String, ByVal sPos As String, _
'            ByVal sRegNo As String, ByVal sName As String, ByVal sSex As String, _
'            ByVal sEmer As String, ByVal sReRun As String, ByVal sOther As String, _
'            ByVal sRst1 As String, ByVal sRst2 As String, ByVal sRegState As String)
    Dim objLD As Object
    Dim i%
    Dim sTIFSeq$
    
    sTIFSeq = ""
    
    For i = 1 To gOrderTable.iOrdCnt
        sTIFSeq = sTIFSeq & gOrderTable.sIFSeq(i) & "|"
    Next
    
    Set objLD = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    Call objLD.Add_IFResult(gsMachineCd, iMode, gOrderTable.sWDate, gOrderTable.sWSeq, sTIFSeq, _
                gOrderTable.sJDate, gOrderTable.sJGbn, gOrderTable.sJNo, _
                gOrderTable.sRack, gOrderTable.sPos, _
                gOrderTable.sRegNo, gOrderTable.sName, gOrderTable.sSex, _
                gOrderTable.sEmer, gOrderTable.sReRun, gOrderTable.sOther, "", "", "0", gOrderTable.iOrdCnt)
                
    Set objLD = Nothing
    
    'LastWSeq를 갱신
    gsLastWSeq = gOrderTable.sWSeq
    
    Exit Sub
    
ErrHandler:
    Set objLD = Nothing
    ViewMsg "RegOrder - Err(" & Err.Description & ")"
End Sub

Public Function RegResult(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
    ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, ByVal sIFSpcCd As String, Optional ByVal iCnt As Integer) As String
    
    On Error GoTo ErrHandler
    
    'iMode = 0 ---> 한 검사항목의 결과를 자동 등록
    'iMode = 1 ---> 한 샘플씩 자동 등록
    'iMode = 2 ---> Batch방식에 사용 여러 샘플 한 번에 등록
    
    Dim vIFItemCnt, vTmp, vChk
    Dim i%, j%, k%, iExist%
    Dim sTmp$, sCIFRstCd$, sCRst1$, sCRst2$
    Dim sWDate$, sWSeq$, sJDate$, sJGbn$, sJNo$, sRack$, sPos$, sRegNo$, sName$, sSex$, sEmer$, sReRun$, sOther$
    Dim sTIFSeq$, sTRst1$, sTRst2$
    Dim sIFSeq$, sRtnVal$
    Dim objLD As Object
    
    Set objLD = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    With gfIFDisplayForm.spdIntList
        Select Case iMode
            Case 0, 1
                sWDate = Format(Now, "YYYYMMDD")
        
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
                    
                    Call .GetText(16, CInt(sCRow), vIFItemCnt)
                    
                    For i = 1 To CInt(vIFItemCnt)
                        Call .GetText(16 + i, CInt(sCRow), vTmp)
                        
                        sTmp = CStr(vTmp)
                        
                        sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
                        
                        If Len(sIFSeq) = 3 Then
                            If ConvertIFItemInfo(8, sIFSeq) = sCIFRstCd Then
                                Exit For
                            End If
                        ElseIf Len(sIFSeq) = 2 Then
                            Exit For
                        End If
                    Next
                    
                    sRtnVal = objLD.Add_IFResult(gsMachineCd, 0, sWDate, sWSeq, _
                                  sIFSeq, sJDate, sJGbn, sJNo, sRack, sPos, sRegNo, sName, sSex, sEmer, sReRun, _
                                  sOther, sCRst1, sCRst2, "0", iRstCnt)
                    
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
                        
                        sTIFSeq = sTIFSeq & sIFSeq & "|"
                        sTRst1 = sTRst1 & sRst1 & "|"
                        sTRst2 = sTRst2 & sRst2 & "|"
                    Next
                    
                    sRtnVal = objLD.Add_IFResult(gsMachineCd, 1, sWDate, sWSeq, _
                                  sTIFSeq, sJDate, sJGbn, sJNo, sRack, sPos, sRegNo, sName, sSex, sEmer, sReRun, _
                                  sOther, sTRst1, sTRst2, "0", CInt(Val(vIFItemCnt)))
                    
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
    
    Set objLD = Nothing
    
    Exit Function
    
ErrHandler:
    Set objLD = Nothing
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

Public Function ReOrder_IFSeq(ByVal sBuf As String)
    Dim aSmall() As String
    Dim aLarge() As String
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
    
    For i = 1 To MAXIFITEM
        For j = 1 To iCnt
            If Format(i, "000") = aBuf(j) Then
                ReOrder_IFSeq = ReOrder_IFSeq & aBuf(j) & "|"
            Else
            End If
        Next
    Next
End Function

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

Public Sub ResultSpdClear()
    With gfIFDisplayForm.spdRst
        .BlockMode = True
        .Row = 1
        .Row2 = 15
        .Col = 1
        .Col2 = 3
        .Action = SS_ACTION_CLEAR_TEXT
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(0, 0, 0)
        .BlockMode = False
    End With
        
    With gfIFDisplayForm.spdRst2
        .BlockMode = True
        .Row = 1
        .Row2 = 95
        .Col = 1
        .Col2 = 3
        .Action = SS_ACTION_CLEAR_TEXT
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(0, 0, 0)
        .BlockMode = False
    End With
End Sub

Public Sub ResultSpdClear1()
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
End Sub

Public Function RPDChk(ByVal sTCd As String, ByVal sBRst As String, ByVal sRst As String, ByVal sSex As String) As String
'    On Error GoTo ErrHandler
    
    Dim i%, iCNo%
    
    If sRst = "" Then
        RPDChk = "|||"
        Exit Function
    End If
    
'    For i = 1 To giOriginIFItemCnt
'        If gIFItem(i).s02 = Mid$(sTCd, 1, 1) And gIFItem(i).s03 = Mid$(sTCd, 2, 2) And _
'            gIFItem(i).s04 = Mid$(sTCd, 4, 3) And gIFItem(i).s05 = Mid$(sTCd, 7, 3) And _
'            gIFItem(i).s06 = Mid$(sTCd, 10, 4) Then
'
'            iCNo = i
'            Exit For
'        End If
'    Next
'
'    If iCNo = 0 Then
'        ViewMsg "결과 등록시 에러가 발생했습니다..."
'        RPDChk = "Error"
'        Exit Function
'    End If
'
'    If gIFItem(iCNo).s12 = "0" Then   'REFGBN=없음
'        RPDChk = "|||"
'        Exit Function
'    End If
'
'    If gIFItem(iCNo).s17 = "0" Then     '판정구분=없음
'        RPDChk = "|||"
'        Exit Function
'    End If
'
'    If gIFItem(iCNo).s12 = "1" Then    'REFGBN=문자
'        '이 경우는 Panic, Delta 없음
'        If gIFItem(iCNo).s17 = "3" Then     '판정구분=OtherFlag
'            If sRst = gIFItem(iCNo).s27 Then
'                RPDChk = "|||"
'            Else
'                RPDChk = gIFItem(iCNo).s26 & "|||"
'            End If
'        Else                                '판정구분=없음
'            RPDChk = "|||"
'        End If
'
'        Exit Function
'    End If
'
'    If IsNumeric(sRst) = True Then
'        If gIFItem(iCNo).s12 = "2" And gIFItem(iCNo).s17 = "1" Then
'            'REFGBN=숫자(Low~High), 판정구분=Low/High
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s20) - CDbl(gIFItem(iCNo).s25)) Then
'                'REFLOWM(-GRAYLOWERM)
'                    RPDChk = RPDChk & "L|"
'                ElseIf (CDbl(gIFItem(iCNo).s20) - CDbl(gIFItem(i).s25)) <= CDbl(sRst) And _
'                    CDbl(sRst) <= (CDbl(gIFItem(iCNo).s21) + CDbl(gIFItem(i).s24)) Then
'                'REFLOWM(-GRAYLOWERM) ~ REFHIGHM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "|"
'                ElseIf CDbl(sRst) > (CDbl(gIFItem(iCNo).s21) + CDbl(gIFItem(iCNo).s24)) Then
'                'REFHIGHM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "H|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s28) - CDbl(gIFItem(iCNo).s33)) Then
'                'REFLOWf(-GRAYLOWERf)
'                    RPDChk = RPDChk & "L|"
'                ElseIf (CDbl(gIFItem(iCNo).s28) - CDbl(gIFItem(i).s33)) <= CDbl(sRst) And _
'                    CDbl(sRst) <= (CDbl(gIFItem(iCNo).s29) + CDbl(gIFItem(i).s32)) Then
'                'REFLOWf(-GRAYLOWERf) ~ REFHIGHf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "|"
'                ElseIf CDbl(sRst) > (CDbl(gIFItem(iCNo).s29) + CDbl(gIFItem(iCNo).s32)) Then
'                'REFHIGHf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "H|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(iCNo).s12 = "2" And gIFItem(iCNo).s17 = "2" Then
'        'REFGBN=숫자(Low~High), 판정구분이 NEG/POS
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s20) - CDbl(gIFItem(iCNo).s25)) Then
'                'REFLOWM(-GRAYLOWERM)
'                    RPDChk = RPDChk & "Pos|"
'                ElseIf (CDbl(gIFItem(iCNo).s20) - CDbl(gIFItem(i).s25)) <= CDbl(sRst) And _
'                    CDbl(sRst) <= (CDbl(gIFItem(iCNo).s21) + CDbl(gIFItem(i).s24)) Then
'                'REFLOWM(-GRAYLOWERM) ~ REFHIGHM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "|"
'                ElseIf CDbl(sRst) > (CDbl(gIFItem(iCNo).s21) + CDbl(gIFItem(iCNo).s24)) Then
'                'REFHIGHM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "Pos|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s28) - CDbl(gIFItem(iCNo).s33)) Then
'                'REFLOWf(-GRAYLOWERf)
'                    RPDChk = RPDChk & "Pos|"
'                ElseIf (CDbl(gIFItem(iCNo).s28) - CDbl(gIFItem(i).s33)) <= CDbl(sRst) And _
'                    CDbl(sRst) <= (CDbl(gIFItem(iCNo).s29) + CDbl(gIFItem(i).s32)) Then
'                'REFLOWf(-GRAYLOWERf) ~ REFHIGHf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "|"
'                ElseIf CDbl(sRst) > (CDbl(gIFItem(iCNo).s29) + CDbl(gIFItem(iCNo).s32)) Then
'                'REFHIGHf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "Pos|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(iCNo).s12 = "2" And gIFItem(iCNo).s17 = "3" Then   '판정구분이 OtherFlag
'            'REFGBN=숫자(Low~High), 판정구분이 OtherFlag
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s20) - CDbl(gIFItem(iCNo).s25)) Then
'                'REFLOWM(-GRAYLOWERM)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                ElseIf (CDbl(gIFItem(iCNo).s20) - CDbl(gIFItem(i).s25)) <= CDbl(sRst) And _
'                    CDbl(sRst) <= (CDbl(gIFItem(iCNo).s21) + CDbl(gIFItem(i).s24)) Then
'                'REFLOWM(-GRAYLOWERM) ~ REFHIGHM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "|"
'                ElseIf CDbl(sRst) > (CDbl(gIFItem(iCNo).s21) + CDbl(gIFItem(iCNo).s24)) Then
'                'REFHIGHM(+GRAYUPPERM)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s28) - CDbl(gIFItem(iCNo).s33)) Then
'                'REFLOWf(-GRAYLOWERf)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                ElseIf (CDbl(gIFItem(iCNo).s28) - CDbl(gIFItem(i).s33)) <= CDbl(sRst) And _
'                    CDbl(sRst) <= (CDbl(gIFItem(iCNo).s29) + CDbl(gIFItem(i).s32)) Then
'                'REFLOWf(-GRAYLOWERf) ~ REFHIGHf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "|"
'                ElseIf CDbl(sRst) > (CDbl(gIFItem(iCNo).s29) + CDbl(gIFItem(iCNo).s32)) Then
'                'REFHIGHf(+GRAYUPPERf)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'
'
'        ElseIf gIFItem(iCNo).s12 = "3" And gIFItem(iCNo).s17 = "1" Then
'            'REFGBN=숫자(상한), 판정구분이 LOW/HIGH
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) > (CDbl(gIFItem(iCNo).s22) + CDbl(gIFItem(iCNo).s24)) Then
'                'UPPERLIMITM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "H|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) > (CDbl(gIFItem(iCNo).s30) + CDbl(gIFItem(iCNo).s32)) Then
'                'UPPERLIMITf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "H|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(iCNo).s12 = "3" And gIFItem(iCNo).s17 = "2" Then
'            'REFGBN=숫자(상한), 판정구분이 NEG/POS
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) > (CDbl(gIFItem(iCNo).s22) + CDbl(gIFItem(iCNo).s24)) Then
'                'UPPERLIMITM(+GRAYUPPERM)
'                    RPDChk = RPDChk & "Pos|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) > (CDbl(gIFItem(iCNo).s30) + CDbl(gIFItem(iCNo).s32)) Then
'                'UPPERLIMITf(+GRAYUPPERf)
'                    RPDChk = RPDChk & "Pos|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(iCNo).s12 = "3" And gIFItem(iCNo).s17 = "3" Then
'            'REFGBN=숫자(상한), 판정구분이 OtherFlag
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) > (CDbl(gIFItem(iCNo).s22) + CDbl(gIFItem(iCNo).s24)) Then
'                'UPPERLIMITM(+GRAYUPPERM)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) > (CDbl(gIFItem(iCNo).s30) + CDbl(gIFItem(iCNo).s32)) Then
'                'UPPERLIMITf(+GRAYUPPERf)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'
'
'        ElseIf gIFItem(iCNo).s12 = "4" And gIFItem(iCNo).s17 = "1" Then
'            'REFGBN=숫자(하한), 판정구분이 LOW/HIGH
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s23) - CDbl(gIFItem(iCNo).s25)) Then
'                'LOWERLIMITM(-GRAYLOWERM)
'                    RPDChk = RPDChk & "L|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s31) - CDbl(gIFItem(iCNo).s33)) Then
'                'LOWERLIMITf(-GRAYLOWERf)
'                    RPDChk = RPDChk & "L|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(iCNo).s13 = "4" And gIFItem(iCNo).s17 = "2" Then
'            'REFGBN=숫자(하한), 판정구분이 NEG/POS
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s23) - CDbl(gIFItem(iCNo).s25)) Then
'                'LOWERLIMITM(-GRAYLOWERM)
'                    RPDChk = RPDChk & "Pos|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s31) - CDbl(gIFItem(iCNo).s33)) Then
'                'LOWERLIMITf(-GRAYLOWERf)
'                    RPDChk = RPDChk & "Pos|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(iCNo).s13 = "4" And gIFItem(iCNo).s17 = "3" Then
'            'REFGBN=숫자(하한), 판정구분이 OtherFlag
'            If sSex = "0" Or (CInt(sSex) Mod 2) = 1 Then
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s23) - CDbl(gIFItem(iCNo).s25)) Then
'                'LOWERLIMITM(-GRAYLOWERM)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            Else
'                If CDbl(sRst) < (CDbl(gIFItem(iCNo).s31) - CDbl(gIFItem(iCNo).s33)) Then
'                'LOWERLIMITf(-GRAYLOWERf)
'                    RPDChk = RPDChk & gIFItem(iCNo).s26 & "|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        End If
'
'        'Panic
'        If gIFItem(i).s13 = "0" Then
'            RPDChk = RPDChk & "|"
'        ElseIf gIFItem(i).s13 = "1" Then
'            If CDbl(sRst) < CDbl(gIFItem(iCNo).s18) Then
'            'PANIC LOW
'                RPDChk = RPDChk & "P|"
'            ElseIf CDbl(gIFItem(iCNo).s18) <= CDbl(sRst) And _
'                CDbl(sRst) <= CDbl(gIFItem(iCNo).s19) Then
'            'PANIC LOW ~ PANIC HIGH
'                RPDChk = RPDChk & "|"
'            ElseIf CDbl(sRst) > CDbl(gIFItem(iCNo).s19) Then
'            'PANIC HIGH
'                RPDChk = RPDChk & "P|"
'            Else
'                RPDChk = RPDChk & "|"
'            End If
'        Else
'            RPDChk = RPDChk & "|"
'        End If
'
'        'Delta
'        If gIFItem(i).s14 = "0" Then
'            RPDChk = RPDChk & "|"
'        ElseIf gIFItem(i).s14 = "1" Then
'        '절대값인 경우
'            If sBRst = "" Then
'                RPDChk = RPDChk & "|"
'            Else
'                If Abs(CDbl(sRst) - CDbl(sBRst)) >= CDbl(gIFItem(iCNo).s15) Then
'                    RPDChk = RPDChk & "D|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        ElseIf gIFItem(i).s14 = "2" Then
'        '%인 경우
'            If sBRst = "" Then
'                RPDChk = RPDChk & "|"
'            Else
'                If (Abs(CDbl(sRst) - CDbl(sBRst)) / CDbl(sRst) * 100) >= CDbl(gIFItem(iCNo).s15) Then
'                    RPDChk = RPDChk & "D|"
'                Else
'                    RPDChk = RPDChk & "|"
'                End If
'            End If
'        Else
'            RPDChk = RPDChk & "|"
'        End If
'    End If
'
'    Exit Function
'
'ErrHandler:
'    RPDChk = "|||"
End Function

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

Public Sub TxtTypeOnlyAlphaNumeric(SomeTextBox3 As TextBox)
    SomeTextBox3.IMEMode = 3    'IME사용못함
End Sub

Public Function ViewIFResult(ByVal iCRow As Integer, ByVal iRstCnt As Integer, _
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
                    If ConvertIFItemInfo(8, gResultTable(1).sTestCd(j)) = sCIFRstCd Then
                        Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, sCRst2, _
                                    RGB(0, 0, 0), RGB(255, 255, 255))
                        
                        Exit For
                    End If
                Next
            Next
        End With
        
        'Result spdRst에 표시
        Call DisplayResult(iCRow)
        
        If Val(vCRstCnt) >= Val(vIFCnt) Then
            ViewIFResult = "DONE"
        Else
            ViewIFResult = "MORE"
        End If
    End With
End Function

Public Function ViewIFResult1(ByVal iCRow As Integer, ByVal iRstCnt As Integer, _
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
                    If Left$(gResultTable(1).sTestCd(j), 1) = "C" Then
                        If gResultTable(1).sTestCd(j) = sCIFRstCd Then
                            If sCRst2 = "Low" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(220, 220, 255))
                            ElseIf sCRst2 = "High" Or sCRst2 = "Positive" Then
                                Call CurRstDisplay(iCRow, ConvertIFItemInfo(4, gResultTable(1).sTestCd(j)), sCRst1, "", _
                                         RGB(0, 0, 0), RGB(255, 220, 220))
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
        
        'Result spdRst에 표시
        Call DisplayResult1(iCRow)
        
        If Val(vCRstCnt) >= Val(vIFCnt) Then
            ViewIFResult1 = "DONE"
        Else
            ViewIFResult1 = "MORE"
        End If
    End With
End Function

Public Sub ViewMsg(ByVal sMsg As String)
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "ViewMsg.Hwnd")
    
    Call SetWindowText(Val(sBuf), sMsg)
End Sub

Public Function 정수부구하기(ByVal sTmp As String) As String
    Dim Pos%
    
    Pos = InStr(1, sTmp, ".")
    
    If Pos = 0 Then
        정수부구하기 = sTmp
    Else
        정수부구하기 = Left$(sTmp, Pos - 1)
    End If
End Function

Public Function 소수부구하기(ByVal sTmp As String, ByVal iDig As Integer) As String
    Dim Pos%
    
    Pos = InStr(1, sTmp, ".")
    
    If Pos = 0 Then
        소수부구하기 = ""
    Else
        소수부구하기 = Mid$(sTmp, Pos + 1, iDig)
    End If
End Function

Public Function 소수부구하기2(ByVal sTmp As String) As String
    Dim Pos%
    
    Pos = InStr(1, sTmp, ".")
    
    If Pos = 0 Then
        소수부구하기2 = ""
    Else
        소수부구하기2 = Mid$(sTmp, Pos + 1)
    End If
End Function
