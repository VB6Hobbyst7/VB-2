Attribute VB_Name = "modVariable"
Option Explicit

Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, _
                            ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, _
                            ByVal nSize As Long) As Long

Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, _
                            ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As String) As Long

Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A
Public Const HKEY_CURRENT_CONFIG = &H80000005

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

'-- 2002-05-26 JJH 추가
'   특정항목인경우 ComboBox입력으로....
'----------------------------------------
Public gsComboBox_InputItems   As String
Public gsComboBox_InputResults As String
'----------------------------------------

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

Public Function Check_OpLevel(ByVal strOp As String) As Integer
    Check_OpLevel = 0

    Select Case strOp
        Case "+", "-":      Check_OpLevel = 1
        Case "\", "%":      Check_OpLevel = 2
        Case "*", "/":      Check_OpLevel = 3
        Case "^":           Check_OpLevel = 4
    End Select
End Function

Public Function ConvertIFItemInfo(ByVal iMode As Integer, ByVal sComp As String, Optional ByVal sUL$) As String
    Dim i%

    Select Case iMode
        '서버쪽코드를 IFSEQ로
        Case 1
            If sUL = "U" Then
                For i = 1 To giOriginIFItemCnt
                    If UCase(gIFItem(i).s06) = sComp Then
                        ConvertIFItemInfo = gIFItem(i).s01
                        Exit For
                    End If
                Next
    
                For i = 1 To giOriginCalItemCnt
                    If UCase(gCalItem(i).s03) = sComp Then
                        ConvertIFItemInfo = gCalItem(i).s01
                        Exit For
                    End If
                Next
            ElseIf sUL = "L" Then
                For i = 1 To giOriginIFItemCnt
                    If LCase(gIFItem(i).s06) = sComp Then
                        ConvertIFItemInfo = gIFItem(i).s01
                        Exit For
                    End If
                Next
    
                For i = 1 To giOriginCalItemCnt
                    If LCase(gCalItem(i).s03) = sComp Then
                        ConvertIFItemInfo = gCalItem(i).s01
                        Exit For
                    End If
                Next
            Else
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
            End If
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

        Case Else

    End Select
End Function

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

Public Sub GetMachineInfo()
    Dim RetVal As Long
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
    
'-- 2002-05-26 JJH 추가
'   특정항목인경우 ComboBox입력으로....
'ComboBox_InputItems
'-------------------------------------------------------------------------------------------------------
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.ComboBox_InputItems")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.ComboBox_InputItems", "")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
        gsComboBox_InputItems = ""    'Default No
    Else
        gsComboBox_InputItems = sBuf
    End If
    
'ComboBox_InputResults
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.ComboBox_InputIResults")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.ComboBox_InputIResults", "")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
        gsComboBox_InputResults = ""    'Default No
    Else
        gsComboBox_InputResults = sBuf
    End If
'-------------------------------------------------------------------------------------------------------
    
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
    Dim RetVal As Long
    
'Machine Code
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineCd", "", sBuf, 255, App.Path & "\장비코드.ini")
    
    If RetVal = 0 Then
        MsgBox "장비코드 설정이 되어 있지 않습니다. 프로그램을 제대로 실행할 수 없습니다!!", vbCritical, "장비코드.ini 설정"
    End If
    
    gsMachineCd = Left(sBuf, RetVal) 'Machine Name
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineNm", "", sBuf, 255, App.Path & "\장비코드.ini")
    
    If RetVal = 0 Then
        MsgBox "장비코드 설정이 되어 있지 않습니다. 프로그램을 제대로 실행할 수 없습니다!!", vbCritical, "장비코드.ini 설정"
    End If
    
    gsMachineNm = Left(sBuf, RetVal)
    
'Machine Exe
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineExe", "", sBuf, 255, App.Path & "\장비코드.ini")
    
    If RetVal = 0 Then
        MsgBox "장비코드 설정이 되어 있지 않습니다. 프로그램을 제대로 실행할 수 없습니다!!", vbCritical, "장비코드.ini 설정"
    End If
    
    gsMachineExe = Left(sBuf, RetVal)
End Sub

Public Function JudgePanicDelta$(ByVal sIFSeq$, ByVal sCompRst$, ByVal sPrevRst$, ByVal sDateDiff$, sPanFlag$, sDelFlag$)
    On Error GoTo ErrHandler
    
    Dim i%
    Dim sPanL$, sPanH$, sDelGbn$, sDelL$, sDelH$
    
    If Len(sIFSeq) = 3 Then
        For i = 1 To giOriginIFItemCnt
            If sIFSeq = gIFItem(i).s01 Then
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
                sPanL = gCalItem(i).s12
                sPanH = gCalItem(i).s13
                
                sDelGbn = gCalItem(i).s14
                sDelL = gCalItem(i).s15
                sDelH = gCalItem(i).s16
                
                Exit For
            End If
        Next
    End If
    
    sPanFlag = "": sDelFlag = ""
    
    If Trim(sCompRst) = "" Then Exit Function
    
    If IsNumeric(sCompRst) = False Then
        If sPanL <> "" Or sPanH <> "" Then
            sPanFlag = "P"
        End If
        
        JudgePanicDelta = sCompRst
    
        Exit Function
    End If
    
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
    ViewMsg "JudgePanicDelta - Err(" & Err.Description & ")"
End Function

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
        If sJGbn = "1" Then
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
        sDataRow(i) = GetByOneUserSymbol(sCalItem, sCalItem, Chr(3))
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

Public Sub RegViewMsgHwnd(ByVal lnHwnd As Long)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "ViewMsg.Hwnd", CStr(lnHwnd))
    
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
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

Public Sub TxtTypeOnlyAlphaNumeric(SomeTextBox3 As TextBox)
    SomeTextBox3.IMEMode = 3    'IME사용못함
End Sub

Public Sub ViewMsg(ByVal sMsg As String)
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "ViewMsg.Hwnd")
    
    Call SetWindowText(Val(sBuf), sMsg)
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, ByRef DriverName As String, ByRef PrinterPort As String)
    Dim R       As Integer
    Dim iDriver As Integer
    Dim iPort   As Integer
    
    DriverName = ""
    PrinterPort = ""
        
    'The driver name is first in the string terminated by acomma
    iDriver = InStr(Buffer, ",")
    
    If iDriver > 0 Then
        'Strip out the drivername
        DriverName = Left(Buffer, iDriver - 1)
        
        'The port name is the second entry after the drivername
        'separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")
                
        If iPort > 0 Then
            'Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, iPort - iDriver - 1)
        End If
    End If
End Sub

Public Sub Set_DefaultPrinter(ByVal PrinterName As String)
    Dim R       As Integer
    Dim L       As Long
    Dim Buffer  As String
    Dim DeviceName  As String
    Dim DriverName  As String
    Dim PrinterPort As String
    Dim DeviceLine  As String
        
    'Get the printer information for the currently selected printer. The information is taken from the REGISTRY file.
    Buffer = Space(1024)
        
    R = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))
    
    'Parse the driver name and port name out of thebuffer
    GetDriverAndPort Buffer, DriverName, PrinterPort
        
    If Trim(DriverName) = "" Or Trim(PrinterPort) = "" Then
        Exit Sub
    End If
        
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    
    ' Store the new printer information in the Registry
    R = WriteProfileString("windows", "Device", DeviceLine)
    
    ' Cause all applications to reload the INI file:
    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
End Sub

Public Function Get_DefaultPrinter() As String
    Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long, v$, R As Long
    Dim RegPath$
    
    '/* 기본프린터의 내용이 있는 Registry Path */
    RegPath = "System\CurrentControlSet\Control\Print\Printers"
     
    RetVal$ = ""
    R = RegOpenKeyEx(HKEY_CURRENT_CONFIG, RegPath, 0, KEY_ALL_ACCESS, hSubKey)
                                                     'KEY_ALL_CLASSES
    If R <> ERROR_SUCCESS Then GoTo Quit_Now
    
    SZ = 256: v$ = String$(SZ, 0)
    
    R = RegQueryValueEx(hSubKey, "Default", 0, dwType, ByVal v$, SZ)
    
    If R = ERROR_SUCCESS And dwType = REG_SZ Then
        RetVal$ = Left(v$, SZ - 1)
    Else
        RetVal$ = ""
    End If
    
    R = RegCloseKey(hSubKey)

Quit_Now:
    Get_DefaultPrinter = RetVal$
End Function

Public Function fGetCurDSN(ByVal sBuf As String) As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\Ack_if\Program Config\" & sBuf, "DSN", "IFDSN")
        
        If bRetVal = True Then
            fGetCurDSN = "IFDSN"
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            fGetCurDSN = "IFDSN"
        End If
    Else
        fGetCurDSN = sBuf
    End If
End Function

Public Function GetDeltaPanicOrientation() As String
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "DeltaPanic.Orientation")
        
    GetDeltaPanicOrientation = sBuf
End Function

Public Function GetDeltaPanicPrintTitle() As String
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "DeltaPanic.PrintTitle")
        
    GetDeltaPanicPrintTitle = sBuf
End Function

Public Function GetDeltaPanicRemarkTitle() As String
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "DeltaPanic.RemarkTitle")
        
    GetDeltaPanicRemarkTitle = sBuf
End Function

Public Function GetPrintTail() As String
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Print.Tail")
        
    GetPrintTail = sBuf
End Function

Public Sub MsgLog(ByVal sBuf$)
    On Error GoTo ErrHandler
    
    Open App.Path & "\Err_if.log" For Append As #3
    Print #3, sBuf;
    Close #3
    
    Exit Sub
    
ErrHandler:
End Sub
