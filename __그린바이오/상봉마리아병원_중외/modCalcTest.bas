Attribute VB_Name = "modCalcTest"
Option Explicit

Public gTC      As String
Public gTG      As String
Public gHDL     As String
Public gLDLC    As String
Public gBUN     As String
Public gCREA    As String
Public geGFR    As String
Public gBCRatio As String

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

Public Function Check_OpLevel(ByVal strOp As String) As Integer

    Check_OpLevel = 0
    
    Select Case strOp
        Case "+", "-":      Check_OpLevel = 1
        Case "\", "%":      Check_OpLevel = 2
        Case "*", "/":      Check_OpLevel = 3
        Case "^":           Check_OpLevel = 4
    End Select

End Function

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


'-- strInFormula    : 계산식      (25 * 7) /2
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

    ' '**'를 '^'로 바꿈         ==> 제곱승
    Do While InStr(strInFormula, "**") > 0
        nCurrPos = InStr(strInFormula, "**")
        strInFormula = LEFT(strInFormula, nCurrPos - 1) & "^" & Mid(strInFormula, nCurrPos + 2)
    Loop

    ' 'MOD'를 '%'로 바꿈
    Do While InStr(UCase(strInFormula), "MOD") > 0
        nCurrPos = InStr(UCase(strInFormula), "MOD")
        strInFormula = LEFT(strInFormula, nCurrPos - 1) & "%" & Mid(strInFormula, nCurrPos + 3)
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
    'ViewMsg "계산식에 오류가 있습니다."

    If Not IsMissing(nState) Then nState = False

    Exit Function

End Function


