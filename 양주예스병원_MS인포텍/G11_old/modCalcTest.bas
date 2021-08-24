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


'-- strInFormula    : ����      (25 * 7) /2
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
    ' ����, ������, ��ȣ ���� ������ �迭�� ����
    nCnt = 0        ' �迭�� ����� ����
    nStartPos = 0   ' ���ڸ� �ϳ��� �������� ���� ������ġ ����

    ' '**'�� '^'�� �ٲ�         ==> ������
    Do While InStr(strInFormula, "**") > 0
        nCurrPos = InStr(strInFormula, "**")
        strInFormula = LEFT(strInFormula, nCurrPos - 1) & "^" & Mid(strInFormula, nCurrPos + 2)
    Loop

    ' 'MOD'�� '%'�� �ٲ�
    Do While InStr(UCase(strInFormula), "MOD") > 0
        nCurrPos = InStr(UCase(strInFormula), "MOD")
        strInFormula = LEFT(strInFormula, nCurrPos - 1) & "%" & Mid(strInFormula, nCurrPos + 3)
    Loop

    nFlag = False   ' ������ �������� ���� (Ư�����ڵ�)

    For i = 1 To Len(strInFormula)
        strChar = Mid$(strInFormula, i, 1)   ' �ѱ��ھ� �������� ���������� ��ȣ���� ��
        If Trim(strChar) <> "" Then
            If IsNumeric(strChar) Or (strChar = ".") Then   ' ���ڿ� �Ҽ����� ���ڷ� ���
                If nStartPos = 0 Then nStartPos = i
                nFlag = True
            Else
                If nStartPos > 0 Then   ' ���� ���� ���� ���� ���ڰ� ���� ��� ���ڸ� ����
                    strFormula(nCnt) = Mid$(strInFormula, nStartPos, i - nStartPos)
                    nCnt = nCnt + 1
                    nStartPos = 0
                End If
                If (strChar Like "[()]") Or IsOp(strChar) Then
                    ' ��ȣ �� �����ڸ� ����
                    strFormula(nCnt) = strChar
                    nCnt = nCnt + 1
                    nFlag = True
                End If
            End If
            If nFlag = True Then    ' ��ġ, ��ȣ, ������ ���� �̻��� ���ڰ� �ִ��� Ȯ��
                nFlag = False
            Else
                GoTo Err_Process
            End If
        End If
    Next i

    If nStartPos > 0 Then   ' ���ڰ� ������ ���� ���� ��� ���ڸ� ����
        strFormula(nCnt) = Mid$(strInFormula, nStartPos, i - nStartPos)
        nCnt = nCnt + 1
        nStartPos = 0
    End If

    ' ��ȣ(-)�� �����Ѵ�. ('(', '������' ������ ������ '+', '-'�� ��ȣ��.)
    nFlag = True    ' '(', '������'�� ���Դ��� ����

    For i = 0 To nCnt - 1
        If nFlag = True Then
            If strFormula(i) Like "[+-]" Then      ' ��ȣ �߰�
                If strFormula(i) = "-" Then
                    If IsNumeric(strFormula(i + 1)) Then
                        ' ��ȣ(-)�� ������ ������ ���ڿ� ����.
                        strFormula(i + 1) = Trim(Str(Val(strFormula(i + 1)) * -1))
                    Else
                        '��ȣ ������ �����ڰ� ����
                        GoTo Err_Process
                    End If
                End If
                strFormula(i) = ""     ' ��ȣ�� �ִ� �ڸ��� Null�� ü��
            End If
        End If
        ' '(', '������' ������ ������ '+', '-'�� ��ȣ�̹Ƿ� '(', '������' Ȯ��
        If IsOp(strFormula(i)) Then
            nFlag = True
        Else
            nFlag = False
        End If
    Next i

    ' ��ȣ(-)�� �����Ҷ� �߻��� Null ���� (strFormula2�� �ű��� �ٽ� strFormula�� �ű�)
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

    ' �Ŀ� �����ڰ� ������ ��� ��ȣ ���� (��:(1))
    For i = 0 To nCnt - 1
        If IsOp(strFormula(i)) Then
            Exit For
        End If
    Next i

    If i = nCnt Then
        ' ���ʿ��� ��ȣ ����
        For i = 0 To nCnt - 1
            If strFormula(i) Like "[()]" Then
                strFormula(i) = ""
            End If
        Next i

        ' ���ʿ��� ��ȣ�� �����Ҷ� �߻��� Null ���� (strFormula2�� �ű��� �ٽ� strFormula�� �ű�)
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

    ' ��ȣ�� �ִ��� Ȯ�� �� �켱 ������ ���� �����ڸ� ã�Ƽ� ��ȣ�� ���´�. (�ǹ̾��� ��ȣ�� ������.)
    nStartPos = 0       ' �켱 ���� �� ���� ��ġ

    Do
        nFlag = True    ' �۾� ���� Flag
        nFlag2 = False  ' ��ȣ ���� Flag
        nFlag3 = False  ' nStartPos ���濩��
        nOldStartPos = nStartPos

        For i = nStartPos To nCnt - 1
            If IsOp(strFormula(i)) Then
                nOpLevel = Check_OpLevel(strFormula(i))    ' ���� �켱���� Level
                nCurrPos = i
                nLevel = 0          ' ��ȣ�� ���� ������ �ϳ��� �������� ��ȣ�� ���� ���� ��

                For j = (i - 1) To 0 Step -1    ' �������� ������ ��ȣ�� �� ��ġ Ȯ��
                    If nLevel = 0 Then
                        If IsOp(strFormula(j)) Then
                            nLeftPos = j + 1    ' ��ȣ�� ���Ե� ��ġ
                            Exit For
                        End If
                    End If
                    If strFormula(j) = ")" Then nLevel = nLevel + 1
                    If strFormula(j) = "(" Then nLevel = nLevel - 1
                Next j

                If j = -1 Then nLeftPos = 0     ' ��ȣ�� ���Ե� ��ġ

                nLevel = 0          ' ��ȣ�� ���� ������ �ϳ��� �������� ��ȣ�� ���� ���� ��

                For j = (i + 1) To (nCnt - 1)   ' �������� ������ �켱������ �� ���� �����ڰ� �ִ��� Ȯ��
                    If nLevel = 0 Then
                        If IsOp(strFormula(j)) Then
                            If nOpLevel >= Check_OpLevel(strFormula(j)) Then
                                ' ������ �����ڰ� �켱������ ���� ���� ��� ���� �����ڸ� ��ȣ�� ���´�.
                                nRightPos = j    ' ��ȣ�� ���Ե� ��ġ
                                nFlag2 = True
                                nFlag3 = True
                                nFlag = False:  Exit For
                            Else
                                ' ������ �����ڰ� �켱������ �����Ƿ� ����ġ(J)���� �ٽú�
                                nStartPos = j
                                nFlag3 = True
                                nFlag = False:  Exit For
                            End If
                        End If
                        If strFormula(j) = ")" Then
                            nRightPos = j    ' ��ȣ�� ���Ե� ��ġ
                            nFlag2 = True
                            nFlag3 = True
                            nFlag = False:  Exit For
                        End If
                    End If
                    If strFormula(j) = "(" Then nLevel = nLevel + 1
                    If strFormula(j) = ")" Then nLevel = nLevel - 1
                Next j

                If j = nCnt Then
                    nRightPos = nCnt   ' ��ȣ�� ���Ե� ��ġ
                    nFlag2 = True
                    Exit For
                End If

                If nFlag = False Then Exit For
            End If
        Next i

        nOldCnt = nCnt

        If nFlag2 = True Then
            If Not (strFormula(IIf(nLeftPos = 0, 0, nLeftPos - 1)) = "(" And strFormula(nRightPos) = ")") Then
                ' ��ȣ ����
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
            ' ������ �켱������ �з� ��ȣ�� ������ ���� �����ڰ� �տ� ���� �� �����Ƿ�
            If nOldStartPos <> 0 Then nStartPos = 0
        End If
    Loop Until (nLeftPos = 0 Or nLeftPos = 1) And (nRightPos = nOldCnt Or nRightPos = nOldCnt - 1)

    ' PreFix �� �ٲ۴�.
    For i = 0 To nCnt - 1
        If strFormula(i) = "(" Then
            nLevel = 0          ' ��ȣ�� ���� ������ �ϳ��� �������� ��ȣ�� ���� ���� ��
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
                If nLevel = -1 Then Exit For    ' ��ȣ�ȿ� ������ ����
            Next j
            If nFlag = True Then
                If Trim(strChar) <> "" Then
                    strFormula(i) = strChar    ' ��ȣ('(')�� �����ڷ� ��ü
                    strFormula(j) = ""
                End If
            End If
        End If
    Next i

    ' ���ʿ��� ��ȣ ����
    For i = 0 To nCnt - 1
        If strFormula(i) Like "[()]" Then
            strFormula(i) = ""
        End If
    Next i

    ' ���ʿ��� ��ȣ�� �����Ҷ� �߻��� Null ���� (strFormula2�� �ű��� �ٽ� strFormula�� �ű�)
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

    ' ���ؿ� �����鼭 ���
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
                    ' ���
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
    'ViewMsg "���Ŀ� ������ �ֽ��ϴ�."

    If Not IsMissing(nState) Then nState = False

    Exit Function

End Function


