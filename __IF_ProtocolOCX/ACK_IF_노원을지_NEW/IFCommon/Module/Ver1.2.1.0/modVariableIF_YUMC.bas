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

'For Server ��� ��������
Public giServerOK As Integer

'For MiddleWare
Public gObjMW As Object
Public gObjMW2 As Object

'�������̽���� ���
'0=�ܹ���,
'1=�����(Rack Or Tray ��� ��������, But Rack/Pos ǥ��)
'2=�����(Rack Or Tray ��� ��������, But Tray/Pos ǥ��)
'3=�����(Rack Or Tray ��� ��������, But Tray/Cup ǥ��)
'4=�����(Rack/Pos ��� ����),
'5=�����(Tray/Pos ��� ����),
'6=�����(Tray/Cup ��� ����),
Public gsIFMode$
Public gsINITMode$  'Initialize ��ư ��� ��� - 0=������, 1=�����
Public gsTXMode$    '������۹�� ��� - 0=��ġ, 1=����Ÿ��(�׸� ����), 2=����Ÿ��(ȯ�ں� ����)
Public gsAPMode$    '�ڵ���� ���

Public gsIFVar1$, gsIFVar2$, gsIFVar3$, gsIFVar4$, gsIFVar5$

Public giBSRow%
Public giBERow%

'Comment����
Public gCommentCd() As String
'Interface�׸� �˻��׸��ڵ� 'H01001001NNNN
Public vIFItemCd() As Variant

Public iSpdBackColorOption As Integer

'ConverIFItemInfo�ÿ� ��ġ�ϴ� ����
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
' 2002-05-24 JJH �߰�
' ��ũ����Ʈ���뺴�� ( ����: �������Ϲ�ȣ, �������������, �˻��׸� �ߺ��������� �� )
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
                
                '-- ������ ����ŸRow üũ
                If giBERow - giBSRow + 1 > 2 Then
                    MsgBox "3�� �̻��� ������ �� �����ϴ�.", vbExclamation
                    Exit Function
                End If
                
                '--
                If giBERow <> .MaxRows Then
                    MsgBox "������ �� ���� ����Ʈ�� �����ϸ� �ȵ˴ϴ�.", vbExclamation
                    Exit Function
                End If
                
                '--
                If MsgBox("���õ� ����Ÿ�� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Function
                End If
                
                '-- �۾���ȣ
                .Row = giBSRow: .Col = 1: sWSEQ1 = Trim(.Text)
                .Row = giBERow: .Col = 1: sWSEQ2 = Trim(.Text)
                
                '-- ���Ϲ�ȣ üũ
                .Row = giBSRow: .Col = 8: sTmp1 = Trim(.Text)
                .Row = giBERow: .Col = 8: sTmp2 = Trim(.Text)
                If sTmp1 <> sTmp2 Then
                    MsgBox "���Ϲ�ȣ�� Ʋ���Ƿ� ������ �� �����ϴ�.", vbExclamation
                    Exit Function
                End If
                
                '-- ������ۿ���üũ
                .Row = giBSRow: .Col = 15: sTmp1 = Trim(.Text)
                .Row = giBERow: .Col = 15: sTmp2 = .Text
                If sTmp1 <> "N" Or sTmp2 <> "N" Then
                    MsgBox "�̹� ��������� �Ϸ�� ����Ÿ�̹Ƿ� ������ �� �����ϴ�.", vbExclamation
                End If
                
                '-- �˻��׸� �ߺ�üũ
                .Row = giBSRow: .Col = 14: ReDim sOrdCd1(Val(Trim(.Text))) As String
                .Row = giBERow: .Col = 14: ReDim sOrdCd2(Val(Trim(.Text))) As String
                For i = 1 To UBound(sOrdCd1)
                    .Row = giBSRow: .Col = 16 + i: sOrdCd1(i) = Trim(.Text)
                Next
                    
                For i = 1 To UBound(sOrdCd2)
                    .Row = giBERow: .Col = 16 + i: sOrdCd2(i) = Trim(.Text)
                    
                    For j = 1 To UBound(sOrdCd1)
                        If sOrdCd1(j) = sOrdCd2(i) Then
                            MsgBox "�˻��׸��� �ߺ��Ǿ� ������ �� �����ϴ�.", vbExclamation
                            Exit Function
                        End If
                    Next
                Next
                
                iAllCnt = UBound(sOrdCd1) + UBound(sOrdCd2)
                .Row = giBSRow: .Col = 14: .Text = CStr(iAllCnt)
                .Row = giBSRow: .Col = 16: .Text = CStr(iAllCnt)
                
                '�˻��׸� ���� �����
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
    ViewMsg "WMerge_SPD ���� - (" & Err.Description & ")"
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
            Call .SetText(2, .MaxRows, "1")     'üũ
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
            
            '���� Row ���
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
        If MsgBox("Interface List�� �����ϸ� �ش� List�� ����� ���� ���մϴ�." & vbCrLf & vbCrLf & _
            "����� ���� ���� �ʾҴٸ� '�ƴϿ�'�� �����Ͻʽÿ�." & vbCrLf & _
            "Interface List�� ���� �����Ͻðڽ��ϱ�?", vbYesNo, "��ü����Ʈ ȭ�� ���� Ȯ��") = vbYes Then
            
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
                    RGB(0, 0, 0), �����)
        End With
        
        .lblOrder = gOrderTable.sSampID
    End With
    
    'gOrderTable �ʱ�ȭ
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
    ' ����, ������, ��ȣ ���� ������ �迭�� ����
    nCnt = 0        ' �迭�� ����� ����
    nStartPos = 0   ' ���ڸ� �ϳ��� �������� ���� ������ġ ����
    
    ' '**'�� '^'�� �ٲ�
    Do While InStr(strInFormula, "**") > 0
        nCurrPos = InStr(strInFormula, "**")
        strInFormula = Left(strInFormula, nCurrPos - 1) & "^" & Mid(strInFormula, nCurrPos + 2)
    Loop
    
    ' 'MOD'�� '%'�� �ٲ�
    Do While InStr(UCase(strInFormula), "MOD") > 0
        nCurrPos = InStr(UCase(strInFormula), "MOD")
        strInFormula = Left(strInFormula, nCurrPos - 1) & "%" & Mid(strInFormula, nCurrPos + 3)
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
    ViewMsg "���Ŀ� ������ �ֽ��ϴ�."
    
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
                Case "M", "��", "1"
                    vSex = "M"
                Case "F", "��", "2"
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
                    '���ĺ� I(�빮�� ����)�� ã�´�.
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
                
                '���Ŀ� �ʿ��� Interface ����� ���۵Ǿ����� üũ
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
                
                '����� ���� ����� ��� ���۵Ǿ��ٸ�
                '��� ����� �������忡 ��Ÿ��
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
                    
                    '��� �׸��� IFSEQ("C1"�� ����) ���
                    sCRst = JudgeResultBySex(gCalItem(i).s01, sCRst, vSex, "", "", sCRst2, "", "")
                    
                    '������ ���۵Ǿ����� üũ
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
                    
                    '�������� ���� ������ �Ķ���� ��ȯ
                    iRstCnt = iRstCnt + 1
                    sIFRstCd = sIFRstCd & gCalItem(i).s01 & "|"
                    sRst1 = sRst1 & sCRst & "|"
                    'sCRst2�� JudgeResultBySex���� �޾� ��
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
        '�������ڵ带 IFSEQ��
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
            
        'IFSEQ�� �������ڵ��
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
            
        '�˻��׸���� IFSEQ��
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
            
        'IFSEQ�� �˻��׸������
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
        
        'IFORDCD�� IFSEQ��
        Case 5
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s01
                    Exit For
                End If
            Next
                  
        'IFSEQ�� IFORDCD��
        Case 6
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s03
                    Exit For
                End If
            Next
        
        'IFRSTCD�� IFSEQ��
        Case 7
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s04 = sComp Then
                    ConvertIFItemInfo2 = gIFItem(i).s01
                    Exit For
                End If
            Next
        
        'IFSEQ�� IFRSTCD��
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
        '�������ڵ带 IFSEQ��
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
            
        'IFSEQ�� �������ڵ��
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
            
        '�˻��׸���� IFSEQ��
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
            
        'IFSEQ�� �˻��׸������
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
        
        'IFORDCD�� IFSEQ��
        Case 5
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
                  
        'IFSEQ�� IFORDCD��
        Case 6
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s03
                    Exit For
                End If
            Next
        
        'IFRSTCD�� IFSEQ��
        Case 7
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s04 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s01
                    Exit For
                End If
            Next
        
        'IFSEQ�� IFRSTCD��
        Case 8
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s01 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s04
                    Exit For
                End If
            Next
        
        'IFORDCD�� IFSPECIMEN��
        Case 9
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s05
                End If
            Next
            
        'IFORDCD�� ������������
        Case 10
            For i = 1 To giOriginIFItemCnt
                If gIFItem(i).s03 = sComp Then
                    ConvertIFItemInfo = gIFItem(i).s09
                End If
            Next
        
        'IFSEQ�� IFSPCECIMEN��
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
            'NewIFList�� OldIFList�� ���� = üũ O, X
            Call .SetText(2, .MaxRows, "1")     'üũ
            '<--- ���߿� ������ �����ÿ� ����..
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
            
            '���� Row ���
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
                    'sIFSeq�� ���� ��Ȯ�� �Ҽ��ڸ� ó��, ������ ���� ����ġ ó��
                    Call .GetText(10, .MaxRows, vSex)
                    
                    Select Case vSex
                        Case "M", "��", "1"
                            vSex = "M"
                        Case "F", "��", "2"
                            vSex = "F"
                        Case Else
                            vSex = "M"
                    End Select
                    
                    '�����2(����)�� Ư���� ������ �����ϴ� �Լ�
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
            
            '���� Row ���
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
    
    'iAdd = True : ���ο� �׸��� ��� �߰�, iAdd = False : �����׸��� ��� ������
    iAdd = True
    
    With gfIFDisplayForm
        With .spdIntList
            Call .GetText(2, iCRow, vTmp)
                
            For i = 1 To iRstCnt
                Call .GetText(15, iCRow, vCRstCnt)
                Call .GetText(16, iCRow, vIFCnt)
                
                If vCRstCnt = "N" Then
                'ó�� ����� ���۵Ǿ��� ��
                    vCRstCnt = 0
                End If
                
                iCompCnt = Val(vIFCnt)
                
                '��ü sIFRstCd �� �ϳ��� ������
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                sCFlag = GetByOne(sFlag, sFlag)
                
                'IFSeq�� ��ȯ - �ߺ��Ǵ� �׸��� ���� ��������� IFSeq�� ������
                iAllCnt = 0
                
                'sCIFRstCd�� ��ġ�ϴ� ��� IFSeq ����
                For j = 1 To giOriginIFItemCnt
                    If gIFItem(j).s04 = sCIFRstCd Then
                        iAllCnt = iAllCnt + 1
                        ReDim Preserve aIFSeq(iAllCnt)
                        aIFSeq(iAllCnt) = gIFItem(j).s01
                    End If
                Next
                
                iExist = 0
                
                For j = 1 To iCompCnt
                '���� Row�� ��� IFSeq�� ���� ���� IFSeq�� ����
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
                        'ó�� ����
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
                    
                    '������ ���۹��� �˻��׸��� IFSeq�� ����
                    For j = 1 To iCompCnt
                        Call .GetText(16 + j, iCRow, vTmp)
                        sTmp = CStr(vTmp)
                        
                        '���� �˻��׸��� IFSeq�� ������
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
                            '������ Į�������� �ѱ�
                                iCCol = j
                                Exit For
                            Else
                                iAdd = True
                            End If
                        End If
                    Next
                    
                    '������ ���۵� �׸�����, �����۵� �׸������� ����
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
    
    'iAdd = True : ���ο� �׸��� ��� �߰�, iAdd = False : �����׸��� ��� ������
    iAdd = True
    
    With gfIFDisplayForm
        With .spdIntList
            Call .GetText(2, iCRow, vTmp)
                
            For i = 1 To iRstCnt
                Call .GetText(15, iCRow, vCRstCnt)
                Call .GetText(16, iCRow, vIFCnt)
                
                If vCRstCnt = "N" Then
                'ó�� ����� ���۵Ǿ��� ��
                    vCRstCnt = 0
                End If
                
                iCompCnt = Val(vIFCnt)
                
                '��ü sIFRstCd �� �ϳ��� ������
                sCIFRstCd = GetByOne(sIFRstCd, sIFRstCd)
                sCRst1 = GetByOne(sRst1, sRst1)
                sCRst2 = GetByOne(sRst2, sRst2)
                
                'IFSeq�� ��ȯ - �ߺ��Ǵ� �׸��� ���� ��������� IFSeq�� ������
                iAllCnt = 0
                
                'sCIFRstCd�� ��ġ�ϴ� ��� IFSeq ����
                For j = 1 To giOriginIFItemCnt
                    If gIFItem(j).s04 = sCIFRstCd Then
                        iAllCnt = iAllCnt + 1
                        ReDim Preserve aIFSeq(iAllCnt)
                        aIFSeq(iAllCnt) = gIFItem(j).s01
                    End If
                Next
                
                iExist = 0
                
                For j = 1 To iCompCnt
                '���� Row�� ��� IFSeq�� ���� ���� IFSeq�� ����
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
                        'ó�� ����
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
                    
                    '������ ���۹��� �˻��׸��� IFSeq�� ����
                    For j = 1 To iCompCnt
                        Call .GetText(16 + j, iCRow, vTmp)
                        sTmp = CStr(vTmp)
                        
                        '���� �˻��׸��� IFSeq�� ������
                        sPIFSeq = GetByOne(sTmp, sTmp)
                        sPRst1 = GetByOne(sTmp, sTmp)
                        sPRst2 = GetByOne(sTmp, sTmp)
                        
                        If sPIFSeq = "" Then
                            iCCol = j
                            Exit For
                        Else
                            If sPIFSeq = sCIFSeq Then
                                iAdd = False
                            '������ Į�������� �ѱ�
                                iCCol = j
                                Exit For
                            Else
                                iAdd = True
                            End If
                        End If
                    Next
                    
                    'sIFSeq�� ���� ��Ȯ�� �Ҽ��ڸ� ó��, ������ ���� ����ġ ó��
                    Call .GetText(10, iCRow, vSex)
                    
                    Select Case vSex
                        Case "M", "��", "1"
                            vSex = "M"
                        Case "F", "��", "2"
                            vSex = "F"
                        Case Else
                            vSex = "M"
                    End Select
                    
                    Call gfIFDisplayForm.SpecificProcessResult(sCIFRstCd, sCRst1, sCRst2, sCIFSeq, CStr(vSex))
                    
                    '������ ���۵� �׸�����, �����۵� �׸������� ����
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
    
    '���1, ���2 �ٲٱ�
    sRst1 = sNRst1
    sRst2 = sNRst2
    
    Exit Function
    
ErrHandler:
    OldIFListBySex = "NO"
    ViewMsg "OldIFListBySex ���� - (" & Err.Description & ")"
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
    
'���� sRst���� �Ҽ��� ������ ���� �ٲ�
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
            '�ø�
            Case "U"
                Select Case sDot
                    Case "0"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 1)
                        End If
                    Case "1"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.1)
                        End If
                    Case "2"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.01)
                        End If
                    Case "3"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.001)
                        End If
                    Case "4"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.0001)
                        End If
                    Case "5"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.00001)
                        End If
                    Case "6"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.000001)
                        End If
                    Case "7"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.0000001)
                        End If
                    Case "8"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.00000001)
                        End If
                    Case "9"
                        '�ø��� �ƴ϶� �ݴ�� ������ �� ���
                        If (Val(sTmpVal) - Val(sValue)) < 0 Then
                            sTmpVal = CStr(Val(sTmpVal) + 0.000000001)
                        End If
                End Select
                
            '�ݿø�
            Case "H"
                
            '����
            Case "L"
                Select Case sDot
                    Case "0"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 1)
                        End If
                    Case "1"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.1)
                        End If
                    Case "2"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.01)
                        End If
                    Case "3"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.001)
                        End If
                    Case "4"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.0001)
                        End If
                    Case "5"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.00001)
                        End If
                    Case "6"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.000001)
                        End If
                    Case "7"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.0000001)
                        End If
                    Case "8"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
                        If (Val(sTmpVal) - Val(sValue)) > 0 Then
                            sTmpVal = CStr(Val(sTmpVal) - 0.00000001)
                        End If
                    Case "9"
                        '������ �ƴ϶� �ݴ�� �ø��� �� ���
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
        
    'Title ����
    gfIFDisplayForm.Caption = "   " & UCase$(gsMachineNm) & " �������̽� ȭ�� - BY ACK Co., Ltd."
    
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
         
        'Rack, Pos ��뿩��
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

'Interface Mode�� ���� Display
    If gsIFMode = "0" Then
    'Uni-Direction
        With gfIFDisplayForm
            .fraSendOrd.Visible = False
            .fraBarCd.Top = 7500
        End With
    Else
    'Bi-Direction
        '1=�����(Rack Or Tray ��� ��������, But Rack/Pos ǥ��)
        '2=�����(Rack Or Tray ��� ��������, But Tray/Pos ǥ��)
        '3=�����(Rack Or Tray ��� ��������, But Tray/Cup ǥ��)
        '4=�����(Rack/Pos ��� ����),
        '5=�����(Tray/Pos ��� ����),
        '6=�����(Tray/Cup ��� ����),

        With gfIFDisplayForm
            .fraBarCd.Visible = False
            
            If gsIFMode = "1" Then
            'Rack Or Tray ��� ��������, But Rack/Pos ǥ��
                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Rack"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "2" Then
            'Rack Or Tray ��� ��������, But Tray/Pos ǥ��
                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "3" Then
            'Rack Or Tray ��� ��������, But Tray/Cup ǥ��
                .fraSendOrd.Visible = False
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Cup"))
            ElseIf gsIFMode = "4" Then
            'Rack/Pos ��� ����
                .pnlRackTray = "Rack"
                .pnlPosCup = "Pos"
                
                Call .spdIntList.SetText(6, 0, CVar("Rack"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "5" Then
            'Tray/Pos ��� ����
                .pnlRackTray = "Tray"
                .pnlPosCup = "Pos"
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Pos"))
            ElseIf gsIFMode = "6" Then
            'Tray/Cup ��� ����
                .pnlRackTray = "Tray"
                .pnlPosCup = "Cup"
                
                Call .spdIntList.SetText(6, 0, CVar("Tray"))
                Call .spdIntList.SetText(7, 0, CVar("Cup"))
            End If
        End With
    End If
    
'Transmit Mode�� ���� Display
    If gsTXMode = "0" Then
    'Batch
        '��� Option�� Client�� �ϸ� OK
    ElseIf gsTXMode = "1" Then
    'RealTime �� �׸�
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    ElseIf gsTXMode = "2" Then
    'RealTime  �� ȯ�ھ�
        With gfIFDisplayForm.spdIntList
            .Col = 2
            .ColHidden = True
        End With
    End If
    
'Initialize mode�� ���� Display
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
    
'APMode�� ���� ��� Display
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
                'Pos �ʱ�ȭ
                If .txtPos.MaxLength >= 1 And .txtPos.MaxLength <= 10 Then
                    .txtPos = Format("1", RackFormat(.txtPos.MaxLength))
                Else
                    .txtPos = "0"
                End If
                
                'Rack ����
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
                'Pos �ʱ�ȭ
                If .txtPos.MaxLength >= 1 And .txtPos.MaxLength <= 10 Then
                    .txtPos = Format("1", RackFormat(.txtPos.MaxLength))
                Else
                    .txtPos = "0"
                End If
                
                'Rack ����
                If .txtRack.MaxLength >= 1 And .txtRack.MaxLength <= 10 Then
                    .txtRack = Format(Val(.txtRack) + 1, RackFormat(.txtRack.MaxLength))
                Else
                    .txtRack = "0"
                End If
                
                Exit Sub
            End If
        End If
        
        'Pos ����
        If .txtPos.MaxLength >= 1 And .txtPos.MaxLength <= 10 Then
            .txtPos = Format(Val(.txtPos) + 1, RackFormat(.txtPos.MaxLength))
        Else
            .txtPos = "0"
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayNextRackPos ���� - (" & Err.Description & ")"
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
                    
        gfIFDisplayForm.lblCSelList = "�����ȸ : " & CStr(vJDate) & CStr(vJGbn) & CStr(vJNo)
        
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
    ViewMsg "DisplayResult2 �����߻�" & "(" & CStr(Err.Description) & ")"
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
    ViewMsg "EditRegState ���� - (" & Err.Description & ")"
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
                
            'Worklist�� ������ Match
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        gsIFMode = "0"   'Default �ܹ���
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        gsINITMode = "0"   'Default ������
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
    ViewMsg "GetIFTestItem - Local DB ���� ����!!"
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
    retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineCd", "", sBuf, 255, App.Path & "\����ڵ�.ini")
    
    If retval = 0 Then
        MsgBox "����ڵ� ������ �Ǿ� ���� �ʽ��ϴ�. ���α׷��� ����� ������ �� �����ϴ�!!", vbCritical, "����ڵ�.ini ����"
    End If
    
    gsMachineCd = Left(sBuf, retval) 'Machine Name
    
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineNm", "", sBuf, 255, App.Path & "\����ڵ�.ini")
    
    If retval = 0 Then
        MsgBox "����ڵ� ������ �Ǿ� ���� �ʽ��ϴ�. ���α׷��� ����� ������ �� �����ϴ�!!", vbCritical, "����ڵ�.ini ����"
    End If
    
    gsMachineNm = Left(sBuf, retval)
    
'Machine Exe
    sBuf = String(255, 0)
    retval = GetPrivateProfileString("InterfaceMachineCode", "InterfaceMachineExe", "", sBuf, 255, App.Path & "\����ڵ�.ini")
    
    If retval = 0 Then
        MsgBox "����ڵ� ������ �Ǿ� ���� �ʽ��ϴ�. ���α׷��� ����� ������ �� �����ϴ�!!", vbCritical, "����ڵ�.ini ����"
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
            '������
            Case "0"
                sDelFlag = ""
            '��ȭ��
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
            '��ȭ����
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
            '�Ⱓ�� ��ȭ��
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
            '�Ⱓ�� ��ȭ����
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
            '���뺯ȭ����
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
                            JudgeResultWithSex = sLimit1 & " ����"
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
                            JudgeResultWithSex = sLimit2 & " �̻�"
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
        '���� / �̻�
            If IsNumeric(sCompRst) = True Then
                If Val(sCompRst) <= Val(sRef1) Then
                    JudgeResultWithSex = "<" & sRef1
                    sOneRst2 = "����"
                ElseIf Val(sCompRst) > Val(sRef1) And Val(sCompRst) < Val(sRef2) Then
                    JudgeResultWithSex = sCompRst
                    sOneRst2 = ""
                Else
                    JudgeResultWithSex = ">" & sRef2
                    sOneRst2 = "�̻�"
                End If
            Else
                If sCompRst = "LOWER LIMIT" Then
                    If sRef1 = "" Then
                    Else
                        JudgeResultWithSex = "<" & sRef1
                        sOneRst2 = "����"
                    End If
                ElseIf sCompRst = "UPPER LIMIT" Then
                    If sRef2 = "" Then
                    Else
                        JudgeResultWithSex = ">" & sRef2
                        sOneRst2 = "�̻�"
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
        'P/N ���
        
        Case Else
        
    End Select
    
    'LIMIT���п� ���� ó��
    If sLimit1 <> "" And Val(sCompRst) < Val(sLimit1) Then
        Select Case sLimit1Gbn
            Case "0"
                JudgeResultWithSex = sLimit1
            Case "1"
                JudgeResultWithSex = "< " & sLimit1
            Case "2"
                JudgeResultWithSex = sLimit1 & " ����"
        End Select
    End If
    
    If sLimit2 <> "" And Val(sCompRst) > Val(sLimit2) Then
        Select Case sLimit2Gbn
            Case "0"
                JudgeResultWithSex = sLimit2
            Case "1"
                JudgeResultWithSex = "> " & sLimit2
            Case "2"
                JudgeResultWithSex = sLimit2 & " �̻�"
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
    
    'LIMIT���п� ���� LIMIT ó��
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
                    '1.0 ����
                    JudgeRstBySex = sLimit1 & " ����"
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
                    '1.0 �̻�
                    JudgeRstBySex = sLimit2 & " �̻�"
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
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
    
    '������ DLL�� Call�Ͽ� �����ʿ� ��������
    sBuf = gRstcfg.sComponent

    If sBuf = "" Then
        ViewMsg "������ �������� ���� DLL ������ �������� �ʽ��ϴ�!!"
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
    
    'QC�� ���� ����
    If Len(vJNo) <> 13 Or Mid(vJNo, 7, 1) = "Q" Then
        Set objRst = Nothing
        Exit Function
    End If
    
    
    sTIFSeq = ""
    sTSvrCd = ""
    sTRst1 = ""
    iTRstCnt = 0
            
    'ServerCd�� ��ȯ - �������ڵ尡 �����ϴ� �͸� ���
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
                
                sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
                
                'IFSeq�� IFRstCd�� Convert
                If Len(sIFSeq) = 2 And Left(sIFSeq, 1) = "C" Then
                '�����϶��� ������ IFSeq��
                    If sIFSeq = sCRstCd Then
                        'IFSeq�� �������ڵ�� Convert
                        sCSvrCd = ConvertIFItemInfo(2, sIFSeq)
                        Exit For
                    End If
                Else
                '�Ϲ��׸��� ���
                    If ConvertIFItemInfo(8, sIFSeq) = sCRstCd Then
                        'IFSeq�� �������ڵ�� Convert
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
    
    '������� ����
    Call objRst.SetMachineInfo(gsMachineCd, gsMachineNm)

    sRetVal = objRst.RegServer(1, Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD"), CStr(vWSeq) & Chr(124), _
                CStr(vJDate) & Chr(124), CStr(vJGbn) & Chr(124), CStr(vJNo) & Chr(124), _
                CStr(vRack) & Chr(124), CStr(vPos) & Chr(124), _
                CStr(vRegNo) & Chr(124), CStr(vPtNm) & Chr(124), CStr(vSex) & Chr(124), _
                CStr(vEmer) & Chr(124), CStr(vRerun) & Chr(124), CStr(vOther) & Chr(3), _
                CStr(iTRstCnt) & Chr(124), sTIFSeq & Chr(3), sTSvrCd & Chr(3), sTRst1 & Chr(3), sTRst2 & Chr(3), _
                ADOCN1, ADOCN2, gSvrInfo.DBGbn)
                
    If sRetVal = "OK" Then
        ViewMsg CStr(vJNo) & "�� ����� ������ �����Ͽ����ϴ�!!"
    Else
        ViewMsgLog "���� ERR : " & CStr(vJNo)
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
    
    'LastWSeq�� ����
    gsLastWSeq = gOrderTable.sWSeq
    
    Exit Sub
    
ErrHandler:
    Set objld = Nothing
    ViewMsg "RegOrder ���� - (" & Err.Description & ")"
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
        

    'iMode = 1 ---> �� ���þ� LOCAL ���
        Call .GetText(16, CInt(sCRow), vIFItemCnt)
        
        For i = 1 To CInt(vIFItemCnt)
            Call .GetText(16 + i, CInt(sCRow), vTmp)
            
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
            sRst1 = GetByOne(sTmp, sTmp)
            sRst2 = GetByOne(sTmp, sTmp)
            
            sTIFSeq = sTIFSeq & sIFSeq & "|"
            sTRst1 = sTRst1 & sRst1 & "|"
            sTRst2 = sTRst2 & sRst2 & "|"
        Next i
    End With
    
'    If Mid(sJNo, 7, 1) <> "Q" Then
    If Trim(sJNo) <> "" And Mid(sJNo, 7, 1) <> "Q" Then     '2004/6/10 yk
        '--- ������ ���α׷��� �޼��� ����(2003/3/17 yk)
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
    ViewMsg "SendResultSocket ���� - (" & Err.Description & ")"
End Sub

Private Function RegResultTemp(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
                        ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                        ByVal sIFSpcCd As String, Optional ByVal iCnt As Integer) As String
    On Error GoTo ErrHandler
    
    'iMode = 1 ---> �� ���þ� �ڵ� ���
    
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
        

    'iMode = 1 ---> �� ���þ� LOCAL ���
        Call .GetText(16, CInt(sCRow), vIFItemCnt)
        
        For i = 1 To CInt(vIFItemCnt)
            Call .GetText(16 + i, CInt(sCRow), vTmp)
            
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
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
                '--- ������ ���α׷��� �޼��� ����(2003/3/17 yk)
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
                        ViewMsg sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                    Else
                        ViewMsg sJGbn & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                    End If
                ElseIf Len(sJGbn) = 0 Then
                    If Len(sJDate) = 0 Then
                        ViewMsg sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                    Else
                        ViewMsg sJDate & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                    End If
                Else
                    ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                End If
            Else
                ViewMsg sWDate & "-" & sWSeq & "�� ���忡 �����Ͽ����ϴ�..."
            End If
        End If
    End With
    
    Set objld = Nothing
    
    Exit Function
ErrHandler:
    Set objld = Nothing
    ViewMsg "RegResultTemp ���� - (" & Err.Description & ")"
End Function


Public Function RegResult(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
                        ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                        ByVal sIFSpcCd As String, ByVal sFlag As String, Optional ByVal iCnt As Integer) As String
    
    On Error GoTo ErrHandler
    
    'iMode = 0 ---> �� �˻��׸��� ����� �ڵ� ���
    'iMode = 1 ---> �� ���þ� �ڵ� ���
    'iMode = 2 ---> Batch��Ŀ� ��� ���� ���� �� ���� ���
    
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
                
            'iMode = 0 ---> �� �˻��׸��� ����� LOCAL ���
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
                        
                        sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
                        
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
                                    ViewMsg sJNo & "�� ����� �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "�� ����� �����Ͽ����ϴ�..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "�� ����� �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "�� ����� �����Ͽ����ϴ�..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "�� ����� �����Ͽ����ϴ�..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "�� ����� �����Ͽ����ϴ�..."
                        End If
                        
                        Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, CInt(sCRow), CInt(sCRow), _
                                RGB(0, 0, 0), ���ʷ�)
                        
                    Else
                        If Len(sJNo) > 0 Then
                            If Len(sJDate) = 0 Then
                                If Len(sJGbn) = 0 Then
                                    ViewMsg sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "�� ���忡 �����Ͽ����ϴ�..."
                        End If
                    End If
                End If

            'iMode = 1 ---> �� ���þ� LOCAL ���
                If iMode = 1 Then
                    Call .GetText(16, CInt(sCRow), vIFItemCnt)
                    
                    For i = 1 To CInt(vIFItemCnt)
                        Call .GetText(16 + i, CInt(sCRow), vTmp)
                        
                        sTmp = CStr(vTmp)
                        
                        sIFSeq = GetByOne(sTmp, sTmp)  '�˻��׸��ڵ�
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
                                    ViewMsg sJNo & "�� ����� �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "�� ����� �����Ͽ����ϴ�..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "�� ����� �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "�� ����� �����Ͽ����ϴ�..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "�� ����� �����Ͽ����ϴ�..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "�� ����� �����Ͽ����ϴ�..."
                        End If
                    
                        Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, CInt(sCRow), CInt(sCRow), _
                                RGB(0, 0, 0), ���ʷ�)
                        
                    Else
                        If Len(sJNo) > 0 Then
                            If Len(sJDate) = 0 Then
                                If Len(sJGbn) = 0 Then
                                    ViewMsg sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJGbn & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                End If
                            ElseIf Len(sJGbn) = 0 Then
                                If Len(sJDate) = 0 Then
                                    ViewMsg sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                Else
                                    ViewMsg sJDate & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                                End If
                            Else
                                ViewMsg sJDate & "-" & sJGbn & "-" & sJNo & "�� ���忡 �����Ͽ����ϴ�..."
                            End If
                        Else
                            ViewMsg sWDate & "-" & sWSeq & "�� ���忡 �����Ͽ����ϴ�..."
                        End If
                    End If
                End If
                
        'iMode = 2 ---> ���� ���� Batch ���
            Case 2
        
            Case Else
        End Select
    End With
    
    Set objld = Nothing
    
    Exit Function
    
ErrHandler:
    Set objld = Nothing
    ViewMsg "RegResult ���� - (" & Err.Description & ")"
End Function


Public Sub RegViewMsgHwnd(ByVal lnHwnd As Long)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "ViewMsg.Hwnd", CStr(lnHwnd))
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
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
        If vOption = 1 Then     '�� ����
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
        
        If vOption = 2 Then     '�ϴ� �迭 ����
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
        
        If vOption = 3 Then     '��� �迭 ����
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
                    '����׸��� �� : sIFRstCd = sIFSeq
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
                    '�Ϲ��׸��� ��
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
        
'        'Result spdRst�� ǥ��
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
    'Interface �׸� Seq�� ��ġ�ϴ� �˻�� �Ѹ���
        For j = 1 To giOriginIFItemCnt
            If Format$(i, "000") = gIFItem(j).s01 Then
                iCurItemCnt = iCurItemCnt + 1
                Call gfIFDisplayForm.spdIntList.SetText(16 + giTotIFItemCnt + iCurItemCnt, 0, gIFItem(j).s02 & "")
                
                Exit For
            End If
        Next
    Next
    
    For i = 1 To MAXCALITEM
    '����׸�� ��ġ�ϴ� �˻�� �Ѹ���
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
                    ViewMsgLog "��ġ ERR : Rack(Tray) is empty!!"
                    Exit Sub
                Else
                    If gIFPosInfo(j).sRackNo = CStr(vRack) Then
                        sEachRackPos = gIFPosInfo(j).sPosMaxNo
                        Exit For
                    End If
                End If
            Next
            
            If vPos = "" Then
                ViewMsgLog "��ġ ERR : Pos(Cup) is empty!!"
                Exit Sub
            End If
            
            If Val(vPos) > Val(sEachRackPos) Then
                ViewMsgLog "��ġ ERR : Pos(Cup) is over!!"
                Exit Sub
            End If
            
            'Alphabet
            If gIFRack.sRackDigit = 1 And Val(gIFRack.sMaxRack) = 26 Then
                If Val(vPos) = Val(sEachRackPos) Then
                    'Pos �ʱ�ȭ
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format("1", RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack ����
                    If IsNumeric(vRack) = False Then
                        Call .SetText(6, i, CVar(Chr(Asc(CStr(vRack)) + 1) & ""))
                    End If
                Else
                    'Pos ����
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format(Val(vPos) + 1, RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack �״��
                    If IsNumeric(vRack) = False Then
                        Call .SetText(6, i, CVar(Chr(Asc(CStr(vRack))) & ""))
                    End If
                End If
            '1, 01, 001, 0001, ....
            Else
                If Val(vPos) = Val(sEachRackPos) Then
                    'Pos �ʱ�ȭ
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format("1", RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack ����
                    If IsNumeric(vRack) = True Then
                        Call .SetText(6, i, CVar(Format(Val(vRack) + 1, RackFormat(gIFRack.sRackDigit)) & ""))
                    End If
                Else
                    'Pos ����
                    If gIFRack.sPosDigit >= 1 And gIFRack.sPosDigit <= 10 Then
                        Call .SetText(7, i, CVar(Format(Val(vPos) + 1, RackFormat(gIFRack.sPosDigit)) & ""))
                    Else
                        Call .SetText(7, i, CVar("0"))
                    End If
                    
                    'Rack �״��
                    If IsNumeric(vRack) = True Then
                        Call .SetText(6, i, vRack)
                    End If
                End If
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayRackPos ���� - (" & Err.Description & ")"
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
        '�۾����ڸ� ����
        gOrderTable.sWDate = Format(gfIFDisplayForm.dtpLabDate.Value, "YYYYMMDD")
        '�۾��Ϸù�ȣ�� ����
        gOrderTable.sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
        
        '�ش���ڵ��� ���������� �ѱ�
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
        
        '�˻��׸� ���� �����
        For i = 1 To gOrderTable.iOrdCnt
            Call .SetText(16 + i, gOrderTable.iCRow, gOrderTable.sIFSeq(i) & "||||")
        Next i
        
        If sState <> "DISPLAY" Then
            Call SpdForeBack(gfIFDisplayForm.spdIntList, 3, 15, gOrderTable.iCRow, gOrderTable.iCRow, _
                    RGB(0, 0, 0), �����)
        
            gfIFDisplayForm.lblOrder = gOrderTable.sJNo
        End If
    End With
    
    'Order ���� Local MDB�� Insert
    Call RegOrder(1)
    
    'gOrderTable �ʱ�ȭ
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
    
    'gOrderTable �ʱ�ȭ
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
            Case 0  'JDate, JGbn, JNo�� �ѱ�� ���
                .lblResult = sJDate & "-" & sJGbn & "-" & sJNo
                
                iCRow = FindIFListWithJ(sJDate, sJGbn, sJNo)
                
                If iCRow > 0 Then
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag) = "NO" Then
                           
                        Exit Sub
                    End If
                    
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                    
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case 1  'WSeq�� �ѱ�� ���
                .lblResult = sWDate & "-" & sWSeq
                
                iCRow = FindIFListWithW(sWSeq)
                
                If iCRow > 0 Then
                    If OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag) = "NO" Then
                           
                        Exit Sub
                    End If
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                        
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                                                
            Case 2  'CurRow�� �ѱ�� ��� - ��) �Һ���ⰰ�� �ܹ��� ���
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
            
            '����Ʈ ���� �����ϴ� ���� �ܹ����� ���
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                        
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
            
            Case 3  'JNo�� �ѱ�� ���
                .lblResult = sJNo
                
                iCRow = FindIFListWithJNo(sJNo)
                
                If iCRow > 0 Then
                    Call OldIFList(iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sRack, sPos, _
                            sRegNo, sName, sSex, sEmer, sReRun, sOther, sFlag)
                    
                Else
                    If .chkOExist = "1" Then
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        '�ϴ� Order�� ������ �Ѹ��� ����� ��Ÿ��
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
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case 4  '�۾�����Ʈ�� ������Ī
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
                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
                        giAddKey = 1
                    
                        sCWSeq = NewIFList(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
                                    sRack, sPos, sRegNo, sName, _
                                    sSex, sEmer, sReRun, sOther, _
                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
                                    sIFSpcCd, sCurRow, sFlag)
                    Else
                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
                        .lblResult = "No List!!"
                        Exit Sub
                    End If
                End If
                
            Case Else
            
        End Select
        
        '������ ���Ե� �׸��� �����Ͽ� ��Ÿ��
        sChkVal = ChkCalResult(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        
        'Low, High ���� �����Ͽ� ������ ��Ÿ��
        sRetVal = ViewIFResult2(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        
        With .spdIntList
            Call .RowHeightToTwips(1, .RowHeight(1), lngTwipHeight)
            iRowCnt = Format((.Height / lngTwipHeight) - 2, "0")
            
'            If .MaxRows > iRowCnt Then
'                .TopRow = .MaxRows - iRowCnt + 1
'            End If
        End With
        
    'gsTxMode="0" => Batch, gsTxMode="1" => RealTime(�� �׸�), gsTxMode="2" => RealTime(�� ȯ�ھ�)
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
                         gResultTable(1).iCRow, RGB(0, 0, 0), ���ʷ�)
                End If
            End If
        ElseIf gsTXMode = "2" Then
        '���ϴ� �����Ϲ�Ĵ�� ���� ������.
            If sRetVal = "NONE" Then
            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
                'ȯ�ڴ����� ��� ��� �� ���
                If giAddKey = 1 Then
                    If sCWSeq = "" Then
                    Else
                        gsLastWSeq = sCWSeq
                    End If
                End If
                
                Call RegResult(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd, sFlag)
                
                Call SpdForeBack(.spdIntList, 3, 15, gResultTable(1).iCRow, _
                                gResultTable(1).iCRow, RGB(0, 0, 0), ���ʷ�)
            End If
        End If
        
        If gfIFDisplayForm.optRegOpt(0).Value = True Then
'            '���� �ڵ������ �ӽ����̺� ���
'            Call RegResultTemp(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
            'WinSock���� ������ȭ�鿡 ������� ����
            Call SendResultSocket(1, CStr(gResultTable(1).iCRow), iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
        End If
        
        Erase gResultTable
    End With
    
    Exit Sub
    
ErrHandler:
    ViewMsg "DisplayResultOK �����߻� - ( " & Err.Description & " )"
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
'            Case 0  'JDate, JGbn, JNo�� �ѱ�� ���
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
'                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 1  'WSeq�� �ѱ�� ���
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
'                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 2  'CurRow�� �ѱ�� ��� - ��) �Һ���ⰰ�� �ܹ��� ���
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
'            '����Ʈ ���� �����ϴ� ���� �ܹ����� ���
'                Else
'                    If .chkOExist = "1" Then
'                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 3  'JNo�� �ѱ�� ���
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
'                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
'                        '�ϴ� Order�� ������ �Ѹ��� ����� ��Ÿ��
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
'                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case 4  '�۾�����Ʈ�� ������Ī
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
'                    '����Ʈ�� ��� ����ޱ⸦ üũ�� ���
'                        giAddKey = 1
'
'                        sCWSeq = NewIFListBySex(sWDate, sWSeq, sJDate, sJGbn, sJNo, _
'                                    sRack, sPos, sRegNo, sName, _
'                                    sSex, sEmer, sReRun, sOther, _
'                                    iRstCnt, sIFRstCd, sRst1, sRst2, _
'                                    sIFSpcCd, sCurRow)
'                    Else
'                    '����Ʈ�� ��� ����ޱ⸦ üũ���� ���� ���
'                        .lblResult = "No List!!"
'                        Exit Sub
'                    End If
'                End If
'
'            Case Else
'
'        End Select
'
'        '������ ���Ե� �׸��� �����Ͽ� ��Ÿ��
'        sChkVal = ChkCalResult(gResultTable(1).iCRow, iRstCnt, sIFRstCd, sRst1, sRst2, sIFSpcCd)
'
'        'Low, High ���� �����Ͽ� ������ ��Ÿ��
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
'    'gsTxMode="0" => Batch, gsTxMode="1" => RealTime(�� �׸�), gsTxMode="2" => RealTime(�� ȯ�ھ�)
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
'                         gResultTable(1).iCRow, RGB(0, 0, 0), ���ʷ�)
'                End If
'            End If
'        ElseIf gsTXMode = "2" Then
'        '���ϴ� �����Ϲ�Ĵ�� ���� ������.
'            If sRetVal = "NONE" Then
'            ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
'                'ȯ�ڴ����� ��� ��� �� ���
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
'                         gResultTable(1).iCRow, RGB(0, 0, 0), ���ʷ�)
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
'    ViewMsg "DisplayResultOkBySex �����߻� - ( " & Err.Description & " )"
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
                MsgBox "ȯ�漳������ Result Setting�� �����Ͻʽÿ�!!", vbCritical
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
