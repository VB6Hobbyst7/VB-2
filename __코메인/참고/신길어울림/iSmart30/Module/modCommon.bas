Attribute VB_Name = "modCommon"
Option Explicit

'���� ������ üũ
Public Function GetFileSize(szFileName) As Long
    Dim fs, F
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set F = fs.GetFile(szFileName)
        GetFileSize = F.Size
        Set F = Nothing
        Set fs = Nothing
End Function

'�ý�Ʈ�ڽ�  ��弱��
Public Sub TextBoxs_GotFocus(ByVal Obj As TextBox)
    With Obj
        .SelStart = 0
        .SelLength = Len(Obj.Text)
    End With
End Sub

'List View�� Head�����ϱ�
Public Sub InitLvwHead(ByRef objLvw As Listview, ByVal strHead As String, ByVal strSize As String)
    Dim ii As Integer
    Dim aryTitle() As String
    Dim aryWidth() As String
    
    aryTitle = Split(strHead, ",")
    aryWidth = Split(strSize, ",")
    
    If UBound(aryWidth) < UBound(aryTitle) Then
        ReDim Preserve aryWidth(UBound(aryTitle))
    End If
    
    With objLvw
        .ColumnHeaders.Clear
        For ii = 0 To UBound(aryTitle)
            If aryWidth(ii) = "" Then
                aryWidth(ii) = "0"
            End If
            .ColumnHeaders.Add ii + 1, aryTitle(ii), aryTitle(ii), _
                (.Width \ (UBound(aryTitle) + 1)) + Val(aryWidth(ii)), vbLeftJustify
        Next ii
        .View = lvwReport
    End With
End Sub

'List View�� Head�����ϱ�
Public Sub InitLvwHeader(ByRef objLvw As Listview, objHead As Dictionary, Optional ScrollWidth As Boolean = True)
    Dim defWidth    As Long
    Dim lvwWidth    As Long
    Dim aryTitle()  As String
    Dim i As Integer
    Dim ObjX As Variant
    
    With objLvw
        .ColumnHeaders.Clear
        If ScrollWidth Then
            lvwWidth = .Width - 310
        Else
            lvwWidth = .Width - 60
        End If
        defWidth = lvwWidth \ objHead.Count
        i = 1
        For Each ObjX In objHead
            aryTitle = Split(objHead(ObjX), vbTab)
            If UBound(aryTitle) < 1 Then
                .ColumnHeaders.Add i, , aryTitle(0), defWidth, vbLeftJustify
            Else
                .ColumnHeaders.Add i, , aryTitle(1), (aryTitle(0) / 100) * lvwWidth, vbLeftJustify
            End If
            i = i + 1
        Next
        .View = lvwReport
    End With
End Sub

'List View�� ������ �ֱ�
Public Sub DataLoadLvw(ByRef objLvw As Listview, ByVal RowDel As String, _
    ByVal ColDel As String, ByVal strdata As String, Optional strTag As String)
    
    Dim itmX As ListItem
    Dim strTmp As String
    Dim aryTmp() As String
    Dim aryTag() As String
    Dim ii As Long
    Dim jj As Long
    Dim intCol As Long
   
    aryTmp = Split(P(strdata, RowDel, 1), ColDel)
    If IsMissing(strTag) Then
        strTag = ""
    End If
    
    aryTag = Split(strTag, RowDel)
    
    intCol = UBound(aryTmp) + 1
    '
    aryTmp = Split(strdata, RowDel)
    If UBound(aryTmp) > UBound(aryTag) Then
        ReDim Preserve aryTag(UBound(aryTmp))
    End If
    
    If (UBound(aryTmp) + 1) < 1 Then Exit Sub
    
    For ii = 0 To UBound(aryTmp)
        For jj = 1 To intCol
            If jj = 1 Then
                Set itmX = objLvw.ListItems.Add(, , P(aryTmp(ii), ColDel, jj))
            Else
                If P(aryTmp(ii), ColDel, jj) <> "" Then
                    itmX.SubItems(jj - 1) = P(aryTmp(ii), ColDel, jj)
                Else
                    itmX.SubItems(jj - 1) = " "
                End If
            End If
            itmX.Tag = aryTag(ii)
        Next jj
    Next ii
    Set itmX = Nothing
    '
End Sub

'�̹��� �޺� ����Ÿ �ֱ�
Public Sub DataLoadImgCombo(ByRef ImgCombo As ImageCombo, _
                            ByVal RsData As ADODB.Recordset, _
                            ByVal KeyField As Integer, _
                            ByVal DataField As Integer, _
                            Optional ImgList As ImageList)
    Dim CboiX As ComboItem
    
    With ImgCombo
        .ComboItems.Clear
        If RsData.EOF Then
            Set CboiX = .ComboItems.Add(, , "Data Nothing")
            Exit Sub
        End If
        Do Until RsData.EOF
            Set CboiX = .ComboItems.Add(, RsData.Fields(KeyField), RsData.Fields(DataField))
            RsData.MoveNext
        Loop
    End With

End Sub

'����Ʈ �� ����
Public Sub SetListView_Sort(ByVal Listview As Listview, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   With Listview
      .SortKey = ColumnHeader.Index - 1
      .Sorted = True
      If .SortOrder = lvwAscending Then
         .SortOrder = lvwDescending
      Else
         .SortOrder = lvwAscending
      End If
    .Sorted = False
   End With
End Sub

Public Function P(ByVal Var As String, ByVal Del As String, ByVal FromCnt As Long, _
    Optional ByVal optArg1, Optional ByVal optArg2) As String
    
    Dim ChkArg As Integer
    
    ChkArg = IsMissing(optArg1) + IsMissing(optArg2)
    Select Case ChkArg
        Case -2     ' single piece
            P = SinglePiece(Var, Del, FromCnt)
        Case -1
            If TypeName(optArg1) = "Integer" Or TypeName(optArg1) = "Long" Then ' multi piece
                P = MultiPiece(Var, Del, FromCnt, optArg1)
            Else    ' single piece set
                P = SinglePieceSet(Var, Del, FromCnt, optArg1)
            End If
        Case 0      ' multi piece set
            P = MultiPieceSet(Var, Del, FromCnt, optArg1, optArg2)
    End Select
End Function

Private Function SinglePiece(ByVal Var As String, ByVal Del As String, ByVal Cnt As Long) As String
    Dim Prt As Long
    Dim Srt As Long
    Dim Nxt As Long
    
    If Cnt <= 0 Then
        SinglePiece = ""
        Exit Function
    End If
    
    Nxt = (Len(Del) * -1) + 1
    For Prt = 1 To Cnt
        Srt = Nxt + Len(Del)
        Nxt = InStr(Srt, Var, Del)
        
        If Nxt = 0 Then
            Nxt = Len(Var) + Len(Del)
            Exit For
        End If
    Next Prt
    
    If Prt >= Cnt Then SinglePiece = Mid(Var, Srt, (Nxt - Srt))
End Function

Private Function MultiPiece(ByVal Var As String, ByVal Del As String, _
    ByVal FromCnt As Long, ByVal ToCnt As Long) As String
    
    Dim Prt As Long
    Dim Srt As Long
    Dim Nxt As Long
    Dim FromBuf As Long
    
    If FromCnt > ToCnt Then
        MultiPiece = ""
        Exit Function
    End If
    
    If FromCnt < 1 Then FromCnt = 1
    
    Nxt = (Len(Del) * -1) + 1
    
    For Prt = 1 To ToCnt
        Srt = Nxt + Len(Del)
        Nxt = InStr(Srt, Var, Del)
        
        If Prt = FromCnt Then FromBuf = Srt
        
        If Nxt = 0 Then
            Nxt = Len(Var) + Len(Del)
            Exit For
        End If
    Next Prt
    
    If FromBuf = 0 Then
        MultiPiece = ""
        Exit Function
    End If
    
    MultiPiece = Mid(Var, FromBuf, (Nxt - FromBuf))
End Function

Private Function SinglePieceSet(ByVal Var As String, ByVal Del As String, _
    ByVal Cnt As Long, ByVal XCH As String) As String
    
    Dim Prt As Long
    Dim Srt As Long
    Dim Nxt As Long
    
    If Cnt = 0 Then
        SinglePieceSet = ""
        Exit Function
    End If
    
    Nxt = (Len(Del) * -1) + 1
    
    For Prt = 1 To Cnt
        Srt = Nxt + Len(Del)
        Nxt = InStr(Srt, Var, Del)
        
        If Nxt = 0 Then
            Nxt = Len(Var) + Len(Del)
            Exit For
        End If
    Next Prt
    
    If Prt >= Cnt Then
        SinglePieceSet = left(Var, Srt - 1) + XCH + Mid(Var, Nxt, Len(Var) - Nxt + Len(Del))
    Else
        For Srt = 1 To Cnt - Prt
            Var = Var + Del
        Next Srt
        
        SinglePieceSet = Var + XCH
    End If
End Function

Private Function MultiPieceSet(ByVal Var As String, ByVal Del As String, _
    ByVal FromCnt As Long, ByVal ToCnt As Long, ByVal XCH As String) As String
    
    Dim Prt As Long
    Dim Srt As Long
    Dim Nxt As Long
    Dim FromBuf As Long
    
    If FromCnt > ToCnt Then
        MultiPieceSet = ""
        Exit Function
    End If
    
    If FromCnt < 1 Then FromCnt = 1
    
    If Del = "" Then
        MultiPieceSet = left(Var, FromCnt - 1) + XCH + Mid(Var, ToCnt + 1, Len(Var))
        Exit Function
    End If
    
    Nxt = (Len(Del) * -1) + 1
    
    For Prt = 1 To ToCnt
        Srt = Nxt + Len(Del)
        Nxt = InStr(Srt, Var, Del)
        
        If Prt = FromCnt Then FromBuf = Srt
        
        If Nxt = 0 Then
            Nxt = Len(Var) + Len(Del)
            Exit For
        End If
    Next Prt
    
    If FromBuf > 0 Then
        MultiPieceSet = left(Var, FromBuf - 1) + XCH + Mid(Var, Nxt, Len(Var) - Nxt + Len(Del))
    Else
        For Srt = 1 To FromCnt - Prt
            Var = Var + Del
        Next Srt
        
        MultiPieceSet = Var + XCH
    End If
End Function

Public Function L(ByVal Var As String, ByVal Del As String) As Long
    Dim Srt As Long
    Dim Nxt As Long
    Dim Cnt As Long
    
    If Del = "" Then
        L = 0
        Exit Function
    End If
    
    Nxt = (Len(Del) * -1) + 1
    
    Do
        Srt = Nxt + Len(Del)
        Nxt = InStr(Srt, Var, Del)
        Cnt = Cnt + 1
    Loop Until Nxt = 0
    
    L = Cnt
End Function

'���ڿ��� byte�� �ǵ��� �ش�.
Function LengthByte(ByVal Var As String) As Long
    Dim Cnt As Long
    Dim num As Long
    Dim TMP As String
    
    Cnt = 0: num = 0
    If Var = "" Then Exit Function
    Do
        Cnt = Cnt + 1: TMP = Mid(Var, Cnt, 1): num = num + 1
        If Asc(TMP) < 0 Then num = num + 1
    Loop Until Cnt >= Len(Var)
    LengthByte = num
End Function

Public Function HExtract(ByVal Var As String, ByVal Del As String, ByVal GetCnt As Long) As String
    Dim BUF As String
    Dim TMP As String
    Dim num As Long
    Dim Cnt As Long
    
    BUF = ""
    Cnt = 0
    num = 0
    
    If Var = "" Or GetCnt < 2 Then
        HExtract = ""
        Exit Function
    End If
    
    Do
        Cnt = Cnt + 1
        TMP = Mid(Var, Cnt, 1)
        num = num + 1
        
        If Asc(TMP) < 0 Then num = num + 1
        
        If num < GetCnt Then
            BUF = BUF + TMP
        ElseIf num = GetCnt Then
            num = 0
            BUF = BUF + TMP + Del
        ElseIf num > GetCnt Then
            num = 2
            BUF = BUF + Del + TMP
        End If
    Loop Until Cnt >= Len(Var)
    
    If Right(BUF, 1) = Del Then BUF = left(BUF, Len(BUF) - 1)
    
    HExtract = BUF
End Function

Public Function DctToStr(ByRef dctTmp As Scripting.Dictionary) As String
    Dim varKey As Variant
    Dim aryTmp() As String
    Dim blnFirst As Boolean
    
    If dctTmp.Count = 0 Then Exit Function
    For Each varKey In dctTmp.Keys
        If blnFirst = False Then
            ReDim aryTmp(0)
            blnFirst = True
        Else
            ReDim Preserve aryTmp(UBound(aryTmp) + 1)
        End If
        aryTmp(UBound(aryTmp)) = varKey & vbTab & dctTmp.Item(varKey)
    Next
    '
    DctToStr = Join(aryTmp, vbNewLine)
    '
End Function

''spread sheet sort
'Sub SpreadSheetSort(ByRef Spread As vaSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
'    Dim intCount As Integer
'    Dim strDataField As String
'    'SortType
'    ' 0 : none
'    ' 1 : ascending
'    ' 2 : descending
'
'    With Spread
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 1: .Row2 = .DataRowCnt
'        .SortBy = 0
'        .SortKey(1) = Col       '����Ű ����ȣ
'
'        If SortType = 0 Then
'            .SortKeyOrder(1) = SortKeyOrderNone
'        ElseIf SortType = 1 Then
'            .SortKeyOrder(1) = SortKeyOrderAscending
'        ElseIf SortType = 2 Then
'            .SortKeyOrder(1) = SortKeyOrderDescending
'        Else
'            .SortKeyOrder(1) = SortKeyOrderAscending
'        End If
'
'        .Action = ActionSort
'    End With
'End Sub
'
'Public Sub SpreadPrint(ByRef Spread As vaSpread, ByVal Header As String, _
'    Optional ByVal ListHead As String = "", Optional ByVal Footer As String = "")
'
'    Dim strHead1 As String
'    Dim strHead2 As String
'    Dim strHead3 As String
'    Dim strHead4 As String
'    Dim strFoot As String
'    Dim strFont1 As String
'    Dim strFont2 As String
'    Dim strFont3 As String
'
'    Spread.FontName = "����ü"
'
'    strHead1 = "/n/c/f1" & Header
'
'    If ListHead <> "" Then
'        strHead2 = "/n/l" & Space(2) & ListHead
'    End If
'
'    strHead3 = "/n/l" & Space(2) & "��� �Ͻ�: " & Replace(Format(DbSysDate, "yyyy/mm/dd hh:mm"), "-", "/")
'    strHead4 = "/r" & "������: /p" & "//" & "/p        "
'    strFoot = "/n/c/f1" & Footer
'    strFont1 = "/c/fn""����ü"" /fz""18"" /fb1 /fi0 /fu0 /fk0 /fs1"
'    strFont2 = "/fn""����ü"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
'    strFont3 = "/c/fn""����ü"" /fz""12"" /fb1 /fi0 /fu0 /fk0 /fs1"
'
'    With Spread
'        'Print Body
'        If ListHead = "" Then
'            .PrintHeader = strFont1 + strHead1 + "/n" + strFont2 + strHead3 + strHead4 + "/n"
'        Else
'            .PrintHeader = strFont1 + strHead1 + "/n" + strFont2 + strHead2 + "/n" + strFont2 + strHead3 + strHead4 + "/n"
'        End If
'
'        .PrintFooter = strFont3 + strFoot
'        .PrintMarginLeft = 150
'        .PrintMarginRight = 0
'        .PrintMarginTop = 200
'        .PrintMarginBottom = 200
'        .PrintColHeaders = True
'        .PrintRowHeaders = True
'        .PrintBorder = True
'        .PrintColor = False
'        .PrintGrid = True
'        .PrintShadows = True
'        .PrintUseDataMax = True
'        .PrintType = PrintTypeAll
'        .Action = ActionPrint
'    End With
'End Sub

'������Ϸ� ���� ��� ..............................................
'strBirthDate: �������(yyyymmdd)
'strType:���̸� ��,��,�� �� ��� �������� ������ ������(Y, M, D)
'strSysDate : ����� ������ �Ǵ� ��¥(yyyymmdd)
'             strSysDate�� Optional ������ ������ ���ڷ� ���̸� ���
'ReturnValue : ���� ����(Year����)
'...................................................................
Function FindAge(ByVal strBirthDate As String, ByVal strAgeType As String, _
    Optional ByVal strSysDate) As String
    
    Dim strFormatBirth As String
    Dim strFormatSys As String

    strFormatBirth = Format(Format(strBirthDate, "####/##/##"), "yyyy-mm-dd")
    
    If IsMissing(strSysDate) Then
'        strFormatSys = Format(DbSysDate, "yyyy-mm-dd")
        strFormatSys = Format(Now, "yyyy-mm-dd")
    Else
        strFormatSys = Format(Format(strSysDate, "####/##/##"), "yyyy-mm-dd")
    End If
    
    Select Case UCase(strAgeType)
        Case "Y"        '���
            FindAge = DateDiff("yyyy", strFormatBirth, strFormatSys)
        Case "M"        '����
            FindAge = DateDiff("m", strFormatBirth, strFormatSys)
        Case "D"        '�Ϸ�
            FindAge = DateDiff("d", strFormatBirth, strFormatSys)
    End Select
End Function

'�ش� Data�� ��¥�μ� �� Data�� ��ȿ���� Check
'strDate : Check�ϰ��� �ϴ� Data, yyyymmdd(8�ڸ�) ���ĸ� ����
Public Function DateChk(ByVal strDate As String) As Boolean
    DateChk = IsDate(Format(Format(strDate, "####-##-##"), "yyyy-mm-dd"))
End Function

'�Ϸ����� ���̸� �޾Ƽ� ���ϴ� ���̷� �ǵ��� �ش�.
Public Function ConvertAge(ByVal strDayAge As String, Optional ByVal AgeType As String = "Y", _
    Optional ByVal AgeUnit As Boolean = True) As String
    'AgeType - "Y" : ���
    '          "M" : ����
    '          "D" : �Ϸ�
    
    'AgeUnit - "True"  : ������ ������ ����.
    '          "False" : ������ ������ ������ ����.
    
    Select Case AgeType
        Case "Y"
            ConvertAge = Val(strDayAge) / 365
            If AgeUnit = True Then
                ConvertAge = ConvertAge & "��"
            End If
        Case "M"
            ConvertAge = Val(strDayAge) / 30
            If AgeUnit = True Then
                ConvertAge = ConvertAge & "����"
            End If
        Case "D"
            ConvertAge = strDayAge
            If AgeUnit = True Then
                ConvertAge = "��"
            End If
        Case Else
            ConvertAge = Val(strDayAge) / 365
            If AgeUnit = True Then
                ConvertAge = ConvertAge & "��"
            End If
    End Select
End Function

'���ڸ� �޾Ƽ� ���̸� ��� �ؼ� �����ش�.
Public Function GetAge(ByVal strBirthDate As String, Optional ByVal strSysDate) As String
    If Not DateChk(strBirthDate) Then
        GetAge = ""
        Exit Function
    End If
    
    Select Case Val(FindAge(strBirthDate, "D", strSysDate))
        Case Is < 30
            GetAge = FindAge(strBirthDate, "D", strSysDate) & "D"
        Case 31 To 365
            GetAge = FindAge(strBirthDate, "M", strSysDate) & "M"
        Case 366 To 730
            GetAge = Val(FindAge(strBirthDate, "D", strSysDate)) \ 365 & "Y" & (Val(FindAge(strBirthDate, "D", strSysDate)) Mod 365) \ 30 & "M"
        Case Is > 730
            GetAge = FindAge(strBirthDate, "Y", strSysDate) & "Y"
        Case Else
            GetAge = "0"
    End Select
    
End Function

'===========================================================================  �迭 ���� (Qucik Sort)
'in_array() As String       : ���� �ϰ��� �ϴ� �迭
'in_Left As Long            : ��Ʈ�Ϸ����ϴ� �迭�� ÷���� ���� ��ġ ��κ� 0
'in_Right As Long           : ��Ʈ�ϰ����ϴ� �迭�� ÷���� ����ġ ��κ� Ubound ��
'start_Position As Long     : ��Ʈ�ϰ����ϴ� ���ڿ��� ��ġ(�� : 'ABCDEF' C ���� ���ں��� ������ ��� 3)

Public Sub Qsort(in_array() As String, in_Left As Long, in_Right As Long, start_Position As Long)
   Dim in_Current As Long
   Dim in_Last As Long
   
   If in_Left >= in_Right Then Exit Sub
   
   Call arrSwap(in_array, in_Left, (in_Left + in_Right) / 2)
   in_Last = in_Left
   
   For in_Current = in_Left + 1 To in_Right
       If Mid(in_array(in_Current), start_Position) < Mid(in_array(in_Left), start_Position) Then
           in_Last = in_Last + 1
           Call arrSwap(in_array, in_Last, in_Current)
       End If
   Next in_Current
   
   Call arrSwap(in_array, in_Left, in_Last)
   Call Qsort(in_array, in_Left, in_Last - 1, start_Position)
   Call Qsort(in_array, in_Last + 1, in_Right, start_Position)
End Sub

Public Sub arrSwap(in_array() As String, i As Long, j As Long)
   Dim temp As String
   
   temp = in_array(i)
   in_array(i) = in_array(j)
   in_array(j) = temp

End Sub
'-----------------------------------------------------------------------------------------------------
'................................................
'Date       : 25-01-2001
'Functie     : String crypter
'Description: Call this function to crypt and call
'    it also to decrypt a string
'Url: http://vb.netmenu.nl
'.................................................
Public Function CryptString(ptSource As String, _
    ptPassword As String) As String

   Dim tdest As String
   Dim lteller As Long
   Dim lPasswTeller As Long

   tdest = ptSource
   For lteller = 1 To Len(ptSource)
      lPasswTeller = lPasswTeller - 1
      If lPasswTeller < 1 Then lPasswTeller = Len(ptPassword)

      Mid$(tdest, lteller, 1) = _
          Chr$(Asc(Mid$(ptSource, lteller, 1)) Xor _
              Asc(Mid$(ptPassword, lPasswTeller, 1)))
   Next lteller
   CryptString = tdest

End Function
