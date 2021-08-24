Attribute VB_Name = "mod공용_CTLS"
Option Explicit

'한글 편집 관련
Public Const IME_HANGUL = &H1
Public Const IME_ENGLISH = &H0
Public Const IME_NONE = &H0

Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..

Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

'외부 프로그램 실행
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
     ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Function CHG_BYTE_CUT(ByVal str As String, start, Length)
   CHG_BYTE_CUT = StrConv(MidB(StrConv(str, vbFromUnicode), start, Length), vbUnicode)
End Function

Public Sub CHG_ENG(Src As Object)
   Dim hIME As Long
   hIME = ImmGetContext(Src.hwnd)
   ImmSetConversionStatus hIME, IME_ENGLISH, IME_NONE
End Sub

Public Sub CHG_KOR(Src As Object)
   Dim hIME As Long
   hIME = ImmGetContext(Src.hwnd)
   ImmSetConversionStatus hIME, IME_HANGUL, IME_NONE
End Sub

Public Function CHK_QUOTATION(N_Codename As String) As String
    Dim i           As Long
    Dim C_Pos       As Long
    
    i = 1
    Do
        C_Pos = InStr(i, N_Codename, "'")
        If C_Pos Then
            N_Codename = Mid(N_Codename, 1, C_Pos) & "'" & Mid(N_Codename, C_Pos + 1)
            i = C_Pos + 2
        End If
    Loop While C_Pos
    CHK_QUOTATION = N_Codename
End Function

Public Sub EXCEL_EXCHANGE_ADO_RECORDSET(Rs As ADODB.Recordset, xlsFileName As String)
    Dim iRow       As Integer
    Dim iCol       As Integer
    Dim objExcel   As Object     '생성될 엑셀 개체
    Dim xlBook     As Object
    Dim xlSheet    As Object
    Dim sMSG       As String
    Dim gsMSG      As String    '일반 Message
    
    Rs.MoveLast
    Rs.MoveFirst
    
    '개체 생
    Set objExcel = CreateObject("Excel.Application")
    Set xlBook = objExcel.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    With xlSheet
        '열제목 나타내고 굵게 지정
        For iCol = 0 To Rs.Fields.Count - 1
            .Cells(1, iCol + 1).Value = Trim(Rs.Fields(iCol).Name)
        Next iCol
        
        iCol = 0: iRow = 0
        .Range("A1").CurrentRegion.Font.Bold = True
        
        '레코드 Excel로 변환
        For iRow = 0 To Rs.RecordCount - 1
            For iCol = 0 To Rs.Fields.Count - 1

'''                '>>> Cell
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "0.00%"
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "dd/mm/yy"
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "#####0"
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "##0.0000"
'''
'''                '>>> Range
'''                .Workbooks.Application.Worksheets(1).Range("A:A").NumberFormat = "#####0"
'''                .Workbooks.Application.Worksheets(1).Range("B:B").NumberFormat = "######0"
'''                .Workbooks.Application.Worksheets(1).Range("H:H").NumberFormat = "##0.0000"
'''                .Workbooks.Application.Worksheets(1).Range("J:J").NumberFormat = "mm/dd/yyyy"

                '/Replace(Trim(Rs.Fields(iCol).Value & ""), vbCrLf, vbLf)     => 줄바꿈 문자열일때 음표가 엑셀에 나타나는 것을 없애줌.
                .Cells(iRow + 2, iCol + 1).Value = Replace(Trim(Rs.Fields(iCol).Value & ""), vbCrLf, vbLf)
            Next iCol
            Rs.MoveNext
        Next iRow
        
    End With
    
    '지정한 경로로 저장
    xlSheet.SaveAs xlsFileName
    
    objExcel.Quit
    Set objExcel = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
End Sub

Public Sub EXCEL_EXCHANGE_DAO_RECORDSET(Rs As Recordset, xlsFileName As String)
    Dim iRow       As Integer
    Dim iCol       As Integer
    Dim objExcel   As Object     '생성될 엑셀 개체
    Dim xlBook     As Object
    Dim xlSheet    As Object
    Dim sMSG       As String
    Dim gsMSG      As String    '일반 Message
    
    Rs.MoveLast
    Rs.MoveFirst
    
    '개체 생
    Set objExcel = CreateObject("Excel.Application")
    Set xlBook = objExcel.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    With xlSheet
        '열제목 나타내고 굵게 지정
        For iCol = 0 To Rs.Fields.Count - 1
            .Cells(1, iCol + 1).Value = Trim(Rs.Fields(iCol).Name)
        Next iCol
        
        iCol = 0: iRow = 0
        .Range("A1").CurrentRegion.Font.Bold = True
        
        '레코드 Excel로 변환
        For iRow = 0 To Rs.RecordCount - 1
            For iCol = 0 To Rs.Fields.Count - 1

'''                '>>> Cell
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "0.00%"
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "dd/mm/yy"
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "#####0"
'''                .Workbooks.Application.Worksheets(1).Cells(iRow + 2, iCol + 1).NumberFormat = "##0.0000"
'''
'''                '>>> Range
'''                .Workbooks.Application.Worksheets(1).Range("A:A").NumberFormat = "#####0"
'''                .Workbooks.Application.Worksheets(1).Range("B:B").NumberFormat = "######0"
'''                .Workbooks.Application.Worksheets(1).Range("H:H").NumberFormat = "##0.0000"
'''                .Workbooks.Application.Worksheets(1).Range("J:J").NumberFormat = "mm/dd/yyyy"

                '/Replace(Trim(Rs.Fields(iCol).Value & ""), vbCrLf, vbLf)     => 줄바꿈 문자열일때 음표가 엑셀에 나타나는 것을 없애줌.
                .Cells(iRow + 2, iCol + 1).Value = Replace(Trim(Rs.Fields(iCol).Value & ""), vbCrLf, vbLf)
            Next iCol
            Rs.MoveNext
        Next iRow
        
    End With
    
    '지정한 경로로 저장
    xlSheet.SaveAs xlsFileName
    
    objExcel.Quit
    Set objExcel = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
End Sub

'Excel File 열기
Public Sub EXCEL_OPEN_FILE(FilePathName As String, Optional Parameter As String)
    Dim Result As Long
    Dim Param As String
    
    On Error GoTo Err_Handler
    
    Const SW_SHOWNORMAL = 1
    
    If Parameter = "" Then
        Param = vbNull
    Else
        Param = Parameter
    End If
    
    Result = ShellExecute(App.hInstance, "open", FilePathName, Param, vbNullString, SW_SHOWNORMAL)
    '에러 났네
    If Result <= 32 Then Debug.Print Result

Exit Sub

'/----------------------------------------------------------------------------------------------------------------------------------------------------------------------------/

Err_Handler:
    ' MousePointer = vbDefault      '마우스를 원래모양으로
    Call MsgBox(Err.Description, vbCritical)
    Err.Clear
End Sub

Public Function GET_LOCAL_IP() As String
    Dim Ret         As Long
    Dim Tel         As Long
    Dim bBytes()    As Byte
    Dim Listing     As MIB_IPADDRTABLE

    Dim myByte(3) As Byte
    Dim Cnt As Long

    GET_LOCAL_IP = ""

'''    Form1.Text1 = ""

On Error GoTo END1

    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Function

    ReDim bBytes(0 To Ret - 1) As Byte
    GetIpAddrTable bBytes(0), Ret, False '/retrieve the data

    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4

    'MsgBox "IP's found : " & Listing.dEntrys    => Founded ip installed on your PC..
'''    Form1.Text1 = Listing.dEntrys & "   IP addresses found on your PC !!" & vbCrLf
'''    Form1.Text1 = Form1.Text1 & "----------------------------------------" & vbCrLf
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
       ' MsgBox bBytes(tel) & "."
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))

'''         Form1.Text1 = Form1.Text1 & "IP address                   : " & ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr) & vbCrLf
'''         Form1.Text1 = Form1.Text1 & "IP Subnetmask            : " & ConvertAddressToString(Listing.mIPInfo(Tel).dwMask) & vbCrLf
'''         Form1.Text1 = Form1.Text1 & "BroadCast IP address  : " & ConvertAddressToString(Listing.mIPInfo(Tel).dwBCastAddr) & vbCrLf
'''         Form1.Text1 = Form1.Text1 & "**************************************" & vbCrLf
    Next Tel

    GET_LOCAL_IP = ConvertAddressToString(Listing.mIPInfo(1).dwAddr)
Exit Function

'/-------------------------------------------------------------------------------------------------------------------------/

END1:
    MsgBox "Local IP 가져오기 실패", vbInformation, "확인"
End Function

Public Sub SET_CBO_DT_ALL(ArdData As String, ArdComboBox As Object)
    Dim i           As Long
    
    ArdComboBox.ListIndex = -1
    For i = 0 To ArdComboBox.ListCount - 1
        If Trim(ArdData) = Trim(ArdComboBox.List(i)) Then
            ArdComboBox.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Public Sub SET_CBO_DT_L(ArdData As String, ArdComboBox As Object, ArgLength As Integer)
    Dim i           As Long
    
    ArdComboBox.ListIndex = -1
    For i = 0 To ArdComboBox.ListCount - 1
        If Trim(ArdData) = Trim(Left(ArdComboBox.List(i), ArgLength)) Then
            ArdComboBox.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Public Sub SET_CBO_DT_R(ArdData As String, ArdComboBox As Object)
    Dim i           As Long
    
    ArdComboBox.ListIndex = -1
    For i = 0 To ArdComboBox.ListCount - 1
        If Trim(ArdData) = Trim(Right(ArdComboBox.List(i), 10)) Then
            ArdComboBox.ListIndex = i
            Exit For
        End If
    Next i
End Sub

'정해진 공간(ArgLen Byte 단위)에 문자열의 가운데 정렬
Public Function TEXT_CSET(ArgData$, ArgLen%) '입력문자(한글) ,  전체자릿수
    Dim nTemp1#, nTemp2#
    Dim ArgData1     As String
    
    ArgData1 = Left(ArgData, ArgLen)
    nTemp1 = Len(ArgData1)
    
    Do Until lstrlen(ArgData1) <= ArgLen
        nTemp1 = nTemp1 - 1
        ArgData1 = Left(ArgData, nTemp1)
    Loop
    
    nTemp2 = ArgLen - lstrlen(ArgData1)
    If nTemp2 Mod 2 = 0 Then
        TEXT_CSET = Space(nTemp2 / 2) & ArgData1 & Space(nTemp2 / 2)
    Else
        TEXT_CSET = Space(nTemp2 / 2 - 0.5) & ArgData1 & Space(nTemp2 / 2 + 0.5)
    End If
End Function

'정해진 공간(ArgLen Byte 단위)에 문자열의 왼쪽 정렬
Public Function TEXT_LSET(ArgData$, ArgLen%) '입력문자(한글) ,  전체자릿수
    Dim nTemp1%, nTemp2%
    Dim ArgData1     As String
    
    ArgData1 = Left(ArgData, ArgLen)
    nTemp1 = Len(ArgData1)
    
    Do Until lstrlen(ArgData1) <= ArgLen
        nTemp1 = nTemp1 - 1
        ArgData1 = Left(ArgData, nTemp1)
    Loop
    
    nTemp2 = ArgLen - lstrlen(ArgData1)
    TEXT_LSET = ArgData1 & Space(nTemp2)
End Function

'정해진 공간(ArgLen Byte 단위)에 문자열의 오른쪽 정렬
Public Function TEXT_RSET(ArgData$, ArgLen%) '입력문자(한글) ,  전체자릿수
    Dim nTemp1%, nTemp2%
    Dim ArgData1     As String
    
    ArgData1 = Left(ArgData, ArgLen)
    nTemp1 = Len(ArgData1)
    
    Do Until lstrlen(ArgData1) <= ArgLen
        nTemp1 = nTemp1 - 1
        ArgData1 = Left(ArgData, nTemp1)
    Loop
    
    nTemp2 = ArgLen - lstrlen(ArgData1)
    TEXT_RSET = Space(nTemp2) & ArgData1
End Function

'------------------------------------------------------------
' 목적: 화면의 Text의 내용을 선택한다.
' 입력: TEXTSELECT()
' 반환: 없음
'------------------------------------------------------------
Public Sub TEXTGF(Argctl As Object)
    Dim lngSelLength As Long
    
    If Argctl Is Nothing Then
        Exit Sub
    End If
    
    With Argctl
        If TypeOf Argctl Is TextBox Then
            lngSelLength = Len(.Text)
        ElseIf TypeOf Argctl Is MaskEdBox Then
            If Trim(.Mask) = vbNullString Then
                lngSelLength = Len(.Text)
            Else
                lngSelLength = Len(.Mask)
            End If
        ElseIf TypeOf Argctl Is ComboBox Then
            If .Style = 2 Then Exit Sub
            lngSelLength = Len(.Text)
        Else
            Exit Sub
        End If
        
       .SelStart = 0
       .SelLength = lngSelLength
    End With
End Sub
