Attribute VB_Name = "modSlideImage"
Option Explicit
Option Compare Text

Public Enum opgParsePath
    FILE_ONLY
    PATH_ONLY
    DRIVE_ONLY
    FILEEXT_ONLY
End Enum


'Public P_SLIDE_CLIENT_PATH As String
Global Const SLIDE_DIAGNOSIS_IMAGE = "진단"
'Global Const SLIDE_GROSS_IMAGE = "육안"
Global Const MTS_COL = "|"
Global Const MTS_ROW = "^"
Global Const MTS_TCOL = "∮"
Global Const MTS_TROW = "ː"


'Const P_BLOCK_SIZE = 16384
'#
'Declare Function LockWindowUpdate Lib "user32" _
                (ByVal hwndLock As Long) As Long
                                
Public Function ParseFilePath(strTmp As String)
'파일 패스명에서 해당하는 파일패스를 가져온다.
   '
   ParseFilePath = Mid(strTmp, 1, InStrRev(strTmp, "\", , vbTextCompare))
   '
End Function

Public Function ParseFileName(strTmp As String)
'파일 패스명에서 해당하는 파일명을 가져온다.
   '
   ParseFileName = Mid(strTmp, InStrRev(strTmp, "\", , vbTextCompare) + 1, Len(strTmp))
   '
End Function
                
Public Function P(ByVal expression As String, ByVal Delim As String, _
           ByVal Piece As Integer) As String
Dim CNTA, CNTB As Integer
    P = ""
    CNTA = 0
    CNTB = 0
    Do
        If CNTB = Piece - 1 Then Exit Do
        CNTA = InStr(CNTA + 1, expression, Delim)
        If CNTA <> 0 Then CNTB = CNTB + 1
    Loop Until CNTA = 0
    If CNTA = 0 And Piece <> 1 Then Exit Function
    CNTB = InStr(CNTA + 1, expression, Delim)
    If CNTB = 0 Then CNTB = Len(expression) + 1
    P = Mid$(expression, CNTA + 1, CNTB - CNTA - 1)
End Function

Public Function ParsePath(strPath As String, _
                          lngPart As opgParsePath) As String
    Dim lngPos          As Long
    Dim strPart         As String
    Dim blnIncludesFile As Boolean
    lngPos = InStrRev(strPath, "\")
    blnIncludesFile = InStrRev(strPath, ".") > lngPos
    If lngPos > 0 Then
        Select Case lngPart
            Case opgParsePath.FILE_ONLY
                If blnIncludesFile Then
                    strPart = Mid(strPath, lngPos + 1)
                Else
                    strPart = ""
                End If
            Case opgParsePath.PATH_ONLY
                If blnIncludesFile Then
                    strPart = Mid(strPath, 1, lngPos)
                Else
                    strPart = strPath
                End If
            Case opgParsePath.DRIVE_ONLY
                strPart = Mid(strPath, 1, 3)
            Case opgParsePath.FILEEXT_ONLY
                If blnIncludesFile Then
                    strPart = Mid(strPath, InStrRev(strPath, ".") + 1, 3)
                Else
                    strPart = ""
                End If
            Case Else
                strPart = ""
        End Select
    End If
    ParsePath = strPart

ParsePath_End:
    Exit Function

End Function

Public Function FillDictionary(strPath As String) As Scripting.Dictionary
    Dim fsoSysObj As Scripting.FileSystemObject
    Dim fdrFolder As Scripting.Folder
    Dim filFile As Scripting.File
    Dim dctImages As Scripting.Dictionary
    Set fsoSysObj = New FileSystemObject
    'Set fdrFolder = fsoSysObj.GetFolder("C:\ANA\Anatomic\SlideImage\")
    Set fdrFolder = fsoSysObj.GetFolder(P_SLIDE_CLIENT_PATH)
    Set dctImages = New Scripting.Dictionary
    
    For Each filFile In fdrFolder.Files
        Select Case ParsePath(filFile.Path, FILEEXT_ONLY)
            Case "bmp", "wmf", "gif", "jpg"
                dctImages.Add filFile.Path, filFile.Name
        End Select
    Next
    Set FillDictionary = dctImages
End Function

Public Function HExtract(ByVal Var As String, ByVal Del As String, ByVal GetCnt As Long) As String
' (c)Copyright 1998.11.05  made by hyuntae Jo  한글 단어보호
    Dim BUF As String, tmp As String, Num As Long, Cnt As Long
    
    BUF = "": Cnt = 0: Num = 0
    If Var = "" Or GetCnt < 2 Then HExtract = "": Exit Function
    Do
        Cnt = Cnt + 1: tmp = Mid(Var, Cnt, 1): Num = Num + 1
        If Asc(tmp) < 0 Then Num = Num + 1
        If Num < GetCnt Then
            BUF = BUF + tmp
        ElseIf Num = GetCnt Then
            Num = 0: BUF = BUF + tmp + Del
        ElseIf Num > GetCnt Then
            Num = 2: BUF = BUF + Del + tmp
        End If
    Loop Until Cnt >= Len(Var)
    If VBA.Strings.Right(BUF, 1) = Del Then BUF = VBA.Strings.Left(BUF, Len(BUF) - 1)
    HExtract = BUF
End Function

Function L(ByVal Var As String, ByVal Del As String) As Long
' (c)Copyright 1997.05.03  made by hyuntae Jo   multi-delimiter support
    Dim Srt As Long, Nxt As Long, Cnt As Long
    
    If Del = "" Then L = 0: Exit Function
    Nxt = (Len(Del) * -1) + 1
    Do
        Srt = Nxt + Len(Del): Nxt = InStr(Srt, Var, Del)
        Cnt = Cnt + 1
    Loop Until Nxt = 0
    L = Cnt
End Function

'Public Sub medInitLvwHead(ByRef objLvw As ListView, ByVal strHead As String, _
'   ByVal strSize As String)
'Dim ii As Integer
'Dim aryTitle() As String
'Dim aryWidth() As String
'   aryTitle = Split(strHead, ",")
'   aryWidth = Split(strSize, ",")
'   If UBound(aryWidth) < UBound(aryTitle) Then
'      ReDim Preserve aryWidth(UBound(aryTitle))
'   End If
'   With objLvw
'      .ColumnHeaders.Clear
'      For ii = 0 To UBound(aryTitle)
'         If aryWidth(ii) = "" Then
'            aryWidth(ii) = "0"
'         End If
'        .ColumnHeaders.Add ii + 1, aryTitle(ii), aryTitle(ii), _
'            (.Width \ (UBound(aryTitle) + 1)) + Val(aryWidth(ii)), vbLeftJustify
'      Next ii
'      .View = lvwReport            ' Report Style
'   End With
'End Sub

Public Function DctToStr(ByRef dctTmp As Scripting.Dictionary) As String
Dim varKey As Variant
Dim aryTmp() As String
Dim blnFirst As Boolean
   'varkeyTmp = dctTmp.Keys
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

Public Sub BlobToFile(fld As ADODB.Field, ByVal FName As String, _
                     Optional FieldSize As Long = -1, _
                     Optional Threshold As Long = 1048576)
'
' Assumes file does not exist '1048576
' Data cannot exceed approx. 2Gb in size
'
Dim F As Long, bData() As Byte, sData As String
  F = FreeFile
  Open FName For Binary As #F
  Select Case fld.Type
    Case adLongVarBinary
      If FieldSize = -1 Then   ' blob field is of unknown size
        WriteFROMUnsizedBinary F, fld
      Else                     ' blob field is of known size
        If FieldSize > Threshold Then   ' very large actual data
          WriteFROMBinary F, fld, FieldSize
        Else                            ' smallish actual data
          bData = fld.Value
          Put #F, , bData  ' PUT tacks on overhead if use fld.Value
        End If
      End If
    Case adLongVarChar, adLongVarWChar
      If FieldSize = -1 Then
        WriteFROMUnsizedText F, fld
      Else
        If FieldSize > Threshold Then
          WriteFROMText F, fld, FieldSize
        Else
          sData = fld.Value
          Put #F, , sData  ' PUT tacks on overhead if use fld.Value
        End If
      End If
  End Select
  Close #F
End Sub

Public Sub WriteFROMBinary(ByVal F As Long, fld As ADODB.Field, _
                    ByVal FieldSize As Long)
Dim Data() As Byte, BytesRead As Long
  Do While FieldSize <> BytesRead
    If FieldSize - BytesRead < P_BLOCK_SIZE Then
      Data = fld.GetChunk(FieldSize - P_BLOCK_SIZE)
      BytesRead = FieldSize
    Else
      Data = fld.GetChunk(P_BLOCK_SIZE)
      BytesRead = BytesRead + P_BLOCK_SIZE
    End If
    Put #F, , Data
  Loop
End Sub

Public Sub WriteFROMUnsizedBinary(ByVal F As Long, fld As ADODB.Field)
Dim Data() As Byte, Temp As Variant
  Do
    Temp = fld.GetChunk(P_BLOCK_SIZE)
    If IsNull(Temp) Then Exit Do
    Data = Temp
    Put #F, , Data
  Loop While LenB(Temp) = P_BLOCK_SIZE
End Sub

Public Sub WriteFROMText(ByVal F As Long, fld As ADODB.Field, _
                  ByVal FieldSize As Long)
Dim Data As String, CharsRead As Long
  Do While FieldSize <> CharsRead
    If FieldSize - CharsRead < P_BLOCK_SIZE Then
      Data = fld.GetChunk(FieldSize - P_BLOCK_SIZE)
      CharsRead = FieldSize
    Else
      Data = fld.GetChunk(P_BLOCK_SIZE)
      CharsRead = CharsRead + P_BLOCK_SIZE
    End If
    Put #F, , Data
  Loop
End Sub

Public Sub WriteFROMUnsizedText(ByVal F As Long, fld As ADODB.Field)
Dim Data As String, Temp As Variant
  Do
    Temp = fld.GetChunk(P_BLOCK_SIZE)
    If IsNull(Temp) Then Exit Do
    Data = Temp
    Put #F, , Data
  Loop While Len(Temp) = P_BLOCK_SIZE
End Sub

Public Sub FileToBlob(ByVal FName As String, fld As ADODB.Field, _
               Optional Threshold As Long = 1048576)
'
' Assumes file exists
' Assumes calling routine does the UPDATE
' File cannot exceed approx. 2Gb in size
'
Dim F As Long, Data() As Byte, FileSize As Long
  F = FreeFile
  Open FName For Binary As #F
  FileSize = LOF(F)
  Select Case fld.Type
    Case adLongVarBinary
      If FileSize > Threshold Then
        ReadToBinary F, fld, FileSize
      Else
        Data = InputB(FileSize, F)
        fld.Value = Data
      End If
    Case adLongVarChar, adLongVarWChar
      If FileSize > Threshold Then
        ReadToText F, fld, FileSize
      Else
        fld.Value = Input(FileSize, F)
      End If
  End Select
  Close #F
End Sub

Public Sub ReadToBinary(ByVal F As Long, fld As ADODB.Field, _
                 ByVal FileSize As Long)
Dim Data() As Byte, BytesRead As Long
  Do While FileSize <> BytesRead
    If FileSize - BytesRead < P_BLOCK_SIZE Then
      Data = InputB(FileSize - BytesRead, F)

      BytesRead = FileSize
    Else
      Data = InputB(P_BLOCK_SIZE, F)
      BytesRead = BytesRead + P_BLOCK_SIZE
    End If
    fld.AppendChunk Data
  Loop
End Sub

Public Sub ReadToText(ByVal F As Long, fld As ADODB.Field, _
               ByVal FileSize As Long)
Dim Data As String, CharsRead As Long
  Do While FileSize <> CharsRead
    If FileSize - CharsRead < P_BLOCK_SIZE Then
      Data = Input(FileSize - CharsRead, F)
      CharsRead = FileSize
    Else
      Data = Input(P_BLOCK_SIZE, F)
      CharsRead = CharsRead + P_BLOCK_SIZE
    End If
    fld.AppendChunk Data
  Loop
End Sub

