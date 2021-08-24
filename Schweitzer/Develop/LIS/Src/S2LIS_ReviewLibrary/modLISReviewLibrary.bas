Attribute VB_Name = "modLISReviewLibrary"
Option Explicit
'

Global gBuildingCd      As String
Global gBuildingNm      As String
Global gBuildingNo      As Long
Global objBuildings     As New clsDictionary
Global objDeptDic       As New clsDictionary
'
'
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Global gIsDeveloper     As Boolean
Global gEmpId           As String
Global gDeptCd          As String
Global gPatientId       As String

Global gUsingInWardMenu As Boolean

Global Const MTS_COL = "|"
Global Const MTS_ROW = "^"
Global Const MTS_TCOL = "∮"
Global Const MTS_TROW = "ː"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, ByVal lpOperation As String, _
             ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function IsLastForm() As Boolean

    Dim I As Long
    Dim tmpFrm As Form
    
    I = 0
    IsLastForm = False
    
    For Each tmpFrm In Forms
        I = I + 1
    Next
    If I = 0 Then IsLastForm = True

End Function


Public Sub HighlightText(ByVal pTextBox As Object, ByVal pText As String, _
                        ByVal InitFg As Boolean, Optional ByVal FtName As String, _
                        Optional COLOR As Long = &H80&, Optional ByVal FtSize As Long)
   With pTextBox
      If InitFg Then
         .SelStart = 0
         .SelLength = Len(.Text)
         .SelProtected = False
         .SelColor = &H0&
         '.SelBold = False
      End If
      
      Dim Point2 As Long
      Point2 = .Find(pText, 0, , rtfWholeWord)
      If Point2 <> -1 Then
         .SelStart = Point2
         .SelLength = Len(pText)
         .SelProtected = False
         .SelColor = COLOR         '&HFF8080       '&H8080FF           '&HDF6A3E
         '.SelBold = True
      End If
      .SelLength = 0
   End With
End Sub

Public Sub BlobToFile(fld As ADODB.Field, ByVal FName As String, _
                     Optional FieldSize As Long = -1, _
                     Optional Threshold As Long = 1048576)
    Dim F       As Long
    Dim bData() As Byte
    Dim sData   As String
    
    F = FreeFile
    Open FName For Binary As #F
    
    Select Case fld.Type
        Case adLongVarBinary
            If FieldSize = -1 Then   ' blob field is of unknown size
                WriteFromUnsizedBinary F, fld
            Else                     ' blob field is of known size
                If FieldSize > Threshold Then   ' very large actual data
                    WriteFromBinary F, fld, FieldSize
                Else                            ' smallish actual data
                    bData = fld.Value
                    Put #F, , bData  ' PUT tacks on overhead if use fld.Value
                End If
            End If
        Case adLongVarChar, adLongVarWChar
            If FieldSize = -1 Then
                WriteFromUnsizedText F, fld
            Else
                If FieldSize > Threshold Then
                    WriteFromText F, fld, FieldSize
                Else
                    sData = fld.Value
                    Put #F, , sData  ' PUT tacks on overhead if use fld.Value
                End If
            End If
    End Select
    Close #F
End Sub

Public Sub WriteFromBinary(ByVal F As Long, fld As ADODB.Field, _
                           ByVal FieldSize As Long)
    Dim Data()      As Byte
    Dim BytesRead   As Long
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

Public Sub WriteFromUnsizedBinary(ByVal F As Long, fld As ADODB.Field)
    Dim Data()  As Byte
    Dim Temp    As Variant
    
    Do
        Temp = fld.GetChunk(P_BLOCK_SIZE)
        If IsNull(Temp) Then Exit Do
        Data = Temp
        Put #F, , Data
    Loop While LenB(Temp) = P_BLOCK_SIZE
End Sub

Public Sub WriteFromText(ByVal F As Long, fld As ADODB.Field, _
                  ByVal FieldSize As Long)
    Dim Data        As String
    Dim CharsRead   As Long
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

Public Sub WriteFromUnsizedText(ByVal F As Long, fld As ADODB.Field)
    Dim Data As String
    Dim Temp As Variant
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
    Dim F        As Long
    Dim Data()   As Byte
    Dim FileSize As Long
    
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
    Dim Data()      As Byte
    Dim BytesRead   As Long
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
    Dim Data        As String
    Dim CharsRead   As Long
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

'Public Sub InitLvwHead(ByRef objLvw As ListView, ByVal strHead As String, _
'                       ByVal strSize As String)
'    Dim ii As Integer
'    Dim aryTitle() As String
'    Dim aryWidth() As String
'
'    aryTitle = Split(strHead, ",")
'    aryWidth = Split(strSize, ",")
'    If UBound(aryWidth) < UBound(aryTitle) Then
'        ReDim Preserve aryWidth(UBound(aryTitle))
'    End If
'    With objLvw
'        .ColumnHeaders.Clear
'        For ii = 0 To UBound(aryTitle)
'           If aryWidth(ii) = "" Then
'              aryWidth(ii) = "0"
'           End If
'          .ColumnHeaders.Add ii + 1, aryTitle(ii), aryTitle(ii), _
'              (.Width \ (UBound(aryTitle) + 1)) + Val(aryWidth(ii)), vbLeftJustify
'        Next ii
'        .View = lvwReport            ' Report Style
'    End With
'End Sub
'
'Public Sub DataLoadLvw(ByRef objLvw As ListView, _
'                       ByVal RowDel As String, ByVal ColDel As String, _
'                       ByVal strData As String)
'    Dim itmx        As ListItem
'    Dim strTmp      As String
'    Dim aryTmp()    As String
'    Dim ii          As Integer
'    Dim jj          As Integer
'    Dim intCol      As Integer
'
'   aryTmp = Split(medGetP(strData, 1, RowDel), ColDel)
'   intCol = UBound(aryTmp) + 1
'   '
'   aryTmp = Split(strData, RowDel)
'   If (UBound(aryTmp) + 1) < 1 Then Exit Sub
'   For ii = 0 To UBound(aryTmp)
'      For jj = 1 To intCol
'         If jj = 1 Then
'            Set itmx = objLvw.ListItems.Add(, , medGetP(aryTmp(ii), jj, ColDel))
'         Else
'            If medGetP(aryTmp(ii), jj, ColDel) <> "" Then
'               itmx.SubItems(jj - 1) = medGetP(aryTmp(ii), jj, ColDel)
'            Else
'               itmx.SubItems(jj - 1) = " "
'            End If
'         End If
'      Next jj
'
'   Next ii
'   '
'End Sub

'[검사항목별 Comment]-----
Public Function TestItemToolTipString(ByVal TestCd As String) As String
    '
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim aryTmp()    As String
    Dim strTmp      As String
    Dim I As Long
    
    SSQL = " SELECT * FROM " & T_LAB034 & _
           " WHERE " & DBW("cdindex=", LC4_TestItemComment) & _
           " AND " & DBW("cdval1=", TestCd)
           
    Set RS = New Recordset
    
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        strTmp = "  검사항목 Comment :" & vbCrLf
        aryTmp() = Split(RS.Fields("text1").Value & "", vbCrLf)
        
        For I = LBound(aryTmp()) To UBound(aryTmp())
            strTmp = strTmp & "     " & aryTmp(I) & vbCrLf
        Next
    End If
    RS.Close
    Set RS = Nothing
    
    TestItemToolTipString = strTmp
End Function

