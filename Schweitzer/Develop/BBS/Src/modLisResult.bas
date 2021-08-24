Attribute VB_Name = "modLab"
Option Explicit

'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long

Global Const Splt_Delimeter = "$"
Global Const vbLockColor = vb3DFace
Global Const ScrEmpId$ = "E0102"

Type ptInfo
   name As String
   Location As String
   Sex As String
   Age As String
   DOB As String
End Type

Type tLabno
    sWorkArea As String
    sAccDt As String
    iAccSeq As Integer
End Type


Public Type ResultTextTbl
    sTCd As String * 1
    TPCD As String
    TPNM As String
    TPDATA As String
    
End Type

Public Type SpeItemTbl
    STITEM As String
    TestCd As String
End Type

Global SpecialItem() As SpeItemTbl ' temporary table for special item
Global ResultText() As ResultTextTbl
Global MaxCnt ' max number of array
Global SMaxCnt ' max number of special item
Global formLoadCase As Integer
Global ChosenCodeNm As String
Global gDeptCd As String
Global gPatientId As String
Global gUsingInWardMenu As Boolean

Global PtDemo As ptInfo


Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1

Public CmdLine As String

Private Sub Main()
'    Dim CmdLine As String
    
    CmdLine = Command()
    
    medMain.Show
End Sub

'Public Sub DemoInit()
'   With PtDemo
'      .name = "Test Patient"
'      .Sex = "Female"
'      .Age = "29"
'      .DOB = "01/08/1971"
'      .Location = "042W 12R 3B"
'   End With
'End Sub

'Public Sub FillList(ListControl As ListBox, ParamArray Items())
'    Dim i As Variant
'    With ListControl
'        .Clear
'        For Each i In Items
'            .AddItem i
'        Next
'    End With
'End Sub
'Public Sub FocusMe(ctlName As Control)
'    With ctlName
'        .SelStart = 0
'        .SelLength = Len(ctlName)
'    End With
'End Sub
'Public Sub CenterForm(Frm As Form)
'    Frm.Move (Screen.Width - Frm.Width) \ 2, _
'        (Screen.Height - Frm.Height) \ 2
'End Sub


Public Function IsLeapYear(iYear As Integer)
    '-- Check for leap year
    If (iYear Mod 4 = 0) And _
    ((iYear Mod 100 <> 0) Or (iYear Mod 400 = 0)) Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If
End Function

Public Function DateStr(ByVal pDate As Date) As String
   DateStr = Format(pDate, "yyyymmdd")
End Function

Public Function DateDpt(ByVal pDate As String) As Date
   DateDpt = CDate(Format(pDate, "####-##-##"))
End Function

Public Function DateSys(ByVal pDate As Date) As Date
   DateSys = Format(pDate, "yyyy-mm-dd")
End Function

Public Function LvwClickData(ByVal Item As ListItem) As String
Dim ii As Integer
Dim strTmpRecord As String
   Item.Ghosted = Abs(Item.Ghosted) - 1
   LvwClickData = Item.Text
   For ii = 1 To Item.ListSubItems.Count
      LvwClickData = LvwClickData & vbTab & CStr(Item.SubItems(ii))
   Next ii
End Function


   '
'Public Sub DataLoadLvw(ByRef objLvw As ListView, _
'   ByVal RowDel As String, ByVal ColDel As String, _
'   ByVal strData As String)
'Dim iTmx As ListItem
'Dim strTmp As String
'Dim aryTmp() As String
'Dim ii As Integer
'Dim jj As Integer
'Dim intCol As Integer
'   aryTmp = Split(medGetP(strData, 1, RowDel), ColDel)
'   intCol = UBound(aryTmp) + 1
'   '
'   aryTmp = Split(strData, RowDel)
'   If (UBound(aryTmp) + 1) < 1 Then Exit Sub
'   For ii = 0 To UBound(aryTmp)
'      For jj = 1 To intCol
'         If jj = 1 Then
'            Set iTmx = objLvw.ListItems.Add(, , medGetP(aryTmp(ii), jj, ColDel))
'         Else
'            If medGetP(aryTmp(ii), jj, ColDel) <> "" Then
'               iTmx.SubItems(jj - 1) = medGetP(aryTmp(ii), jj, ColDel)
'            Else
'               iTmx.SubItems(jj - 1) = " "
'            End If
'         End If
'      Next jj
'
'   Next ii
'   '
'End Sub


Public Sub GetPtTelInfo(ByVal strWorkArea As String, ByVal strAccDt As String, ByVal strAccSeq As String, _
                        ByVal objTel As Object)
    
    Dim RS          As Recordset
    Dim strCdval1   As String
    Dim SSQL        As String
    
    objTel.Caption = ""
    
    SSQL = " select ptid,wardid,deptcd from " & T_LAB201 & _
           " where " & _
                     DBW("workarea=", strWorkArea) & _
           " and " & DBW("accdt=", strAccDt) & _
           " and " & DBW("accseq=", strAccSeq)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        If Trim(RS.Fields("wardid").value & "") = "" Then
            strCdval1 = RS.Fields("deptcd").value & ""
        Else
            strCdval1 = RS.Fields("wardid").value & ""
        End If
        
    
        SSQL = "select * from " & T_LAB032 & _
               " where " & _
                           DBW("cdindex=", LC2_TelePhone) & _
               " and   " & DBW("cdval1=", strCdval1)
        
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        
        If Not RS.EOF Then
            objTel.Caption = "[" & strCdval1 & "]   " & RS.Fields("field1").value & ""
        End If
    End If
    Set RS = Nothing
End Sub

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
          bData = fld.value
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
          sData = fld.value
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

