Attribute VB_Name = "modLab"
Option Explicit

'Global Const Splt_Delimeter = "$"
'Global Const vbLockColor = vb3DFace
'Global Const ScrEmpId$ = "E0102"

'Type ptInfo
'   Name As String
'   Location As String
'   Sex As String
'   Age As String
'   DOB As String
'End Type
'
'Type tLabno
'    sWorkArea As String
'    sAccDt As String
'    iAccSeq As Integer
'End Type
'
'
'Public Type ResultTextTbl
'    sTCd As String * 1
'    TPCD As String
'    TPNM As String
'    TPDATA As String
'
'End Type
'
'Public Type SpeItemTbl
'    STITEM As String
'    TestCd As String
'End Type
'
'Public Enum ColNo          '미생물 Growh,Nogrowth 에서 사용되는것임.
'    cnCOL0 = 0
'    cnAccNo
'    cnPtid
'    cnPtNm
'    cnSA
'    cnSpcNm
'    cnLstRst
'    cnCurRst
'    cnMic
'    cnTestCd
'    cnWsCd
'    cnWsUnit
'    cnHold
'    cnSpcCd
'    cnWarn
'End Enum

'Global SpecialItem() As SpeItemTbl ' temporary table for special item
'Global ResultText() As ResultTextTbl
'Global MaxCnt ' max number of array
'Global SMaxCnt ' max number of special item
'Global formLoadCase As Integer
'Global ChosenCodeNm As String

'Global PtDemo As ptInfo

'Public Sub DemoInit()
'   With PtDemo
'      .Name = "Test Patient"
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
'Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
'  '
'  Call SendMessage(FormName.hWnd, WM_RBUTTONDOWN, 0, 0&)
'  '
'  FormName.PopupMenu MenuName
'  '
'End Sub

'Public Sub 'FormOnTop(HANDle As Integer, OnTop As Boolean)
'    Dim wFlags As Long, PosFlag As Long
'    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or _
'        SWP_SHOWWINDOW Or SWP_NOACTIVATE
'    SELECT Case OnTop
'        Case True
'            PosFlag = HWND_TOPMOST
'        Case False
'            PosFlag = HWND_NOTOPMOST
'    End SELECT
'    SetWindowPos HANDle, PosFlag, 0, 0, 0, 0, wFlags
'End Sub

'Public Sub CenterForm(Frm As Form)
'    Frm.Move (Screen.Width - Frm.Width) \ 2, _
'        (Screen.Height - Frm.Height) \ 2
'End Sub

'Public Function IsLeapYear(iYear As Integer)
'    '-- Check for leap year
'    If (iYear Mod 4 = 0) And _
'    ((iYear Mod 100 <> 0) Or (iYear Mod 400 = 0)) Then
'        IsLeapYear = True
'    Else
'        IsLeapYear = False
'    End If
'End Function

'Public Function DateStr(ByVal pDate As Date) As String
'   DateStr = Format(pDate, "yyyymmdd")
'End Function

'Public Function DateDpt(ByVal pDate As String) As Date
'   DateDpt = CDate(Format(pDate, "####-##-##"))
'End Function

'Public Function DateSys(ByVal pDate As Date) As Date
'   DateSys = Format(pDate, "yyyy-mm-dd")
'End Function

'Public Function LvwClickData(ByVal Item As ListItem) As String
'Dim ii As Integer
'Dim strTmpRecord As String
'   Item.Ghosted = Abs(Item.Ghosted) - 1
'   LvwClickData = Item.Text
'   For ii = 1 To Item.ListSubItems.Count
'      LvwClickData = LvwClickData & vbTab & CStr(Item.SubItems(ii))
'   Next ii
'End Function

'Public Sub medDataLoadLvw(ByRef objLvw As ListView, _
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
