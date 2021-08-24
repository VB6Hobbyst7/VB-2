Attribute VB_Name = "modLISCollectLibrary"
Option Explicit
'

Global gBuildingCd As String
Global gBuildingNm As String
Global gBuildingNo As Long
Global gUsingInWardMenu As Boolean

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Global gIsDeveloper As Boolean
Global gEmpId As String
Global gDeptCd As String
Global gWardId As String
Global gWardNm As String

Public Function IsLastForm() As Boolean
    Dim i As Long
    Dim tmpFrm As Form
    
    i = 0
    IsLastForm = False
    
    For Each tmpFrm In Forms
        i = i + 1
    Next
    If i = 0 Then IsLastForm = True

End Function

'Public Sub DataLoadLvw(ByRef objLvw As MSComctlLib.ListView, _
'   ByVal RowDel As String, ByVal ColDel As String, _
'   ByVal strData As String, Optional strTag As String)
'Dim itmX As ListItem
'Dim strtmp As String
'Dim AryTmp() As String
'Dim aryTag() As String
'Dim ii As Integer
'Dim jj As Integer
'Dim intCol As Integer
'   AryTmp = Split(medGetP(strData, 1, RowDel), ColDel)
'   If IsMissing(strTag) Then
'      strTag = ""
'   End If
'   aryTag = Split(strTag, RowDel)
'   intCol = UBound(AryTmp) + 1
'   '
'   AryTmp = Split(strData, RowDel)
'   If UBound(AryTmp) > UBound(aryTag) Then
'      ReDim Preserve aryTag(UBound(AryTmp))
'   End If
'   If (UBound(AryTmp) + 1) < 1 Then Exit Sub
'   For ii = 0 To UBound(AryTmp)
'      For jj = 1 To intCol
'         If jj = 1 Then
'            Set itmX = objLvw.ListItems.Add(, , medGetP(AryTmp(ii), jj, ColDel))
'         Else
'            If medGetP(AryTmp(ii), jj, ColDel) <> "" Then
'               itmX.SubItems(jj - 1) = medGetP(AryTmp(ii), jj, ColDel)
'            Else
'               itmX.SubItems(jj - 1) = " "
'            End If
'         End If
'         itmX.Tag = aryTag(ii)
'      Next jj
'
'   Next ii
'   Set itmX = Nothing
'   '
'End Sub



'Public Sub GetBarInfo(ByVal strOrdDiv As String)
'
'    '바코드 출력양식 읽어오기
'    SELECT Case strOrdDiv
'    Case "A"
'        If Not blnAPSBarFg Then
'            Set objAPSbarcode = New clsBarcode
'            Set objAPSbarcode.MyDB = dbconn
'            objAPSbarcode.ProjectCd = "APS"
'            Call objAPSbarcode.GetBarConfig
'            blnAPSBarFg = True
'        End If
'    Case "B"
'        If Not blnBBSBarFg Then
'            Set objBBSbarcode = New clsBarcode
'            Set objBBSbarcode.MyDB = dbconn
'            objBBSbarcode.ProjectCd = "BBS"
'            Call objBBSbarcode.GetBarConfig
'            blnBBSBarFg = True
'        End If
'    Case "L"
'        If Not blnLISBarFg Then
'            Set objLISbarcode = New clsBarcode
'            Set objLISbarcode.MyDB = dbconn
'            objLISbarcode.ProjectCd = "LIS"
'            Call objLISbarcode.GetBarConfig
'            blnLISBarFg = True
'        End If
'    End SELECT
'
'End Sub
'
'
