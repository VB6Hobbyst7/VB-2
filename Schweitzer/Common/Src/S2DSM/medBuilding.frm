VERSION 5.00
Begin VB.Form frmBuilding 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Building Information"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인(&O)"
      Height          =   400
      Left            =   1980
      TabIndex        =   3
      Top             =   3450
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소(&X)"
      Height          =   400
      Left            =   915
      TabIndex        =   2
      Top             =   3450
      Width           =   1000
   End
   Begin VB.ListBox lstBldInfo 
      Height          =   2940
      Left            =   60
      TabIndex        =   0
      Top             =   465
      Width           =   2910
   End
   Begin VB.Label Label1 
      Caption         =   "건물 정보"
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "frmBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fCurIndex As Integer

Public Sub ApplyButton(ByVal pChk As String)
   If pChk = "Onlyreg" Then cmdCancel.Enabled = False
   If pChk = "Movebld" Then cmdCancel.Enabled = True
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sTemp As String, sBldCd As String, sBldNm As String, sBldNo As Integer

   If fCurIndex < 0 Then Exit Sub

   sTemp = lstBldInfo.List(lstBldInfo.ListIndex)
   
   BuildingNo = drGetP(sTemp, 1, Chr$(9))
   BuildingCd = drGetP(sTemp, 2, Chr$(9))
   BuildingNm = drGetP(sTemp, 3, Chr$(9))

   SaveSetting RegHdBld, RegSsBld, RegK1Bld, BuildingCd
   SaveSetting RegHdBld, RegSsBld, RegK2Bld, BuildingNm
   SaveSetting RegHdBld, RegSsBld, RegK3Bld, BuildingNo

   medMain.lblBld.Caption = BuildingNm

   Unload Me

End Sub

Private Sub Form_Load()

   Call LoadBuildingInfo
   
   lstBldInfo.ListIndex = fCurIndex

End Sub

Private Sub LoadBuildingInfo()
Dim sqlBld As String, dsBld As New DrSqlOcx.DrRecordSet, iBldCol As Integer
Dim sBldCd As String, sBldNm As String, sBldNo As Integer

    ' 균 코드 로드
    sqlBld = "SELECT * FROM " & TB_LAB032 & " WHERE cdindex='" & CD2_Buildings & "'" & _
             " ORDER BY field2 asc"
    iBldCol = dsBld.OpenCursor(DbConn, sqlBld)

    lstBldInfo.clear

    Dim i As Integer
    fCurIndex = -1: i = 0
    Do While dsBld.FetchCursor(iBldCol)
        i = i + 1
        sBldCd = "" & dsBld.GetValue("cdval1")
        sBldNm = "" & dsBld.GetValue("field1")
        sBldNo = Val("" & dsBld.GetValue("field2"))
        If BuildingNo = sBldNo Then fCurIndex = i - 1
        lstBldInfo.AddItem sBldNo & Chr$(9) & sBldCd & Chr$(9) & sBldNm
    Loop
    
    dsBld.CloseCursor: Set dsBld = Nothing
    
End Sub

Private Sub lstBldInfo_Click()

   fCurIndex = lstBldInfo.ListIndex

End Sub
