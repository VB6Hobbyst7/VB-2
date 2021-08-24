VERSION 5.00
Begin VB.Form frmBuilding 
   BackColor       =   &H00D8DEDA&
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
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인(&O)"
      Height          =   400
      Left            =   1965
      TabIndex        =   3
      Top             =   3465
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소(&X)"
      Height          =   400
      Left            =   900
      TabIndex        =   2
      Top             =   3465
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
      BackStyle       =   0  '투명
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

Private mvarBuildingList As String

Private mvarCurBuildingNo As String
Private mvarCurBuildingCd As String
Private mvarCurBuildingNm As String

Private fCurIndex As Long

Public Event SetBuilding()

'[속성] - 건물번호
Public Property Let CurBuildingNo(ByVal vData As String)
    mvarCurBuildingNo = vData
End Property
Public Property Get CurBuildingNo() As String
    CurBuildingNo = mvarCurBuildingNo
End Property

'[속성] - 건물명
Public Property Let CurBuildingNm(ByVal vData As String)
    mvarCurBuildingNm = vData
End Property
Public Property Get CurBuildingNm() As String
    CurBuildingNm = mvarCurBuildingNm
End Property

'[속성] - 건물코드
Public Property Let CurBuildingCd(ByVal vData As String)
    mvarCurBuildingCd = vData
End Property
Public Property Get CurBuildingCd() As String
    CurBuildingCd = mvarCurBuildingCd
End Property

'[속성] - 건물리스트
Public Property Let BuildingList(ByVal vData As String)
    mvarBuildingList = vData
End Property
Public Property Get BuildingList() As String
    BuildingList = mvarBuildingList
End Property


Public Sub ApplyButton(ByVal pChk As String)
   If pChk = "Onlyreg" Then cmdCancel.Enabled = False
   If pChk = "Movebld" Then cmdCancel.Enabled = True
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   Set frmBuilding = Nothing
End Sub

Private Sub cmdOK_Click()
Dim sTemp As String, sBldCd As String, sBldNm As String, sBldNo As Integer

   If fCurIndex < 0 Then Exit Sub

   sTemp = lstBldInfo.List(lstBldInfo.ListIndex)
   
   mvarCurBuildingNo = medGetP(sTemp, 1, Chr$(9))
   mvarCurBuildingCd = medGetP(sTemp, 2, Chr$(9))
   mvarCurBuildingNm = medGetP(sTemp, 3, Chr$(9))

   SaveSetting RegAppName, RegSsBld, RegK1Bld, mvarCurBuildingCd
   SaveSetting RegAppName, RegSsBld, RegK2Bld, mvarCurBuildingNm
   SaveSetting RegAppName, RegSsBld, RegK3Bld, mvarCurBuildingNo

   RaiseEvent SetBuilding

   Unload Me
   Set frmBuilding = Nothing

End Sub

Private Sub lstBldInfo_Click()

   fCurIndex = lstBldInfo.ListIndex

End Sub

Public Sub LoadBuildingList()

    Dim i As Long
    Dim strRow As String
    
    i = 1
    lstBldInfo.Clear
    strRow = medGetP(mvarBuildingList, i, vbCr)
    
    While (Trim(strRow) <> "")
        lstBldInfo.AddItem strRow
        If medGetP(strRow, 1, vbTab) = mvarCurBuildingNo Then fCurIndex = i - 1
        i = i + 1: strRow = medGetP(mvarBuildingList, i, vbCr)
    Wend
   
    lstBldInfo.ListIndex = fCurIndex
    
End Sub

