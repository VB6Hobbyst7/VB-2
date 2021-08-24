VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDonorFind 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "헌혈자 선택"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmDonorFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00F4F0F2&
      Caption         =   "선택(&O)"
      Default         =   -1  'True
      Height          =   510
      Left            =   345
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   3510
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Cancel          =   -1  'True
      Caption         =   "취소(&C)"
      Height          =   510
      Left            =   1680
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   3510
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   3345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "이    름"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "   주민    번호"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   " 생년  월일"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "성 별"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "혈액형"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "헌혈횟수"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "총헌혈량"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmDonorFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Donorid As String
Public donornm As String
Public dob As String
Public sex As String
Public ABO As String
Public cnt As String
Public totvol As String
Public ssn As String

Public isSelect As Boolean



Private Sub cmdCancel_Click()
    isSelect = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim itmX As ListItem
    Dim strAge As String
    
    Set itmX = lvwPtList.SelectedItem
    If itmX Is Nothing Then Exit Sub

    With itmX
        Donorid = .Text
        donornm = .SubItems(1)
        strAge = DateDiff("yyyy", Format(.SubItems(3), "yyyy-MM-dd"), GetSystemDate)
        dob = .SubItems(3)
        sex = .SubItems(4) & "/" & strAge
        ABO = .SubItems(5)
        cnt = .SubItems(6)
        totvol = .SubItems(7)
        If Mid(.SubItems(2), 8, 1) = "1" Or Mid(.SubItems(2), 8, 1) = "2" Then
            ssn = .SubItems(2)
            ssn = Replace(ssn, "-", "")
        Else
            ssn = .SubItems(2)
            ssn = Replace(ssn, "-", "")
        End If
    End With

    isSelect = True
    Unload Me
End Sub

Private Sub lvwPtList_DblClick()
    Call cmdOk_Click
End Sub
