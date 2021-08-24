VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDSM005P 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "인적 조회"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmDSM005P.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin MSComctlLib.ListView lvwEmpList 
      Height          =   3465
      Left            =   150
      TabIndex        =   2
      Top             =   135
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   6112
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776191
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "직원ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "이름"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00EBF3ED&
      Caption         =   "취 소(&C)"
      Height          =   510
      Left            =   2790
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00EBF3ED&
      Caption         =   "확 인(&O)"
      Height          =   510
      Left            =   1560
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmDSM005P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private clsRef As New clsDSMUserInfo
'Private strItm As String
'
'Private Sub cmdOK_Click()
'    Query_EmpInfo
'
'    With frmDSM005
'         .LislblEmpID(0).Caption = clsRef.EmpId
'         .LislblEmpID(1).Caption = clsRef.EmpNm
'         .LislblDept.Caption = clsRef.DeptCd
'    End With
'
'    Unload Me
'End Sub
'
'Private Sub Command2_Click()
'    Unload Me
'    Set clsRef = Nothing
'End Sub
'
'Private Sub Form_Load()
'    strItm = ""
'    Query_lvwEmpList
'End Sub
'
'Private Sub Query_lvwEmpList()
'    Dim Rs As New Recordset
'    Dim LvwItm As Object
'    Dim strSQL As String
'
'    On Error GoTo ErrlvwEmpList
'
'    lvwEmpList.ListItems.clear
'
'    strSQL = " select empid, empnm, deptcd from " & T_COM006
'
'    Rs.Open strSQL, DBConn
'
''    If Rs.DBerror = True Then
''       'Call DisplayErrors
''       Set Rs = Nothing
''       Exit Sub
''    End If
'
'    While Rs.EOF = False
'          Set LvwItm = lvwEmpList.ListItems.Add()
'
'          With LvwItm
'               .Text = "" & Rs.Fields("empid").Value
'               .SubItems(1) = IIf(IsNull(Rs.Fields("empnm").Value) = True, "", "" & Rs.Fields("empnm").Value)
'
'               Rs.MoveNext
'          End With
'    Wend
'
'    Set Rs = Nothing
'    Exit Sub
'ErrlvwEmpList:
'    Set Rs = Nothing
'    MsgBox Err.Description, vbCritical, "일반 오류"
'
'End Sub
'
'Private Sub lvwEmpList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Static i As Integer
'
'    With lvwEmpList
'         .SortKey = ColumnHeader.Index - 1
'         .SortOrder = IIf(i = 0, lvwAscending, lvwDescending)
'         .Sorted = True
'    End With
'
'    i = (i + 1) Mod 2
'End Sub
'
'Private Sub Query_EmpInfo()
'    If strItm <> "" Then
'       If clsRef.EmpInfo(strItm) = False Then Exit Sub
'    End If
'End Sub
'
'Private Sub lvwEmpList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    strItm = Item.Text
'End Sub
