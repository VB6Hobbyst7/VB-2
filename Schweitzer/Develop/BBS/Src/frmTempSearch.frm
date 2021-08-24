VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTempSearch 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Templete Edit"
   ClientHeight    =   6630
   ClientLeft      =   450
   ClientTop       =   4335
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11715
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraTemplete 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6680
      Left            =   0
      TabIndex        =   6
      Top             =   -90
      Width           =   11715
      Begin VB.TextBox txtComment 
         Height          =   4995
         Left            =   4140
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   15
         Top             =   540
         Width           =   7395
      End
      Begin VB.Frame fraReason 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   465
         Left            =   5040
         TabIndex        =   13
         Top             =   5580
         Visible         =   0   'False
         Width           =   6495
         Begin VB.ComboBox cmbReason 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmTempSearch.frx":0000
            Left            =   2450
            List            =   "frmTempSearch.frx":000A
            TabIndex        =   2
            Top             =   90
            Width           =   4065
         End
         Begin VB.Label lblReason 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "Reason of Modification : "
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   14
            Top             =   180
            Width           =   2220
         End
      End
      Begin VB.TextBox txtTemplete 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1230
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   8
         Top             =   4680
         Width           =   3475
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "&Save"
         Height          =   460
         Left            =   7785
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   6070
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "E&xit"
         Height          =   460
         Left            =   10335
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   6070
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Cancel          =   -1  'True
         Caption         =   "&Clear"
         Height          =   460
         Left            =   9060
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   6070
         Width           =   1215
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         TabIndex        =   1
         Top             =   2880
         Width           =   340
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   4095
         Left            =   165
         TabIndex        =   0
         Top             =   540
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16056319
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "템플릿 명"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblCode 
         BorderStyle     =   1  '단일 고정
         Height          =   285
         Left            =   990
         TabIndex        =   12
         Top             =   5490
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  '단일 고정
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   5970
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblName 
         BackStyle       =   0  '투명
         Caption         =   "Edit Comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4140
         TabIndex        =   10
         Top             =   270
         Width           =   3510
      End
      Begin VB.Label lblTempleteNm 
         BackStyle       =   0  '투명
         Caption         =   "Templete Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   9
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Height          =   240
         Left            =   2625
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmTempSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public button As Integer    '0:Save 1:Exit
Public template As String   '선택한 내용

Private IsFirst As Boolean


Private Sub cmdClear_Click()
   txtComment.Text = ""
   lblCode.Caption = ""
   txtTemplete.Text = ""
End Sub

Private Sub cmdExit_Click()
    button = 1
    Unload Me
End Sub

Private Sub cmdMove_Click()
    If txtComment.Text = "" Then
        txtComment.Text = txtTemplete.Text
    Else
        If Val(medGetP(lblInfo.Caption, 2, "$")) = 1 Then
            txtComment.Text = txtTemplete.Text
        Else
            txtComment.Text = txtComment.Text & vbNewLine & txtTemplete.Text
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    button = 0
    template = txtComment
    Unload Me
End Sub

Private Sub Form_Activate()
    If IsFirst = False Then Exit Sub
    
    IsFirst = True
    txtComment = template
End Sub

Private Sub Form_Load()
    IsFirst = True
    button = 1
    '
    Label4.Caption = App.Path & "(" & App.PrevInstance & ")" & App.hInstance
    '
    txtComment = ""
    txtComment.Height = 5400
    txtTemplete.Text = ""
    
    SetComment
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwList.SortKey = ColumnHeader.Index - 1
    lvwList.Sorted = True
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim objEditABO As clsEditABO
    Dim DrRS As DrRecordSet

    txtTemplete = ""
    
    Set objEditABO = New clsEditABO
    Set DrRS = objEditABO.GetCommentText(Item.Text)
    If Not (DrRS Is Nothing) Then
        With DrRS
            If .RecordCount > 0 Then
                txtTemplete = .Fields("text1")
            End If
        End With
        Set DrRS = Nothing
    End If
    Set objEditABO = Nothing
End Sub



Private Sub SetComment()
    Dim objEditABO As clsEditABO
    Dim DrRS As DrRecordSet
    Dim i As Long
    Dim itmX As ListItem
    
    lvwList.ListItems.Clear
    
    Set objEditABO = New clsEditABO
    Set DrRS = objEditABO.GetComment
    If Not (DrRS Is Nothing) Then
        With DrRS
            If .RecordCount > 0 Then
                For i = 1 To .RecordCount
                    Set itmX = lvwList.ListItems.Add()
                    itmX.Text = .Fields("cdval1")
                    itmX.SubItems(1) = .Fields("field1")
                    .MoveNext
                Next i
            End If
            .RsClose
        End With
        Set DrRS = Nothing
    End If
    
    Set objEditABO = Nothing
End Sub
