VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmTempSearch 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "Templete Edit"
   ClientHeight    =   5295
   ClientLeft      =   450
   ClientTop       =   4335
   ClientWidth     =   8460
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
   ScaleHeight     =   5295
   ScaleWidth      =   8460
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin MSComctlLib.ListView lvwList 
      Height          =   3120
      Left            =   150
      TabIndex        =   8
      Top             =   465
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   5503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ÄÚµå"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ÅÛÇÃ¸´ ¸í"
         Object.Width           =   3572
      EndProperty
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00C0E0FF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   5
      Top             =   1620
      Width           =   340
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Cancel          =   -1  'True
      Caption         =   "È­¸éÁö¿ò(&C)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5460
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Á¾·á(&X)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6735
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "µî·Ï(&S)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4185
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtTemplete 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '¾ç¹æÇâ
      TabIndex        =   1
      Top             =   3975
      Width           =   3255
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   3915
      MultiLine       =   -1  'True
      ScrollBars      =   3  '¾ç¹æÇâ
      TabIndex        =   0
      Top             =   465
      Width           =   4380
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   3900
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   150
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Templete Edit"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   150
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Templete ¸®½ºÆ®"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   135
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3660
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Templete ¹Ì¸®º¸±â"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmTempSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Selected(ByVal vSelectedComment As String)

Private mvarCurrentComment As String

Public Property Let CurrentComment(ByVal vData As String)
    mvarCurrentComment = vData
End Property

Private Sub cmdClear_Click()
   txtComment.Text = ""
   txtTemplete.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click()
    If lvwList.ListItems.Count = 0 Then Exit Sub
    If lvwList.SelectedItem Is Nothing Then Exit Sub
    
    txtComment.Text = txtComment.Text & lvwList.SelectedItem.SubItems(1) & vbNewLine
End Sub

Private Sub cmdSave_Click()
    RaiseEvent Selected(txtComment.Text)
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadComments
    If mvarCurrentComment <> "" Then txtComment.Text = mvarCurrentComment & vbNewLine
End Sub

Private Sub LoadComments()
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    Dim itmX As ListItem
    
    Set objABOSql = New clsABOSql
    Set Rs = New Recordset
    Set Rs = objABOSql.LoadComment
    
    lvwList.ListItems.Clear
    Do Until Rs.EOF
        Set itmX = lvwList.ListItems.Add()
        itmX.Text = Rs.Fields("cdval1").Value & ""
        itmX.SubItems(1) = Rs.Fields("text1").Value & ""
        
        Rs.MoveNext
    Loop
    
    If lvwList.ListItems.Count > 0 Then
        Call lvwList_ColumnClick(lvwList.ColumnHeaders(1))
    End If
    
    Set Rs = Nothing
    Set objABOSql = Nothing
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnToggle() As Boolean
    Static blnFirst As Boolean
    Dim i As Long
    
    If blnFirst = False Then
        ReDim blnToggle(lvwList.ColumnHeaders.Count - 1)
        blnFirst = True
    End If
    
    '¡ã¡å
    
    For i = 1 To lvwList.ColumnHeaders.Count
        lvwList.ColumnHeaders(i).Text = Trim(Replace(lvwList.ColumnHeaders(i).Text, "¡ã", ""))
        lvwList.ColumnHeaders(i).Text = Trim(Replace(lvwList.ColumnHeaders(i).Text, "¡å", ""))
    Next
    
    With lvwList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(blnToggle(ColumnHeader.Index - 1), lvwDescending, lvwAscending)
        .Sorted = True
        
        ColumnHeader.Text = ColumnHeader.Text & " " & IIf(.SortOrder = lvwAscending, "¡ã", "¡å")
        
        blnToggle(ColumnHeader.Index - 1) = IIf(blnToggle(ColumnHeader.Index - 1), False, True)
    End With
    
    If lvwList.ListItems.Count <> 0 Then
        lvwList.ListItems(1).Selected = True
        lvwList.ListItems(1).EnsureVisible
    End If
End Sub

Private Sub lvwList_DblClick()
    If lvwList.ListItems.Count = 0 Then Exit Sub
    If lvwList.SelectedItem Is Nothing Then Exit Sub
    
    txtComment.Text = txtComment.Text & lvwList.SelectedItem.SubItems(1) & vbNewLine
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item Is Nothing Then Exit Sub
    
    txtTemplete.Text = Item.SubItems(1)
End Sub
