VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm230TempSearch 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "Templete Edit"
   ClientHeight    =   9420
   ClientLeft      =   735
   ClientTop       =   960
   ClientWidth     =   12690
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   12690
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
      Height          =   9555
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   12645
      Begin VB.ComboBox cmbReason 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Lis230.frx":0000
         Left            =   7580
         List            =   "Lis230.frx":000A
         TabIndex        =   13
         Top             =   8450
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.TextBox txtTemplete 
         BackColor       =   &H00F1F5F4&
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
         Height          =   2415
         Left            =   5040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '¼öÁ÷
         TabIndex        =   8
         Text            =   "Lis230.frx":0027
         Top             =   6000
         Width           =   7395
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "ÀúÀå(&S)"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8520
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   3
         Top             =   8900
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Á¾·á(&X)"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   11220
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   5
         Top             =   8900
         Width           =   1320
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Cancel          =   -1  'True
         Caption         =   "È­¸éÁö¿ò(&C)"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9900
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   4
         Top             =   8900
         Width           =   1320
      End
      Begin VB.CommandButton cmdMove 
         BackColor       =   &H00CDE7FA&
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
         Height          =   1290
         Left            =   4680
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   1
         Top             =   2160
         Width           =   340
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   5400
         Left            =   5040
         TabIndex        =   2
         Top             =   540
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   9525
         _Version        =   393217
         BackColor       =   15857140
         ScrollBars      =   2
         TextRTF         =   $"Lis230.frx":0039
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   8655
         Left            =   120
         TabIndex        =   0
         Top             =   540
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   15266
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15857140
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
         NumItems        =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   315
         Left            =   5055
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   210
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Edit Comment"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTempleteNm 
         Height          =   315
         Left            =   180
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   210
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Templete Select"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReason 
         Height          =   300
         Left            =   5080
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   8450
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   529
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Reason of Modification"
         Appearance      =   0
      End
      Begin VB.Label lblCode 
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   285
         Left            =   990
         TabIndex        =   10
         Top             =   6720
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Height          =   240
         Left            =   2625
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frm230TempSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event CopyTemplete()

Dim gstrKey As String
Private mvarField1      As String

Public Property Let qField1(ByVal vData As String)
    mvarField1 = vData
End Property

'Hottracking is TRUE
Private Sub cmdClear_Click()
   rtfText.Text = ""
   lblCode.Caption = ""
   txtTemplete.Text = ""
End Sub

Private Sub cmdExit_Click()
   RaiseEvent CopyTemplete
   Call FrameUnlock
   Unload Me
End Sub

Private Sub cmdMove_Click()
   If rtfText.Text = "" Then
      rtfText.Text = txtTemplete.Text
   Else
      If Val(medGetP(lblInfo.Caption, 2, "$")) = 1 Then
         rtfText.Text = txtTemplete.Text
      Else
         rtfText.Text = rtfText.Text & vbNewLine & txtTemplete.Text
      End If
   End If
End Sub

Private Sub cmdSave_Click()
   RaiseEvent CopyTemplete
   Call FrameUnlock
   Unload Me
End Sub

Private Sub Form_Activate()
    Dim objLab034 As New clsComcode034
    Dim aryData() As String
    Dim aryKey() As String
    Dim aryMergy() As String
    Dim strReason As String
    Dim ii As Long
   
   LoadLvwHead
   With objLab034
      Select Case Val(medGetP(lblInfo.Caption, 2, "$"))
         Case 1:
            .LoadTable LC4_Remark
            rtfText.Enabled = False
            rtfText.BackColor = DCM_LightGray
         Case 2:
            .LoadTable LC4_TextResult
            rtfText.Enabled = False
            rtfText.BackColor = vbWhite
         Case 3:
            .LoadTable LC4_FootNote
            rtfText.Enabled = True
            rtfText.BackColor = vbWhite
         Case 4:
            .LoadTable LC4_ModifyReason
            rtfText.Enabled = True
            rtfText.BackColor = vbWhite
      End Select
      
      If .RecordCount > 0 Then
         gstrKey = .GetString("CdVal1")
         aryKey = Split(.GetString("CdVal1"), "$")
         aryData = Split(.GetString("Text1"), "$")
         
         For ii = 0 To UBound(aryKey)
            ReDim Preserve aryMergy(ii)
            aryMergy(ii) = aryKey(ii) & "`" & aryData(ii)
         Next ii
         medDataLoadLvw lvwList, COL_DIV, "`", Join(aryMergy, COL_DIV)
      End If
   End With
   
   Set objLab034 = Nothing
End Sub

Private Sub Form_Load()
   Call CenterForm(Me)
   Label4.Caption = App.Path & "(" & App.PrevInstance & ")" & App.hInstance
   rtfText.Text = ""
   rtfText.Height = 5400
   txtTemplete.Text = ""
End Sub

Private Sub Form_Terminate()
   Call FrameUnlock
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call FrameUnlock
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lvwList.SortKey = ColumnHeader.Index - 1
   lvwList.Sorted = True
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim strTmpRecord As String
   Item.Ghosted = Abs(Item.Ghosted) - 1
   txtTemplete.Text = CStr(Item.SubItems(1))
   lblCode.Caption = CStr(Item.Text)
End Sub

Private Sub FrameUnlock()
'
End Sub

Private Sub LoadLvwHead()
    Dim colHead As ColumnHeader

    medInitLvwHead lvwList, "ÄÚµå,ÅÛÇÃ¸´¸í", "-1500,1200"
End Sub

