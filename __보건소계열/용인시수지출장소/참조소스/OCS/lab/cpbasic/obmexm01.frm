VERSION 5.00
Begin VB.Form FrmViewOrder 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Å©±â °íÁ¤ ´ëÈ­ »óÀÚ
   Caption         =   "Order ÄÚµåÁ¶È¸"
   ClientHeight    =   6930
   ClientLeft      =   6180
   ClientTop       =   975
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6930
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   1180
      TabIndex        =   26
      Top             =   520
      Width           =   4410
   End
   Begin VB.PictureBox Panel 
      BackColor       =   &H00C0FFFF&
      Height          =   6360
      Index           =   0
      Left            =   20
      ScaleHeight     =   6300
      ScaleWidth      =   1080
      TabIndex        =   27
      Top             =   525
      Width           =   1140
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "ÀüÃ¼"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   26
         Left            =   120
         TabIndex        =   28
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   25
         Left            =   600
         TabIndex        =   25
         Top             =   5520
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   24
         Left            =   600
         TabIndex        =   24
         Top             =   5115
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   23
         Left            =   600
         TabIndex        =   23
         Top             =   4725
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   22
         Left            =   600
         TabIndex        =   22
         Top             =   4320
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   21
         Left            =   600
         TabIndex        =   21
         Top             =   3915
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   20
         Left            =   600
         TabIndex        =   20
         Top             =   3525
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   19
         Left            =   600
         TabIndex        =   19
         Top             =   3120
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   18
         Left            =   600
         TabIndex        =   18
         Top             =   2715
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   17
         Left            =   600
         TabIndex        =   17
         Top             =   2325
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   16
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   15
         Left            =   600
         TabIndex        =   15
         Top             =   1515
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   14
         Left            =   600
         TabIndex        =   14
         Top             =   1125
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   13
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   12
         Left            =   120
         TabIndex        =   12
         Top             =   5520
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   11
         Left            =   120
         TabIndex        =   11
         Top             =   5115
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   4725
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   3915
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   3525
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   2715
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   2325
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1515
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   1125
         Width           =   340
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   340
      End
   End
   Begin VB.PictureBox Panel 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   5535
      TabIndex        =   29
      Top             =   0
      Width           =   5595
      Begin VB.CommandButton CmdEnd 
         Caption         =   "Á¾·á [&X]"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4200
         TabIndex        =   32
         Top             =   0
         Width           =   1185
      End
      Begin VB.ComboBox ComboBun 
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1170
         Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
         TabIndex        =   30
         Top             =   45
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   " °Ë»çÇ×¸ñ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   90
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmViewOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSlipNo(99)           As String
Dim strSearchE(26)          As String
Dim strSearchK(26)          As String
Dim nIndex                  As Integer


Sub Read_Order(ByVal ArgForm, ArgFrom As String, ByVal ArgIndex As Integer)

    Dim GstrSql             As String
    Dim strSql              As String
    Dim i                   As Integer

    

    GoSub Option_Sql_Made

    Exit Sub

'/----------------------------------------------------------------------------------------/

Option_Sql_Made:
    Dim strOrderCode        As String * 10
    Dim strOrderName        As String * 50

    If Trim(ArgForm) <> "ALL" Then
        GstrSql = " WHERE   OrderName Like '" & ArgFrom & "%' "
    Else
        GstrSql = " WHERE   OrderName >= '0' "
    End If
    
    GstrSql = GstrSql & " AND SlipNo  = '" & strSlipNo(ArgIndex) & "' "
'y  GstrSql = GstrSql & " AND GbInput = '1' "                            'RETURN °ªÀÌ 0:ÀÔ·Â¾øÀ½ 1:ÀÔ·ÂÀÖÀ½.


    strSql = ""
    strSql = strSql & " SELECT OrderCode, OrderName  "
    strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_ORDERCODE "
    strSql = strSql & GstrSql
    strSql = strSql & " GROUP BY Seqno, OrderCode, OrderName "
    
    If adoSetOpen(strSql, adoSet) = False Then Exit Sub
    Do Until adoSet.EOF
        strOrderCode = adoSet.Fields("OrderCode").Value & ""
        strOrderName = adoSet.Fields("OrderName").Value & ""

        List1.AddItem strOrderCode & strOrderName
    
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return


End Sub



Private Sub CmdEnd_Click()
    
    Unload Me
    
End Sub

Private Sub CmdSearch_Click(Index As Integer)

    Dim strForm             As String
    Dim strFrom             As String
    Dim strTo               As String
    

    If CmdSearch(Index).Caption = "" Then Exit Sub
    
    List1.Clear
    
    If Index = 26 Then
        strForm = "ALL"
    Else
        strFrom = CmdSearch(Index).Caption & "%"
    End If

    Call Read_Order(strForm, strFrom, nIndex)

End Sub



Private Sub ComboBun_Click()

    nIndex = ComboBun.ListIndex
    List1.Clear
    
End Sub


Private Sub Form_Load()
    
    Dim i               As Integer
    Dim strSql          As String
    
    strSql = ""
    strSql = strSql & " SELECT  OrderName, SlipNo "
    strSql = strSql & " FROM    TW_MIS_OCS.TWOCS_ORDERCODE   "
    strSql = strSql & " WHERE   SlipNo >= '0006'  "
    strSql = strSql & " AND     SlipNo  < 'A'     "
    strSql = strSql & " AND     SeqNo   =  0      "
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    i = 0
    Do Until adoSet.EOF
        ComboBun.AddItem adoSet.Fields("OrderName").Value & ""
        strSlipNo(i) = adoSet.Fields("SlipNo").Value & ""
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    
    ComboBun.ListIndex = 0
    
    strSearchE(0) = "A":    strSearchE(1) = "B":    strSearchE(2) = "C"
    strSearchE(3) = "D":    strSearchE(4) = "E":    strSearchE(5) = "F"
    strSearchE(6) = "G":    strSearchE(7) = "H":    strSearchE(8) = "I"
    strSearchE(9) = "J":    strSearchE(10) = "K":   strSearchE(11) = "L"
    strSearchE(12) = "M":   strSearchE(13) = "N":   strSearchE(14) = "O"
    strSearchE(15) = "P":   strSearchE(16) = "Q":   strSearchE(17) = "R"
    strSearchE(18) = "S":   strSearchE(19) = "T":   strSearchE(20) = "U"
    strSearchE(21) = "V":   strSearchE(22) = "W":   strSearchE(23) = "X"
    strSearchE(24) = "Y":   strSearchE(25) = "Z":   strSearchE(26) = "ÀüÃ¼"
    
    
    For i = 0 To 26
        CmdSearch(i).Caption = strSearchE(i)
    Next i
    
    
End Sub

Private Sub List1_DblClick()

    Dim i                   As Integer

    If List1.ListIndex = -1 Then Exit Sub
    
    frmExamInfo.txtOrderCode.Text = Left$(List1.List(List1.ListIndex), 8)
    frmExamInfo.txtOrdername.Text = Right$(List1.List(List1.ListIndex), 50)
    
    Unload Me
    
    frmExamInfo.txtOrderCode.SetFocus
    SendKeys "{TAB}"
    
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call List1_DblClick

End Sub






