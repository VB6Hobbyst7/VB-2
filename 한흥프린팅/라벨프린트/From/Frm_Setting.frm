VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Frm_Setting 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Color????"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "Frm_Setting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   6270
   WindowState     =   2  '?ִ?ȭ
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "????"
      BeginProperty Font 
         Name            =   "????"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   3495
      TabIndex        =   1
      Top             =   75
      Width           =   2715
      Begin VB.CommandButton Cmd_Close 
         Caption         =   "??         ??"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   285
         TabIndex        =   9
         Top             =   5640
         Width           =   2205
      End
      Begin VB.CommandButton Cmd_InSert 
         Caption         =   "??          ??"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   285
         TabIndex        =   8
         Top             =   3345
         Width           =   2205
      End
      Begin VB.CommandButton Cmd_Delete 
         Caption         =   "??          ??"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   300
         TabIndex        =   7
         Top             =   4005
         Width           =   2205
      End
      Begin VB.TextBox Txt_ColorB 
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   6
         Top             =   2010
         Width           =   2085
      End
      Begin VB.TextBox Txt_ColorA 
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   4
         Top             =   1035
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "????"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   255
         TabIndex        =   5
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "????(????)"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   285
         TabIndex        =   3
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Height          =   6150
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   2460
      End
   End
   Begin FPSpread.vaSpread Spr_Setting 
      Height          =   6390
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   3240
      _Version        =   393216
      _ExtentX        =   5715
      _ExtentY        =   11271
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "Frm_Setting.frx":1272
   End
End
Attribute VB_Name = "Frm_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Close_Click()
 Unload Me
End Sub


Private Sub Cmd_Delete_Click()
    Dim sqlDoc  As String
    Dim sqlRet  As Long

             sqlDoc = " Delete From ColorList "
    sqlDoc = sqlDoc & "  Where ColorCode = '" & Trim(Txt_ColorA.Text) & "'"
    
    AdoCn_Jet.Execute sqlDoc, sqlRet
    
    Txt_ColorA.Text = ""
    Txt_ColorB.Text = ""
    
    Call LoadColorList

End Sub


Private Sub Cmd_InSert_Click()
    Dim sqlDoc  As String
    Dim sqlRet  As Long
    Dim intCnt As Integer
'    Dim intRow As Integer
'    Dim varTmp As Variant
    
    intCnt = 0
    
    intCnt = chkInsUpData(Trim(Txt_ColorA.Text))
        
    '-- insert
    If intCnt = 0 Then
                 sqlDoc = " Insert into ColorList "
        sqlDoc = sqlDoc & "        (ColorCode,ColorName,Remark)"
        sqlDoc = sqlDoc & "  Values ("
        sqlDoc = sqlDoc & "        '" & Trim(Txt_ColorA.Text) & "',"
        sqlDoc = sqlDoc & "        '" & Trim(Txt_ColorB.Text) & "',"
        sqlDoc = sqlDoc & "        '')"
        
        AdoCn_Jet.Execute sqlDoc, sqlRet
            
    Else
        '-- update
                 sqlDoc = " Update ColorList Set " & vbNewLine
        sqlDoc = sqlDoc & "        ColorName  = '" & Trim(Txt_ColorB.Text) & "'"
        sqlDoc = sqlDoc & "  Where ColorCode  = '" & Trim(Txt_ColorA.Text) & "'"
        
        AdoCn_Jet.Execute sqlDoc, sqlRet
    
    End If
    
    Txt_ColorA.Text = ""
    Txt_ColorB.Text = ""
    
    Call LoadColorList

End Sub


Private Function chkInsUpData(ByVal strColorCode As String) As Integer
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    chkInsUpData = 0
    
             sqlDoc = "Select count(*) as CNT From ColorList "
    sqlDoc = sqlDoc & " Where ColorCode = '" & strColorCode & "'"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        chkInsUpData = Trim$(adoRS("CNT") & "")
    End If
    
    adoRS.Close:    Set adoRS = Nothing
    
End Function




Private Sub LoadColorList()
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim intRow  As Integer
    
    Spr_Setting.MaxRows = 0
    Spr_Setting.RowHeight(-1) = 15
    
             sqlDoc = " Select * From ColorList "
    sqlDoc = sqlDoc & "  Order By ColorCode "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    With Spr_Setting
        intRow = 0
        .MaxRows = adoRS.RecordCount
        Do While Not adoRS.EOF
            intRow = intRow + 1
            .SetText 1, intRow, Trim$(adoRS("ColorCode") & "")
            .SetText 2, intRow, Trim$(adoRS("ColorName") & "")
            
            adoRS.MoveNext
        
        Loop
    End With
    
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

'***********************************************************************************
'***  Description   : From ?ε?
'***  Modification Log : 2006/03/20  ?赿??  Initial Coding
'***********************************************************************************

Private Sub Form_Load()
 Dim Ls_Path As String
 Dim Li_Count As Integer
 Dim Ls_TempData As String
 Dim Ls_StrarryCount As Integer
 Dim Li_RowCount As Integer
 Dim Li_RowMaxCount As Integer
 Dim Ls_FileNumber As Integer

 Me.Width = 6435
 Me.Height = 7185
 
 
    Call DbConnect_Jet
        
    Call LoadColorList
    
' Ls_Path = App.Path & "\Setting\Setting.ini"
'
' Open Ls_Path For Input As #2
'
'      While Not EOF(2)
'           Line Input #2, Ls_TempData
'      Wend
' Close #2
'
' Li_Count = 0
'
' LS_Strarry = Split(Ls_TempData, ",")
'
' Ls_StrarryCount = UBound(LS_Strarry)
'
' If Ls_StrarryCount > 0 Then
'
'       Do
''          Debug.Print LS_Strarry(Li_Count)
'          Li_Count = Li_Count + 1
'
'       Loop Until Li_Count = Ls_StrarryCount
'
' End If
'
' With Spr_Setting
'     .MaxRows = (Ls_StrarryCount / 2)
'      Li_RowCount = 0
'
'      For Li_RowMaxCount = 1 To .MaxRows
'          Li_RowCount = Li_RowCount + 1
'
'         .Row = Li_RowCount
'         .Col = 1
'         .Text = LS_Strarry(Ls_Count)
'
'         .Row = Li_RowCount
'         .Col = 2
'         .Text = LS_Strarry(Ls_Count + 1)
'
'          Ls_Count = Ls_Count + 2
'
'      Next Li_RowMaxCount
'
'End With


End Sub

'***********************************************************************************
'***  Description   : ???????? Ŭ?? ????
'***  Modification Log : 2006/03/20  ?赿??  Initial Coding
'***********************************************************************************

Private Sub Spr_Setting_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 
 Dim i As Integer
    
    If Row <> NewRow And NewRow > 0 Or Col <> NewCol And NewCol > 0 Then
        Me.Spr_Setting.Col = NewCol
        Me.Spr_Setting.Row = NewRow
           
        If Me.Spr_Setting.Col = 1 Or Me.Spr_Setting.Col = 2 Or Me.Spr_Setting.Col = Me.Spr_Setting.MaxCols Then
        
        Else
            Me.Spr_Setting.Row = 0
            Cbo_Size.Text = Me.Spr_Setting.Text
        End If
        
        Me.Spr_Setting.Row = NewRow
        Me.Spr_Setting.Col = 1
        Txt_ColorA.Text = Me.Spr_Setting.Text
        Me.Spr_Setting.Col = 2
        Txt_ColorB.Text = Me.Spr_Setting.Text
    End If
End Sub
