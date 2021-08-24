VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Frm_Size 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Szie설정"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "Frm_Size.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "설정"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Begin VB.TextBox Txt_ColorA 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   1035
         Width           =   2085
      End
      Begin VB.TextBox Txt_ColorB 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   4
         Top             =   2010
         Width           =   2085
      End
      Begin VB.CommandButton Cmd_Delete 
         Caption         =   "삭          제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   300
         TabIndex        =   3
         Top             =   4005
         Width           =   2205
      End
      Begin VB.CommandButton Cmd_InSert 
         Caption         =   "추          가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   285
         TabIndex        =   2
         Top             =   3345
         Width           =   2205
      End
      Begin VB.CommandButton Cmd_Close 
         Caption         =   "닫         기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   285
         TabIndex        =   1
         Top             =   5640
         Width           =   2205
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SIZE(약자)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   285
         TabIndex        =   7
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SIZE"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   255
         TabIndex        =   6
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Height          =   6150
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Width           =   2460
      End
   End
   Begin FPSpread.vaSpread Spr_Setting 
      Height          =   6390
      Left            =   0
      TabIndex        =   9
      Top             =   90
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
      SpreadDesigner  =   "Frm_Size.frx":1272
   End
End
Attribute VB_Name = "Frm_Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Close_Click()
 Unload Me
End Sub

'***********************************************************************************
'***  Description   : Color 목록 삭제
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************
Private Sub Cmd_Delete_Click()
 Dim Ls_data(10000) As String
 Dim Li_DataCount As Integer
 Dim Ls_MainData As String
 Dim Ls_Path As String
 Dim Ls_TempData As String
 Dim Li_RowCount As Integer
 Dim Li_RowMaxCount As Integer
 Dim Ls_FileNumber As Integer
  
  With Spr_Setting

      Li_RowCount = 0
      Li_DataCount = 0

      For Li_RowMaxCount = 1 To .MaxRows
          
          Li_RowCount = Li_RowCount + 1
         .Row = Li_RowCount
         .Col = 1
          Ls_data(Li_DataCount) = .Text
          
          If Ls_data(Li_DataCount) = Txt_ColorA.Text Then
           
          Else
           
                Ls_MainData = Ls_MainData & Ls_data(Li_DataCount) & ","
                Li_DataCount = Li_DataCount + 1
           
          End If
           
         .Row = Li_RowCount
         .Col = 2
          Ls_data(Li_DataCount) = .Text
          
          If Ls_data(Li_DataCount) = Txt_ColorB.Text Then
           
          Else
           
                Ls_MainData = Ls_MainData & Ls_data(Li_DataCount) & ","
                Li_DataCount = Li_DataCount + 1
                Ls_Count = Ls_Count + 2
                
           End If
           
      Next Li_RowMaxCount
 
 End With
 
 Ls_FileNumber = FreeFile

 Ls_Path = App.Path & "\Setting\size.ini"

 Open Ls_Path For Output As #Ls_FileNumber

      Print #Ls_FileNumber, Ls_MainData
 
 Close #Ls_FileNumber
 
 Open Ls_Path For Input As #2
      
      While Not EOF(2)
        
         Line Input #2, Ls_TempData
    
      Wend
      
 Close #2

 Li_Count = 0

 LS_Strarry = Split(Ls_TempData, ",")

 cnt = UBound(LS_Strarry)

 If cnt > 0 Then
       
       Do
       
          Li_Count = Li_Count + 1
       
       Loop Until Li_Count = cnt
 End If

 With Spr_Setting
      .MaxRows = (cnt / 2)
       Li_RowCount = 0
       Ls_Count = 0
       Li_RowMaxCount = 0
        
        For Li_RowMaxCount = 1 To .MaxRows
            
            Li_RowCount = Li_RowCount + 1
           .Row = Li_RowCount
           .Col = 1
           .Text = LS_Strarry(Ls_Count)
           
           .Row = Li_RowCount
           .Col = 2
           .Text = LS_Strarry(Ls_Count + 1)
                             
            Ls_Count = Ls_Count + 2
       
        Next Li_RowMaxCount
 
 End With
 
 Txt_ColorA.Text = ""
 Txt_ColorB.Text = ""

End Sub

'***********************************************************************************
'***  Description   : Color 목록 등록
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Cmd_InSert_Click()

 Dim Ls_data(10000) As String
 Dim Li_DataCount As Integer
 Dim Ls_MainData As String
 Dim Li_RowCount As Integer
 Dim Li_RowMaxCount As Integer
 Dim Ls_FileNumber As Integer
 Dim Ls_Path As String
 
 With Spr_Setting

      Li_RowCount = 0

      For Li_RowMaxCount = 1 To .MaxRows
          
          Li_RowCount = Li_RowCount + 1
         .Row = Li_RowCount
         .Col = 1
          
          If Txt_ColorA.Text = .Text Then
            
                MsgBox " 해당 Color를 사용하고 있습니다.", vbInformation, "중복체크"
                Exit Sub
          
          End If
        
      Next Li_RowMaxCount
   
     .MaxRows = .MaxRows + 1
      Li_RowCount = Li_RowCount + 1
   
     .Row = Li_RowCount
     .Col = 1
     .Text = Txt_ColorA.Text
    
     .Row = Li_RowCount
     .Col = 2
     .Text = Txt_ColorB.Text
   
  End With
  
  
  With Spr_Setting

       Li_RowCount = 0
       Li_DataCount = 0
       Li_RowMaxCount = 0

       For Li_RowMaxCount = 1 To .MaxRows
           Li_RowCount = Li_RowCount + 1
           
           .Row = Li_RowCount
           .Col = 1
            Ls_data(Li_DataCount) = .Text
           
            Ls_MainData = Ls_MainData & Ls_data(Li_DataCount) & ","
            Li_DataCount = Li_DataCount + 1
           .Row = Li_RowCount
           .Col = 2
            Ls_data(Li_DataCount) = .Text
           
            Ls_MainData = Ls_MainData & Ls_data(Li_DataCount) & ","
            Li_DataCount = Li_DataCount + 1
            Ls_Count = Ls_Count + 2
       
       Next Li_RowMaxCount
 
  End With
  
  Ls_FileNumber = FreeFile
  Ls_Path = App.Path & "\Setting\Setting.ini"

  Open Ls_Path For Output As #Ls_FileNumber

       Print #Ls_FileNumber, Ls_MainData
  
  Close #Ls_FileNumber


End Sub

'***********************************************************************************
'***  Description   : From 로드
'***  Modification Log : 2006/03/20  김동후  Initial Coding
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
 
 Ls_Path = App.Path & "\Setting\size.ini"

 Open Ls_Path For Input As #2
      
      While Not EOF(2)
           Line Input #2, Ls_TempData
      Wend
 Close #2

 Li_Count = 0

 LS_Strarry = Split(Ls_TempData, ",")

 Ls_StrarryCount = UBound(LS_Strarry)

 If Ls_StrarryCount > 0 Then
    
       Do
          Debug.Print LS_Strarry(Li_Count)
          Li_Count = Li_Count + 1
       
       Loop Until Li_Count = Ls_StrarryCount

 End If

 With Spr_Setting
     .MaxRows = (Ls_StrarryCount / 2)
      Li_RowCount = 0

      For Li_RowMaxCount = 1 To .MaxRows
          Li_RowCount = Li_RowCount + 1
         
         .Row = Li_RowCount
         .Col = 1
         .Text = LS_Strarry(Ls_Count)
          
         .Row = Li_RowCount
         .Col = 2
         .Text = LS_Strarry(Ls_Count + 1)
                            
          Ls_Count = Ls_Count + 2
      
      Next Li_RowMaxCount
 
End With

End Sub

'***********************************************************************************
'***  Description   : 스프레드 클릭 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
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

