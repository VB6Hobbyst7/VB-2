VERSION 5.00
Object = "{F493C2CC-2117-47DA-B779-6610022E1179}#1.0#0"; "TiffViewer.ocx"
Begin VB.Form frmImageView 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "이미지 확인"
   ClientHeight    =   7305
   ClientLeft      =   2400
   ClientTop       =   2100
   ClientWidth     =   11790
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11790
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   210
      TabIndex        =   6
      Top             =   750
      Width           =   2355
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   525
      Left            =   10080
      TabIndex        =   5
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2580
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   2820
      TabIndex        =   3
      Top             =   1230
      Visible         =   0   'False
      Width           =   2985
   End
   Begin TIFFVIEWERLib.TiffViewer TiffViewer1 
      Height          =   6585
      Left            =   2670
      TabIndex        =   1
      Top             =   630
      Width           =   9045
      _Version        =   65536
      _ExtentX        =   15954
      _ExtentY        =   11615
      _StockProps     =   201
   End
   Begin VB.FileListBox flFile 
      Height          =   4410
      Left            =   90
      Pattern         =   "*.gif"
      TabIndex        =   0
      Top             =   2790
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "폴더 목록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   300
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "파일 목록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2340
      Width           =   1155
   End
End
Attribute VB_Name = "frmImageView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    Dim i As Integer
On Err GoTo errCHECK
    
    Dir1.Path = gImagePath
    
    List1.Clear
    'cobas_6500_ResultReport_
    For i = 1 To Dir1.ListCount
        If InStr(1, Dir1.List(i - 1), "cobas_6500_ResultReport_" & GetText(frmInterface.vasID, CInt(frmInterface.lblRowNum), 2)) > 0 Then
            List1.AddItem Replace(Dir1.List(i - 1), gImagePath & "\cobas_6500_ResultReport_", "")
        End If
    Next i
    Exit Sub

errCHECK:
    
End Sub

Private Sub flFile_Click()
    Dim intH As Long
    Dim intW As Long
    
On Err GoTo errCHECK
    intH = 400
    intW = 550
    TiffViewer1.ClearImage
    
    
    TiffViewer1.LoadImage (gImagePath & "\cobas_6500_ResultReport_" & List1.Text & "\" & flFile.FileName)
    Call TiffViewer1.Resize(intW, intH)
    Exit Sub

errCHECK:
End Sub

Private Sub Form_Load()
    
    Command1_Click
    
    intH = 400
    intW = 550
    TiffViewer1.ClearImage
    
    
    
    TiffViewer1.LoadImage ("")
    Call TiffViewer1.Resize(intW, intH)
    
    
    
End Sub

Private Sub List1_Click()
On Err GoTo errCHECK
    flFile.Path = gImagePath & "\cobas_6500_ResultReport_" & List1.Text
    Exit Sub

errCHECK:
End Sub
