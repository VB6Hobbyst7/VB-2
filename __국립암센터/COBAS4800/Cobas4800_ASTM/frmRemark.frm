VERSION 5.00
Begin VB.Form frmRemark 
   Caption         =   "Comment 등록"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   18645
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame1 
      Height          =   8085
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   18375
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   15450
         TabIndex        =   6
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txtCmntND 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6180
         Left            =   9240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   750
         Width           =   8955
      End
      Begin VB.TextBox txtCmntMD 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6180
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   750
         Width           =   8955
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   16860
         TabIndex        =   1
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblNegInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Mutation detected 코멘트]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   9270
         TabIndex        =   8
         Top             =   7080
         Width           =   8895
      End
      Begin VB.Label lblPosInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Mutation detected 코멘트]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   7050
         Width           =   8895
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Not detected 코멘트]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   9300
         TabIndex        =   3
         Top             =   450
         Width           =   2160
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Mutation detected 코멘트]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   480
         Width           =   2685
      End
   End
End
Attribute VB_Name = "frmRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim FilNum

    FilNum = FreeFile

    Open App.Path & "\eGFR_MD코멘트.txt" For Output As FilNum
    Print #FilNum, txtCmntMD.Text
    Close FilNum
    
    FilNum = FreeFile

    Open App.Path & "\eGFR_ND코멘트.txt" For Output As FilNum
    Print #FilNum, txtCmntND.Text
    Close FilNum
    
    MsgBox "코멘트가 변경되었습니다", vbOKOnly + vbInformation, Me.Caption
    
End Sub

Private Sub Form_Load()
    Dim TextLine
    Dim strBuffer
    Dim strInfo     As String
    
    ClearText
    
    Open App.Path & "\eGFR_MD코멘트.txt" For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine & vbNewLine
    Loop
    
    Close #1 ' 파일을 닫습니다

    txtCmntMD.Text = strBuffer

    strBuffer = ""

    Open App.Path & "\eGFR_ND코멘트.txt" For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine & vbNewLine
    Loop
    
    Close #1 ' 파일을 닫습니다

    txtCmntND.Text = strBuffer
    
    strInfo = "[Mutation detected 코멘트 수정시 주의사항]" & vbNewLine
    strInfo = strInfo & "XXXXX mutation이 발견되었습니다." & vbNewLine
    strInfo = strInfo & "검사방법:" & vbNewLine
    strInfo = strInfo & "- 보 고 자:" & vbNewLine
    
    lblPosInfo.Caption = strInfo

    strInfo = "[Mutation detected 코멘트 수정시 주의사항]" & vbNewLine
    strInfo = strInfo & "  1. XXXXX mutation이 발견되었습니다." & vbNewLine
    strInfo = strInfo & "  2. 검사방법:" & vbNewLine
    strInfo = strInfo & "  3. - 보 고 자:" & vbNewLine
    strInfo = strInfo & "  위의 세줄은 본문에서 변경하지 마세요"
    lblPosInfo.Caption = strInfo

    strInfo = "[Not detected 코멘트 수정시 주의사항]" & vbNewLine
    strInfo = strInfo & "  1. 검사방법:" & vbNewLine
    strInfo = strInfo & "  2. - 보 고 자:" & vbNewLine
    strInfo = strInfo & "  위의 두줄은 본문에서 변경하지 마세요"
    lblNegInfo.Caption = strInfo

End Sub

Private Sub ClearText()

    txtCmntMD.Text = ""
    txtCmntND.Text = ""
    
End Sub


