VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   5160
   ClientLeft      =   4110
   ClientTop       =   3090
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5535
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00F1F2E3&
      BorderStyle     =   0  '없음
      Height          =   3090
      Left            =   15
      ScaleHeight     =   3090
      ScaleWidth      =   5535
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Ver 1.0.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005E632E&
         Height          =   270
         Left            =   2910
         TabIndex        =   9
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Copyright 1999  Daeryun MTS Co., Ltd."
         Height          =   225
         Left            =   1065
         TabIndex        =   8
         Top             =   1950
         Width           =   3150
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "BBS"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   27.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005E632E&
         Height          =   645
         Left            =   1650
         TabIndex        =   7
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblHomePage 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "http://www.pomis.co.kr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00888F43&
         Height          =   255
         Left            =   1230
         MouseIcon       =   "medAbout.frx":0000
         MousePointer    =   99  '사용자 정의
         TabIndex        =   6
         Top             =   2460
         Width           =   2715
      End
      Begin VB.Label lblNation 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Seoul, Korea"
         Height          =   255
         Left            =   1245
         TabIndex        =   5
         Top             =   2190
         Width           =   2640
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   120
         Picture         =   "medAbout.frx":030A
         Top             =   90
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   840
         Picture         =   "medAbout.frx":3554
         Top             =   330
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F1F2E3&
      Caption         =   "확인(&O)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4110
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   3705
      Width           =   1140
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   795
      Left            =   150
      Shape           =   4  '둥근 사각형
      Top             =   4215
      Width           =   5250
   End
   Begin VB.Label lblAvailMem 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"medAbout.frx":3D43
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   570
      Left            =   345
      TabIndex        =   10
      Top             =   4335
      Width           =   4965
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 이 제품은 다음 사용자에게 사용이 허가되었습니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   3
      Top             =   3330
      Width           =   4995
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "임상병리과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   3825
      Width           =   1140
   End
   Begin VB.Label lblHospital 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "가천의과대학 부속 길병원"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   225
      TabIndex        =   1
      Top             =   3615
      Width           =   2100
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarProductName As String
Private mvarVersion As String
Private mvarHospital As String
Private mvarCopyright As String

'------------------ 속    성 ------------------------'

Public Property Let ProductName(ByVal vData As String)
    mvarProductName = vData
End Property
Public Property Let Version(ByVal vData As String)
    mvarVersion = vData
End Property
Public Property Let Hospital(ByVal vData As String)
    mvarHospital = vData
End Property
Public Property Let Copyright(ByVal vData As String)
    mvarCopyright = vData
End Property
'
'---------------- 메  쏘  드 --------------------'
'
Public Sub SetValues()
   
   lblProductName.Caption = mvarProductName     ' 프로젝트명
   lblVersion.Caption = mvarVersion             ' 버전
   lblHospital.Caption = mvarHospital           ' 병원명
   lblCopyright.Caption = mvarCopyright         ' 저작권
   lblUser.Caption = medGetComNm                ' 컴퓨터(Client)명
    
End Sub

'---------------- 본     문  --------------------'

Private Sub cmdOK_Click()
    Unload Me
    Set frmAbout = Nothing
End Sub

Private Sub Form_Activate()
   
'   Dim tmpTotMem As Long, tmpAvailMem As Long
'
'   Call medSysMem(tmpTotMem, tmpAvailMem)
'   lblTotMem.Caption = Format(tmpTotMem / 1024, "###,###,###") & " KB"
'   lblAvailMem.Caption = Format((tmpAvailMem / tmpTotMem) * 1000, "###") & " %"
'   DoEvents

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHomePage.FontUnderline = False

End Sub

Private Sub lblHomePage_Click()

   Dim i As Double
   Dim MyHomePage As String
   Dim FileName As String
   Dim FileNumber As Integer
   Dim BrowserExec As String * 255
   Dim BrowserExecNm As String

   Dim RetVal As Long

   Call medAlwaysOn(Me, 0)

   MyHomePage = lblHomePage.Caption
   
   BrowserExec = Space(255)

   FileName = App.Path & "\temphtm.HTM"
   FileNumber = FreeFile() ' Get unused file number

   '연결된 브라우저의 Path와 명칭을 알기위해
   '일시적으로 HTML파일을 만든다
   '==> 프로그램의 끝부분에서 삭제한다
   
   Open FileName For Output As #FileNumber
   Write #FileNumber, " " ' Output text
   Close #FileNumber ' Close file

   ' Then find the application associated with it.

   'RetVal = FindExecutable(FileName, Dummy, BrowserExec)
   Call medLoadBrowser(FileName, BrowserExec, RetVal)
   For i = 1 To Len(BrowserExec)
       If Mid(BrowserExec, i, 1) < " " Then
           Mid(BrowserExec, i, 1) = " "
       End If
   Next i

   BrowserExecNm = Trim$(BrowserExec)
   BrowserExecNm = Mid(BrowserExecNm, 1, InStr(1, BrowserExecNm, " ") - 1)
   
   ' If an application is found, launch it!
   If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
       MsgBox "Could not find a browser"
   Else
       i = Shell(BrowserExec & " " & MyHomePage, vbNormalFocus)
   End If
   Kill FileName ' delete temp HTML file

End Sub

Private Sub lblHomePage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHomePage.FontUnderline = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHomePage.FontUnderline = False
End Sub
