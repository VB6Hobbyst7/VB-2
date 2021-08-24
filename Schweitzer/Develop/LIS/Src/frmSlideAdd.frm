VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmSlideAdd 
   BackColor       =   &H00DBE6E6&
   Caption         =   "진단검사 의학과 이미지로드"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   1215
   ClientWidth     =   7260
   Icon            =   "frmSlideAdd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleMode       =   0  '사용자
   ScaleWidth      =   7260
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdInsert 
      BackColor       =   &H00DBE6E6&
      Caption         =   "추가(&A)"
      Height          =   510
      Left            =   2205
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   5745
      Width           =   1320
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00DBE6E6&
      Caption         =   "취소(&U)"
      Height          =   510
      Left            =   3540
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   5745
      Width           =   1320
   End
   Begin VB.DriveListBox drvDrive 
      BackColor       =   &H00EEFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   90
      Width           =   2295
   End
   Begin VB.DirListBox dirList 
      BackColor       =   &H00FFFBF2&
      Height          =   1770
      Left            =   210
      TabIndex        =   1
      Top             =   570
      Width           =   2295
   End
   Begin VB.FileListBox flsFiles 
      BackColor       =   &H00F5F4FF&
      Height          =   2250
      Left            =   2610
      TabIndex        =   0
      Top             =   90
      Width           =   4455
   End
   Begin VB.Frame fraOuter 
      BackColor       =   &H00DBE6E6&
      Caption         =   "이미지 리스트"
      Height          =   2175
      Left            =   150
      TabIndex        =   5
      Top             =   2430
      Width           =   6915
      Begin VB.Frame fraBar 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   1815
         Left            =   75
         TabIndex        =   14
         Top             =   225
         Width           =   6765
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   165
            Left            =   75
            TabIndex        =   16
            Top             =   900
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblBar 
            BackColor       =   &H00DBE6E6&
            Caption         =   "슬라이드 이미지 로딩..."
            Height          =   165
            Left            =   2550
            TabIndex        =   15
            Top             =   675
            Width           =   1965
         End
      End
      Begin VB.Frame fraInner 
         BorderStyle     =   0  '없음
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6675
         Begin VB.PictureBox pctContainer 
            Height          =   1455
            Left            =   0
            ScaleHeight     =   1395
            ScaleWidth      =   6570
            TabIndex        =   8
            Top             =   0
            Width           =   6630
            Begin VB.Shape shpBorder 
               BorderWidth     =   4
               Height          =   1250
               Index           =   0
               Left            =   100
               Top             =   100
               Visible         =   0   'False
               Width           =   1250
            End
            Begin VB.Image img 
               BorderStyle     =   1  '단일 고정
               Height          =   1215
               Index           =   0
               Left            =   120
               Stretch         =   -1  'True
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.HScrollBar scrHorizontal 
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   1470
            Width           =   6630
         End
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4650
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "환자   ID"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblReceptNo 
      Height          =   315
      Left            =   150
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5010
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "이미지경로"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   7
      Left            =   3450
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4650
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "성      명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   150
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "선택이미지"
      Appearance      =   0
   End
   Begin VB.Label lblSexAge 
      BackColor       =   &H00DBE6E6&
      Height          =   255
      Left            =   5610
      TabIndex        =   13
      Top             =   4680
      Width           =   630
   End
   Begin VB.Label lblPtNm 
      BackColor       =   &H00DBE6E6&
      Height          =   315
      Left            =   4515
      TabIndex        =   12
      Top             =   4665
      Width           =   1020
   End
   Begin VB.Label lblPtId 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   1260
      TabIndex        =   11
      Top             =   4680
      Width           =   2160
   End
   Begin VB.Label lblimgPath 
      BackColor       =   &H00DBE6E6&
      Height          =   285
      Left            =   1260
      TabIndex        =   10
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label lblPicName 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   1245
      TabIndex        =   9
      Top             =   5400
      Width           =   4950
   End
End
Attribute VB_Name = "frmSlideAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSel     As Long
Private mlngSelLast As Long
Private Const FORM_HEIGHT As Long = 7095
Private Const INTEGER_MAX As Integer = 32767
Public Event ImageAddLoad()
Public Event ImageAdd(ByVal ImgDiv As String, ByVal ImgFileName As String)
Dim blnFirst As Boolean

Private Sub cmdClose_Click()
    '
    Unload Me
    '
End Sub

Private Sub cmdInsert_Click()
    Dim lngRet As Long
    '
    If lblPicName.Caption = "" Or lblPicName.Caption = "없음" Then Exit Sub
    lngRet = MsgBox("이미지 : " & lblPicName.Caption & "을 새로 등록하시겠읍니까?", _
        vbYesNo + vbQuestion, "이미지 등록 확인")
    '
    If lngRet = vbYes Then
'        If optDiv(0).Value = True Then
'            RaiseEvent ImageAdd("G", lblimgPath & "\" & lblPicName)
'        Else
            RaiseEvent ImageAdd("", lblimgPath & "\" & lblPicName)
'        End If
    End If
    '
    Unload Me
   '
End Sub

Private Sub ImgLoad()
Dim dctImages       As Scripting.Dictionary
Dim strPath         As String
Dim lngCount        As Long
Dim lngFileCount    As Long
Dim lngImgCount     As Long
Dim varItem         As Variant
    '
    strPath = dirList.Path
    '
    If Len(strPath) = 0 Then
        GoTo Event_End
    End If
    '
    Set dctImages = FillDictionary(strPath)
    '
    lngFileCount = dctImages.Count
    prgBar.Max = lngFileCount * 4
    prgBar.Min = 0
    
    '
    If lngFileCount = 0 Then
        GoTo Event_End
    End If
    '
    On Error Resume Next
    For lngCount = 0 To lngFileCount - 1
        Load img(lngCount)
        img(lngCount).Visible = True
        Load shpBorder(lngCount)
        shpBorder(lngCount).ZOrder
        prgBar.Value = prgBar.Value + 2
    Next
    '
    On Error Resume Next
    lngImgCount = 0
    For Each varItem In dctImages
        Set img(lngImgCount).Picture = LoadPicture(varItem)
        img(lngImgCount).Tag = ParseFile(varItem)
        If Err = 0 Then
            lngImgCount = lngImgCount + 1
        End If
        prgBar.Value = prgBar.Value + 1
    Next
    On Error GoTo 0

    lngImgCount = lngImgCount - 1
    For lngCount = 0 To lngImgCount - 1
        img(lngCount + 1).Left = img(lngCount).Left + img(lngCount).Width + 100
        shpBorder(lngCount + 1).Left = img(lngCount + 1).Left - 4
        prgBar.Value = prgBar.Value + 1
    Next
    '
    pctContainer.Width = img(lngImgCount).Left + img(lngImgCount).Width
    '
    With scrHorizontal
      .SmallChange = 1
      .LargeChange = 2
      If lngImgCount < 5 Then
         .Max = 1
      Else
         .Max = lngImgCount
      End If
    End With
    '
Event_End:
    Exit Sub

End Sub

Private Sub dirList_Change()
    drvDrive.Drive = ParsePath(dirList.Path, DRIVE_ONLY)
    flsFiles.Path = dirList.Path
    lblimgPath.Caption = dirList.Path
End Sub

Private Sub drvDrive_Change()
    dirList.Path = drvDrive.Drive
    flsFiles.Path = dirList.Path
    lblimgPath.Caption = dirList.Path
End Sub

Private Sub flsFiles_Click()
Dim ii As Long
Dim jj As Long
   '
   For ii = 0 To (flsFiles.ListCount - 1)
      If flsFiles.Selected(ii) = True Then
         For jj = 0 To (img.Count - 1)
            If img(jj).Tag = flsFiles.List(ii) Then
               img_Click (jj)
               lblPicName = img(jj).Tag
            End If
         Next
         ii = flsFiles.ListCount - 1
      End If
   Next
   '
End Sub

Private Sub Form_Activate()
   '
   If blnFirst = True Then
      Me.MousePointer = 13
      DoEvents
      '
      'LockWindowUpdate (Me.hwnd)
      pctContainer.Visible = False
      RaiseEvent ImageAddLoad
      blnFirst = False
      ImgLoad
      'LockWindowUpdate (0&)
      Me.MousePointer = 1
      '
      fraBar.Visible = False
      pctContainer.Visible = True
      '
   End If
   '
End Sub

Private Sub Form_Load()
Dim strPath As String
    On Error GoTo Event_Err
    strPath = ""
    dirList.Path = P_SLIDE_CLIENT_PATH
    lblimgPath.Caption = dirList.Path
    flsFiles.Pattern = "*.bmp;*.gif;*.jpg"
    mlngSelLast = 0
    blnFirst = True
    '
Event_End:
    Exit Sub

Event_Err:
'    AddInErr Err
    Resume Event_End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSlideAdd = Nothing
End Sub

Private Sub img_Click(Index As Integer)
Dim lngCount As Long
Dim ii As Long
    mlngSel = Index
    shpBorder(mlngSelLast).Visible = False
    With shpBorder(Index)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 4
    End With
    mlngSelLast = Index
    For ii = 0 To (flsFiles.ListCount - 1)
      If flsFiles.List(ii) = img(Index).Tag Then
         flsFiles.ListIndex = ii
         lblPicName = flsFiles.List(flsFiles.ListIndex)
         ii = flsFiles.ListCount - 1
      End If
    Next
    '
End Sub


Private Sub scrHorizontal_Change()
    '
    
    With scrHorizontal
      If .Max < 5 Then Exit Sub
      If .Value = 0 Then
         pctContainer.Left = img(0).Left
      ElseIf .Value >= img.Count - 5 Then
         pctContainer.Left = -img((img.Count - 5)).Left
      Else
         pctContainer.Left = -img((.Value)).Left
      End If
    End With
    '
End Sub


Private Function ParseFile(ByVal pFileName) As String
    Dim aryTmp() As String
    Dim strTmp As String
    Dim N As Integer

    ParseFile = ""

    If pFileName = "" Then Exit Function

    aryTmp = Split(pFileName, "\")

    ParseFile = aryTmp(UBound(aryTmp))
    
End Function
