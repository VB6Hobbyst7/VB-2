VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmESignAdd 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "전자서명 이미지 등록"
   ClientHeight    =   6495
   ClientLeft      =   210
   ClientTop       =   1200
   ClientWidth     =   7110
   Icon            =   "frmESignAdd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleMode       =   0  '사용자
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdInsert 
      BackColor       =   &H00DBE6E6&
      Caption         =   "등록(&A)"
      Height          =   450
      Left            =   2145
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00DBE6E6&
      Caption         =   "취소(&C)"
      Height          =   450
      Left            =   3600
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   5880
      Width           =   1335
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
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   2415
   End
   Begin VB.DirListBox dirList 
      BackColor       =   &H00FFFBF2&
      Height          =   2190
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   2415
   End
   Begin VB.FileListBox flsFiles 
      BackColor       =   &H00F5F4FF&
      Height          =   2610
      Left            =   2550
      TabIndex        =   0
      Top             =   90
      Width           =   4455
   End
   Begin VB.Frame fraOuter 
      BackColor       =   &H00DBE6E6&
      Caption         =   "이미지 리스트"
      Height          =   2175
      Left            =   90
      TabIndex        =   5
      Top             =   2730
      Width           =   6915
      Begin VB.Frame fraBar 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   1815
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   6765
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   165
            Left            =   75
            TabIndex        =   11
            Top             =   900
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblFileName 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00DBE6E6&
            ForeColor       =   &H80000011&
            Height          =   225
            Left            =   1560
            TabIndex        =   20
            Top             =   660
            Width           =   3465
         End
         Begin VB.Label Label1 
            BackColor       =   &H00DBE6E6&
            Caption         =   "잠시만 기다려 주십시요."
            ForeColor       =   &H80000011&
            Height          =   165
            Left            =   2160
            TabIndex        =   19
            Top             =   1140
            Width           =   2085
         End
         Begin VB.Label lblBar 
            BackColor       =   &H00DBE6E6&
            Caption         =   "이미지 로딩..."
            ForeColor       =   &H80000011&
            Height          =   165
            Left            =   2610
            TabIndex        =   10
            Top             =   375
            Width           =   1185
         End
      End
      Begin VB.Frame fraInner 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   6675
         Begin VB.PictureBox pctContainer 
            BackColor       =   &H00DBE6E6&
            Height          =   1455
            Left            =   0
            ScaleHeight     =   1395
            ScaleWidth      =   6570
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
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
            Top             =   1530
            Width           =   6630
         End
         Begin VB.Label lblIMsg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "이미지가 없습니다."
            ForeColor       =   &H8000000C&
            Height          =   165
            Left            =   2400
            TabIndex        =   21
            Top             =   540
            Width           =   1605
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   90
      TabIndex        =   12
      Top             =   4860
      Width           =   6915
      Begin VB.Label lblPicName 
         BackColor       =   &H00DBE6E6&
         Height          =   210
         Left            =   1260
         TabIndex        =   18
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "선택  이미지  : "
         Height          =   210
         Left            =   60
         TabIndex        =   17
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DBE6E6&
         Caption         =   "이미지 경로   : "
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblimgPath 
         BackColor       =   &H00DBE6E6&
         Height          =   210
         Left            =   1320
         TabIndex        =   15
         Top             =   450
         Width           =   5475
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전자서명자    :"
         Height          =   210
         Left            =   60
         TabIndex        =   14
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label lblSignNm 
         BackColor       =   &H00DBE6E6&
         Height          =   210
         Left            =   1380
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmESignAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSel             As Long
Private mlngSelLast         As Long
Private Const FORM_HEIGHT   As Long = 7095
Private Const INTEGER_MAX   As Integer = 32767

Private blnFirst As Boolean

Public Event ImageAddLoad()
Public Event ImageAdd(ByVal ImgFileName As String)
Private Sub cmdClose_Click()
    '

    Unload Me
    '
End Sub

Private Sub cmdInsert_Click()
Dim lngRet As Long
   '
   If lblPicName.Caption = "" Or lblPicName.Caption = "없음" Then Exit Sub
   lngRet = MsgBox("전자서명 이미지 : " & lblPicName.Caption & "을 전자서명 이미지 출력파일로 등록하시겠읍니까?", _
      vbYesNo + vbQuestion, "슬라이드 등록 확인")
   '
   If lngRet = vbYes Then
      RaiseEvent ImageAdd(lblimgPath & "\" & lblPicName)
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
    Me.MousePointer = 13
    DoEvents
    strPath = dirList.Path
    '
    
    fraBar.Visible = True
    pctContainer.Visible = False
    DoEvents
    
    If Len(strPath) = 0 Then
        GoTo Event_End
    End If
    '
    Set dctImages = FillDictionary(strPath)
    '
    lngFileCount = dctImages.Count
    
    If lngFileCount = 0 Then
        GoTo Event_End
    End If
    
    prgBar.Max = lngFileCount * 4
    prgBar.Min = 0
    
    '
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
        If prgBar.Value >= prgBar.Max Then
            prgBar.Value = 0
        End If
    Next
    On Error GoTo 0
    
    DoEvents
    
    lngImgCount = lngImgCount - 1
    For lngCount = 0 To lngImgCount - 1
        img(lngCount + 1).Left = img(lngCount).Left + img(lngCount).Width + 100
        shpBorder(lngCount + 1).Left = img(lngCount + 1).Left - 4
        prgBar.Value = prgBar.Value + 1
        If prgBar.Value >= prgBar.Max Then
            prgBar.Value = 0
        End If
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
    If lngImgCount > 0 Then
        fraBar.Visible = False
        pctContainer.Visible = True
        scrHorizontal.Enabled = True
        lblIMsg.Visible = False
    Else
        fraBar.Visible = False
        pctContainer.Visible = False
        scrHorizontal.Enabled = False
        lblIMsg.Visible = True
        lblimgPath = ""
        lblPicName = ""
    End If
    Me.MousePointer = 1
    DoEvents
    Exit Sub

End Sub

Private Sub dirList_Change()
    '
    drvDrive.Drive = ParsePath(CStr(dirList.Path), DRIVE_ONLY)
    flsFiles.Path = dirList.Path
    lblimgPath.Caption = dirList.Path
    ImgLoad
    '
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
               lblimgPath.Caption = dirList.Path
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
      'fraBar.Visible = False
      'pctContainer.Visible = True
      '
   End If
   '
End Sub

Private Sub Form_Load()
Dim strPath As String
    On Error GoTo Event_Err
    strPath = ""
    dirList.Path = App.Path & "\"                                      ' SLIDE_CLIENT_PATH
    lblimgPath.Caption = dirList.Path
    flsFiles.Pattern = "*.bmp;*.gif;*.jpg"
    mlngSelLast = 0
    blnFirst = True
    '
Event_End:
    Exit Sub

Event_Err:
    'AddInErr Err
    Resume Event_End
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

Dim N As Integer

ParseFile = ""
For N = Len(pFileName) To 1 Step -1

    If Mid(pFileName, N, 1) = "\" Then
        ParseFile = Mid(pFileName, N + 1, Len(pFileName))
        N = -1
    End If
    
Next N
    
End Function

Private Function FillDictionary(ByVal strPath As String) As Scripting.Dictionary
    Dim fsoSysObj As Scripting.FileSystemObject
    Dim fdrFolder As Scripting.Folder
    Dim filFile As Scripting.File
    Dim dctImages As Scripting.Dictionary
    Dim strFile As String
    '
    Set fsoSysObj = New FileSystemObject
    Set fdrFolder = fsoSysObj.GetFolder(strPath)
    Set dctImages = New Scripting.Dictionary
    
    For Each filFile In fdrFolder.Files
        strFile = ParsePath(filFile.Path, FILEEXT_ONLY)
        Select Case strFile

            Case "jpg"
                dctImages.Add filFile.Path, filFile.Name
                
            Case Else
                
        End Select
        lblFileName = "(" & ParsePath(filFile.Path, FILE_ONLY) & ")"
        DoEvents
    Next
    lblFileName = ""
    Set FillDictionary = dctImages
    '
End Function



