VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAPS905 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "소견 결과 조회"
   ClientHeight    =   10200
   ClientLeft      =   7080
   ClientTop       =   735
   ClientWidth     =   8085
   Icon            =   "APS905.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   8085
   Begin VB.Frame fraImageSlide 
      BackColor       =   &H00E8EEEE&
      Height          =   7695
      Left            =   225
      TabIndex        =   18
      Top             =   450
      Visible         =   0   'False
      Width           =   7605
      Begin VB.Image imgImage 
         BorderStyle     =   1  '단일 고정
         Height          =   7500
         Left            =   45
         Picture         =   "APS905.frx":06EA
         Stretch         =   -1  'True
         Top             =   135
         Width           =   7500
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EBF3ED&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   9525
      Width           =   8055
      Begin VB.PictureBox picESign 
         Height          =   500
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   1140
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00DBE6E6&
         Caption         =   "출 력(&P)"
         Height          =   510
         Left            =   5295
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DBE6E6&
         Caption         =   "닫 기(&X)"
         Height          =   510
         Left            =   6645
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   45
         Width           =   1320
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfResultText 
      Height          =   9360
      Left            =   0
      TabIndex        =   2
      Top             =   150
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   16510
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"APS905.frx":16243
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraImage 
      BackColor       =   &H80000009&
      Height          =   9405
      Left            =   30
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   8055
      Begin MSComctlLib.TabStrip tabSlide 
         Height          =   315
         Left            =   15
         TabIndex        =   12
         Top             =   3540
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   556
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S1999-1902A"
               Key             =   "S1999-1902A"
               Object.Tag             =   "1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S1999-1902B"
               Key             =   "S1999-1902B"
               Object.Tag             =   "2"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S1999-1902C"
               Key             =   "S1999-1902C"
               Object.Tag             =   "3"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S1999-1902D"
               Key             =   "S1999-1902D"
               Object.Tag             =   "4"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkMemo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "이미지 메모"
         Height          =   345
         Left            =   6510
         TabIndex        =   17
         Top             =   3540
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame fraOuter 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   3690
         Left            =   30
         TabIndex        =   13
         Top             =   135
         Width           =   7950
         Begin VB.Frame fraInner 
            BorderStyle     =   0  '없음
            Height          =   3615
            Left            =   30
            TabIndex        =   14
            Top             =   120
            Width           =   7890
            Begin VB.PictureBox pctContainer 
               BackColor       =   &H00FFFFF7&
               Height          =   3000
               Left            =   0
               ScaleHeight     =   2940
               ScaleWidth      =   7830
               TabIndex        =   16
               Top             =   15
               Width           =   7890
               Begin VB.Image img 
                  BorderStyle     =   1  '단일 고정
                  Height          =   2700
                  Index           =   0
                  Left            =   210
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   2700
               End
               Begin VB.Shape shpBorder 
                  BorderWidth     =   4
                  Height          =   2775
                  Index           =   0
                  Left            =   180
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   2775
               End
            End
            Begin VB.HScrollBar scrHorizontal 
               Height          =   255
               Left            =   0
               TabIndex        =   15
               Top             =   3015
               Width           =   7875
            End
         End
      End
      Begin VB.TextBox txtMemo 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Left            =   4560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2685
         Width           =   3420
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E8EEEE&
         Height          =   5520
         Left            =   30
         TabIndex        =   6
         Top             =   3795
         Width           =   7995
         Begin RichTextLib.RichTextBox txtRstCmt 
            Height          =   2490
            Left            =   30
            TabIndex        =   7
            Top             =   2205
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   4392
            _Version        =   393217
            BackColor       =   16777207
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"APS905.frx":164F3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtSamCmt 
            Height          =   765
            Left            =   30
            TabIndex        =   8
            Top             =   4710
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   1349
            _Version        =   393217
            BackColor       =   15728382
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"APS905.frx":16598
            MouseIcon       =   "APS905.frx":1663D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FPSpread.vaSpread tblResult 
            Height          =   2070
            Left            =   30
            TabIndex        =   9
            Top             =   120
            Width           =   7935
            _Version        =   196608
            _ExtentX        =   13996
            _ExtentY        =   3651
            _StockProps     =   64
            AllowCellOverflow=   -1  'True
            AutoCalc        =   0   'False
            BackColorStyle  =   3
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridShowHoriz   =   0   'False
            GridSolid       =   0   'False
            MaxCols         =   13
            OperationMode   =   1
            ScrollBars      =   2
            ShadowColor     =   15988216
            ShadowDark      =   12632256
            ShadowText      =   0
            SpreadDesigner  =   "APS905.frx":1679F
            UnitType        =   0
            UserResize      =   0
            VisibleCols     =   8
            VisibleRows     =   22
            TextTip         =   4
         End
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   2520
         Left            =   4560
         TabIndex        =   10
         Top             =   120
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   4445
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmAPS905"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'-------------------------------------
'해부병리/혈액은행 결과조회 여부
'-------------------------------------
'#Const AllowAPSResultReview = True

Public Event Click()
Private objSQL          As clsLISSqlStatement
Private objSlide        As clsSlideImage
Private objDiskFile     As clsDiskFile

Private mvarWorkarea    As String
Private mvarAccdt       As String
Private mvarAccSeq      As String
Private mvarPTid        As String
Private mvarRcvDt       As String
Private mvarTestCd      As String
Private mvarOrdDiv      As String
Private mvarSpecial     As Boolean
Private mvarAllResult   As Boolean

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuSave As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_SAVE& = 1

Private mlngSel         As Long
Private mlngSelLast     As Long
Private Const INTEGER_MAX As Integer = 32767

Public Property Let OrdDiv(ByVal vData As String)
    mvarOrdDiv = vData
End Property

Public Property Let Special(ByVal vData As Boolean)
    mvarSpecial = vData
End Property

Public Property Let AllResult(ByVal vData As Boolean)
    mvarAllResult = vData
End Property

Public Property Get OrdDiv() As String
    OrdDiv = mvarOrdDiv
End Property

'Private Sub mnuSave_Click()
'    Dim strImgDir As String
'
'    If lvwList.ListItems.Count = 0 Then Exit Sub
'
'    strImgDir = lvwList.SelectedItem.SubItems(6)
'
'    DlgSave.InitDir = "C:\"
'    DlgSave.Filter = "JPEG"
'    DlgSave.FileName = Mid(strImgDir, InStrRev(strImgDir, "\", , vbTextCompare) + 1, Len(strImgDir))
'    DlgSave.ShowSave
'
'    FileCopy strImgDir, DlgSave.FileName
'End Sub

Private Sub imgImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Set objPop = New clsPopupMenu
        With objPop
            .AddMenu MENU_SAVE, "IMAGE SAVE"
            .PopupMenus Me.hwnd
        End With
        Set objPop = Nothing
'        Set mnuPopup = frmControls.mnuPopup
'        Set mnuSave = frmControls.mnuSub
'        frmControls.mnuSub1.Visible = False
'        mnuSave.Caption = "Image Save"
'        Me.PopupMenu mnuPopup
'
'        Set mnuSave = Nothing
'        Set mnuPopup = Nothing
'        Unload frmControls
'        Set frmControls = Nothing
    End If
End Sub

Private Sub cmdExit_Click()
    RaiseEvent Click
    Unload Me
End Sub

Private Sub Form_Activate()
    If mvarOrdDiv = lis_orddiv Then
        cmdPrint.Visible = False
    Else
        cmdPrint.Visible = True
    End If
End Sub

Private Sub Form_Load()
    rtfResultText.Text = ""
    Set objSQL = New clsLISSqlStatement
    Set objSlide = New clsSlideImage
    Set objDiskFile = New clsDiskFile
End Sub


Public Function DisplayForm(ByVal pTestCd As String, ByVal pWorkArea As String, _
                        ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean
    Dim strSQL  As String
    Dim RS      As Recordset
    
    fraImage.Visible = True
    If mvarSpecial = True Then
        rtfResultText.Top = 3900
        rtfResultText.Height = 5600
    Else
        rtfResultText.Visible = False
        If mvarAllResult = True Then
            txtRstCmt.Top = tblResult.Top
            txtRstCmt.Height = txtRstCmt.Height + tblResult.Height + 20
        End If
    End If
    
    Set objSQL = New clsLISSqlStatement
    
    Set RS = New Recordset
    RS.Open objSQL.SqlGetGraphDataByLabNo(pTestCd, pWorkArea, pAccDt, pAccSeq, False), DBConn
    
    If RS.RecordCount > 0 Then
        mvarPTid = RS.Fields("ptid").Value & ""
        mvarWorkarea = pWorkArea
        mvarAccdt = pAccDt
        mvarAccSeq = pAccSeq
        mvarTestCd = pTestCd
        mvarRcvDt = RS.Fields("rcvdt").Value & ""
    Else
        mvarPTid = ""
        mvarWorkarea = ""
        mvarAccdt = ""
        mvarAccSeq = ""
        mvarTestCd = ""
        mvarRcvDt = ""
    End If
    RS.Close
    
    If lvwList.ListItems.Count > 0 Then If P_SLIDE_SERVER_PATH = "" Then ClearImage
    
    'DB에서 이미지를 Loading
    If P_SLIDE_SERVER_PATH = "" Then Call LoadImage
    'Client에서 이미지를 Loading
    Call LoadSlide
    
    cmdPrint.Visible = False
    
    Set RS = Nothing
    Set objSQL = Nothing

End Function

Private Sub img_DblClick(Index As Integer)
    imgImage.Picture = img(Index).Picture
    fraImageSlide.Visible = True
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_SAVE
            Dim strImgDir As String
            
            If lvwList.ListItems.Count = 0 Then Exit Sub
            
            strImgDir = lvwList.SelectedItem.SubItems(6)
            
            DlgSave.InitDir = "C:\"
            DlgSave.Filter = "JPEG"
            DlgSave.FileName = Mid(strImgDir, InStrRev(strImgDir, "\", , vbTextCompare) + 1, Len(strImgDir))
            DlgSave.ShowSave
        
            FileCopy strImgDir, DlgSave.FileName
    End Select
End Sub

Private Sub rtfResultText_DblClick()
    Dim Domain      As String
    Dim strDoMain   As String
    Dim strFlag     As String
    Dim lngFCnt     As Integer
    Dim lngURLCnt   As Integer
    Dim strURL(10)  As String
    
    
    lngURLCnt = 0
    strURL(1) = "": strURL(2) = "": strURL(3) = "": strURL(4) = "": strURL(5) = ""
    strURL(6) = "": strURL(7) = "": strURL(8) = "": strURL(9) = "": strURL(10) = ""
    Debug.Print Trim(rtfResultText.Text)
    Domain = Trim(rtfResultText.Text)
    strFlag = "E"
    For lngFCnt = 1 To Len(Domain)
        Select Case Mid(Domain, lngFCnt, 1)
         Case vbCr, vbLf, vbCrLf
              If Mid(Domain, lngFCnt + 1, 4) = "http" Then
                 lngURLCnt = lngURLCnt + 1
                 strFlag = "S"
              Else
                 If lngURLCnt > 0 Then strFlag = "E"
              End If
         Case Else
            If lngURLCnt > 0 And strFlag = "S" Then
               strURL(lngURLCnt) = strURL(lngURLCnt) & Mid(Domain, lngFCnt, 1)
            End If
        End Select
    Next
    
    
    Debug.Print strURL(1)
    Debug.Print strURL(2)
    
    If Len(strURL(1)) > 0 And Mid(strURL(1), 1, 4) = "http" Then
       ShellExecute 1, vbNullString, strURL(1), vbNullString, vbNullString, 1
       Sleep 2000
    End If
    
    If Len(strURL(2)) > 0 And Mid(strURL(2), 1, 4) = "http" Then
       ShellExecute 2, vbNullString, strURL(2), vbNullString, vbNullString, 1
       Sleep 2000
    End If
    
    If Len(strURL(3)) > 0 And Mid(strURL(3), 1, 4) = "http" Then
       ShellExecute 0, vbNullString, strURL(3), vbNullString, vbNullString, 1
       Sleep 2000
    End If
    
    If Len(strURL(4)) > 0 And Mid(strURL(4), 1, 4) = "http" Then
       ShellExecute 0, vbNullString, strURL(4), vbNullString, vbNullString, 1
       Sleep 2000
    End If
    
    If Len(strURL(5)) > 0 And Mid(strURL(5), 1, 4) = "http" Then
       ShellExecute 0, vbNullString, strURL(5), vbNullString, vbNullString, 1
       Sleep 2000
    End If
    
'''    If Len(strURL(6)) > 0 And Mid(strURL(6), 1, 4) = "http" Then ShellExecute 0, vbNullString, strURL(6), vbNullString, vbNullString, 0
'''    If Len(strURL(7)) > 0 And Mid(strURL(7), 1, 4) = "http" Then ShellExecute 0, vbNullString, strURL(7), vbNullString, vbNullString, 0
'''    If Len(strURL(8)) > 0 And Mid(strURL(8), 1, 4) = "http" Then ShellExecute 0, vbNullString, strURL(8), vbNullString, vbNullString, 0
'''    If Len(strURL(9)) > 0 And Mid(strURL(9), 1, 4) = "http" Then ShellExecute 0, vbNullString, strURL(9), vbNullString, vbNullString, 0
'''    If Len(strURL(10)) > 0 And Mid(strURL(10), 1, 4) = "http" Then ShellExecute 0, vbNullString, strURL(10), vbNullString, vbNullString, 0
    
'''    Debug.Print Trim(rtfResultText.Text)
'''    If InStr(Domain, "http") > 0 Then
'''        Domain = Mid(Domain, InStr(Domain, "http"))
'''        'Debug.Print Domain
'''
'''        '==>2014-06-05 PSK 하단의 검사소견 내용까지 같이 있어서 URL연결이 잘못되는경우 발생함
'''        ' VBCR, VBLF, VBCRLF 값을 체크하여 뒷단위 내용 잘라냄
'''        strDoMain = ""
'''        For lngFCnt = 1 To Len(Domain)
'''            Select Case Mid(Domain, lngFCnt, 1)
'''             Case vbCr, vbLf, vbCrLf
'''                  Exit For
'''             Case Else
'''                  strDoMain = strDoMain & Mid(Domain, lngFCnt, 1)
'''            End Select
'''        Next
'''        '<== 2014-06-05 PSK
'''
'''        Debug.Print strDoMain
'''        ShellExecute 0, vbNullString, strDoMain, vbNullString, vbNullString, 1
'''    End If
End Sub

Private Sub scrHorizontal_Change()
    '
    pctContainer.Left = -scrHorizontal.Value
    '
End Sub

Private Sub LoadImage()
    Dim cn As ADODB.Connection, RS As ADODB.Recordset, SQL As String
    Dim Cnt As Long
    Dim ii As Long

    Set cn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    cn.CursorLocation = adUseServer
    
    If mvarOrdDiv = lis_orddiv Then

        cn.Open "Driver={Microsoft ODBC for Oracle};" & _
        "Server=" & GetSetting("Schweitzer2000 LIS", "Server", "DB", "") & ";" & _
        "Uid=" & GetSetting("Schweitzer2000 LIS", "Server", "UID", "") & ";" & _
        "Pwd=" & GetSetting("Schweitzer2000 LIS", "Server", "PWD", "") & ";"
        
            SQL = " SELECT * FROM " & T_LAB310 & _
                  "  where " & DBW("workarea", mvarWorkarea, 2) & _
                  "    and " & DBW("accdt", mvarAccdt, 2) & _
                  "    and " & DBW("accseq", mvarAccSeq, 2) & _
                  "    and " & DBW("testcd", mvarTestCd, 2)
    End If
    RS.Open SQL, cn, adOpenStatic, adLockReadOnly
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            BlobToFile RS!imgfile, RS!imgdir
            RS.MoveNext
        Loop
    End If
    RS.Close
    cn.Close
    
    Set RS = Nothing
    Set cn = Nothing
End Sub

Private Sub LoadSlide()
    Dim ii As Integer
    Dim objListImg As ListImages
'
    If Not (objDiskFile Is Nothing) Then
        Set objDiskFile = Nothing
    End If
    Set objDiskFile = New clsDiskFile
   
    lvwList.ListItems.Clear
    tabSlide.Tabs.Clear
    
    DoEvents
    '
   
    DoEvents
    '
    Set objSlide = New clsSlideImage
    With objSlide
        If mvarOrdDiv = lis_orddiv Then
            If P_SLIDE_SERVER_PATH = "" Then
                .LoadSlide P_SLIDE_DB_PATH, mvarPTid, mvarWorkarea & "-" & Mid(mvarAccdt, 3), mvarAccSeq
            Else
                .LoadSlide P_SLIDE_SERVER_PATH, mvarPTid, mvarWorkarea & "-" & Mid(mvarAccdt, 3), mvarAccSeq
            End If
        End If
    End With
    '
    With objSlide
        medInitLvwHead lvwList, _
             "No,Slide No,Status,File Size,File Date,Description,Directory", _
             "-50,950,300,650,1300,3000,3000"
        .MoveFirst
        For ii = 1 To .RecordCount
        ImageTabAdd ii, True
        .MoveNext
        Next ii
        SlideLoading
    End With
   
End Sub

Private Sub ImageTabAdd(ByVal ii As Integer, ByVal blnFirst As Boolean)
   '
   With objSlide
      tabSlide.Tabs.Add ii, .SlideNo, .SlideNo
      tabSlide.Tabs(ii).Tag = .FileName
      tabSlide.Tabs(ii).Key = .SlideNo
      tabSlide.Tabs(ii).ToolTipText = .FileName
      If ii = 1 Then
         tabSlide_Click
      End If
   End With
   '
End Sub

Private Sub tabSlide_Click()
    Dim ii As Long
    Dim intSelect As Long
    Dim lngWidth As Long
   '
    For ii = 1 To tabSlide.Tabs.Count
        If tabSlide.Tabs(ii).Selected = True Then
            intSelect = ii - 1
            
        End If
    Next ii
   '
    shpBorder(intSelect).Visible = False
    If shpBorder.Count - 1 >= mlngSelLast Then
        shpBorder(mlngSelLast).Visible = False
    End If
   '
    With shpBorder(intSelect)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 4
    End With
    
    mlngSelLast = intSelect
    
    
    If lvwList.ListItems.Count > 0 Then
        lvwList.ListItems(intSelect + 1).Selected = True
        If intSelect = 0 Then
            scrHorizontal.Value = 0
        Else
            If scrHorizontal.Max < img(intSelect).Width + 300 Then
                scrHorizontal.Value = scrHorizontal.Max
            Else
                scrHorizontal.Value = img(intSelect).Width + 300
            End If
        End If
    End If
    
    
End Sub

Private Sub SlideLoading(Optional ByVal SelectedIndex As Long = 1)
    Dim dctImages       As Scripting.Dictionary
    Dim lngCount        As Long
    Dim lngFileCount    As Long
    Dim lngImgCount     As Long
    Dim ii As Long
    '
    lngFileCount = objSlide.RecordCount
    '
    If lngFileCount = 0 Then
        img(0).Visible = False
        GoTo Event_End
    End If
    '
    On Error Resume Next
    For lngCount = 0 To lngFileCount - 1
      If lngCount = 0 Then
         shpBorder(lngCount).Visible = True
         img(lngCount).Visible = True
         shpBorder(lngCount).ZOrder
      Else
        Load img(lngCount)
        img(lngCount).Visible = True
        Load shpBorder(lngCount)
        shpBorder(lngCount).ZOrder
      End If
    Next
    '
    On Error Resume Next
    lngImgCount = 0
    objSlide.MoveFirst
    Do Until objSlide.EOF
      Set img(lngImgCount).Picture = LoadPicture(objSlide.FileName)
      img(lngImgCount).Tag = objSlide.FileName
      If Err = 0 Then
         lngImgCount = lngImgCount + 1
      End If
      objSlide.MoveNext
    Loop
'    '
    On Error GoTo 0
'
    lngImgCount = lngImgCount - 1
    For lngCount = 0 To lngImgCount - 1
        img(lngCount + 1).Left = img(lngCount).Left + _
            img(lngCount).Width + 300
        shpBorder(lngCount + 1).Left = img(lngCount + 1).Left - 4
    Next lngCount
    '
    
    pctContainer.Width = img(lngImgCount).Left + img(lngImgCount).Width + 300
    If pctContainer.Width < 7890 Then pctContainer.Width = 7890
    '
    If pctContainer.Width > INTEGER_MAX Then
        pctContainer.Width = INTEGER_MAX
    End If
    '
    With scrHorizontal
        .SmallChange = CInt(0.1 * pctContainer.Width)
        .LargeChange = CInt(0.2 * pctContainer.Width)
        .Max = CInt(Abs(fraInner.Width - pctContainer.Width))
    End With
    '
    If tabSlide.Tabs.Count > 0 Then
      tabSlide.Tabs(SelectedIndex).Selected = True
      mlngSelLast = SelectedIndex - 1
    Else
      mlngSelLast = 0
    End If
    '
    If objSlide.RecordCount > 0 Then
      If lvwList.ListItems.Count > 0 Then
         lvwList.ListItems.Clear
      End If
      medDataLoadLvw lvwList, vbNewLine, vbTab, objSlide.LvwString
    End If
    '/***
    
    Call lvwList_DblClick
Event_End:
    Exit Sub

End Sub

Private Sub img_Click(Index As Integer)
    Dim lngCount As Long
    mlngSel = Index
    If tabSlide.Tabs.Count = 0 Then Exit Sub
    shpBorder(mlngSelLast).Visible = False
    With shpBorder(Index)
        .Visible = True
        .BorderColor = vbBlack
        .BorderWidth = 4
    End With
    mlngSelLast = Index
    tabSlide.Tabs(Index + 1).Selected = True
    lvwList.ListItems(Index + 1).Selected = True

    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set mnuPopup = Nothing
'    Set mnuSave = Nothing
    Set objSQL = Nothing
    Set objSlide = Nothing
    Set objDiskFile = Nothing
    If P_SLIDE_SERVER_PATH = "" Then ClearImage
End Sub

Private Sub ClearImage()
    Dim ii As Long
    
    If Dir(P_SLIDE_DB_PATH, vbDirectory) = "" Then
        MkDir P_SLIDE_DB_PATH
    Else
        If lvwList.ListItems.Count > 0 Then
            For ii = 1 To lvwList.ListItems.Count
                If Dir(Trim(lvwList.ListItems(ii).SubItems(6))) <> "" Then Kill Trim(lvwList.ListItems(ii).SubItems(6))
            Next
        End If
    End If
End Sub

Private Sub imgImage_DblClick()
    fraImageSlide.Visible = False
End Sub

Private Sub lvwList_DblClick()
    If lvwList.ListItems.Count = 0 Then Exit Sub
    
    txtMemo.Text = Replace(lvwList.SelectedItem.SubItems(5), LINE_DIV, vbNewLine)
    
End Sub


