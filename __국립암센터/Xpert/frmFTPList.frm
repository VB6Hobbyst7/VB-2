VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmFTPList 
   Caption         =   "Connection Equipment File"
   ClientHeight    =   8805
   ClientLeft      =   555
   ClientTop       =   855
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11700
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   30
      TabIndex        =   13
      Top             =   8430
      Width           =   11595
      _Version        =   65536
      _ExtentX        =   20452
      _ExtentY        =   556
      _StockProps     =   15
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Label lblFtpLog 
         BackStyle       =   0  '투명
         Caption         =   "test"
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
         Left            =   150
         TabIndex        =   14
         Top             =   60
         Width           =   10335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8475
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   11625
      Begin Xpert_국립암센터.MDButton MDButton1 
         Height          =   435
         Left            =   8430
         TabIndex        =   15
         Top             =   5310
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "초기화"
      End
      Begin Xpert_국립암센터.MDButton cmdFileList 
         Height          =   435
         Left            =   6390
         TabIndex        =   12
         Top             =   5310
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "불러오기"
      End
      Begin Xpert_국립암센터.MDButton cmdClose 
         Height          =   435
         Left            =   10470
         TabIndex        =   11
         Top             =   5310
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "종료"
      End
      Begin Xpert_국립암센터.MDButton cmdResExe 
         Height          =   435
         Left            =   7410
         TabIndex        =   10
         Top             =   5310
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "결과처리"
      End
      Begin Xpert_국립암센터.MDButton cmdFDelete 
         Height          =   435
         Left            =   9450
         TabIndex        =   9
         Top             =   5310
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "파일삭제"
      End
      Begin VB.FileListBox flbData 
         Height          =   1890
         Left            =   180
         TabIndex        =   8
         Top             =   6480
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox txtPath 
         Height          =   270
         Left            =   9210
         TabIndex        =   6
         Top             =   4380
         Visible         =   0   'False
         Width           =   1995
      End
      Begin Threed.SSCommand cmdBackFolder 
         Height          =   495
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "<<"
      End
      Begin VB.Frame Frame2 
         Height          =   2715
         Left            =   6360
         TabIndex        =   3
         Top             =   5700
         Width           =   5175
         Begin VB.TextBox txtEvent 
            Appearance      =   0  '평면
            Height          =   2505
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   4
            Top             =   150
            Width           =   5085
         End
      End
      Begin Xpert_국립암센터.MDButton cmdFileDownLoad 
         Height          =   435
         Left            =   3090
         TabIndex        =   2
         Top             =   7380
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "File Download"
      End
      Begin FPSpread.vaSpread vasFTP 
         Height          =   8205
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   6195
         _Version        =   393216
         _ExtentX        =   10927
         _ExtentY        =   14473
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFTPList.frx":0000
      End
      Begin FPSpread.vaSpread vasLFile 
         Height          =   5085
         Left            =   6360
         TabIndex        =   7
         Top             =   180
         Width           =   5175
         _Version        =   393216
         _ExtentX        =   9128
         _ExtentY        =   8969
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFTPList.frx":5BC7
      End
   End
End
Attribute VB_Name = "frmFTPList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gLocalPath As String

Private MMFTP As New clsFTP
Private MMDirList As New cDirList

'Const gServer = "203.241.227.162"
'Const gPort = "21"
'Const gUser = "lsy05"
'Const gPassWD = "1227"

'Const gServer = "203.241.227.162"
'Const gPort = "21"
'Const gUser = "lsy05"
'Const gPassWD = "1227"

'Const gServer = "203.241.227.20"
'Const gPort = "21"
'Const gUser = "ljs00"
'Const gPassWD = "798520"

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdFDelete_Click()
Dim i As Integer

    For i = 1 To vasLFile.DataRowCnt
        vasLFile.Col = 1
        vasLFile.Row = i
        If vasLFile.Value = 1 Then
            Kill gLocalPath & "\" & Trim(GetText(vasLFile, i, 2))
        End If
    Next
    
    Local_FileList
    
End Sub

Private Sub cmdFileDownLoad_Click()
    Dim sDownLoad As Boolean
    Dim sFileName As String
    Dim sFilePath As String
    Dim sLocalPath As String
    Dim sRow As Long
    Dim i, j As Long
    
    sRow = -1
    For i = 1 To vasFTP.DataRowCnt
        vasFTP.Col = 1
        vasFTP.Row = i
        If vasFTP.Value = 1 Then
            sRow = i
            Exit For
        End If
    Next
    If sRow = -1 Then Exit Sub
    
    sFileName = Trim(GetText(vasFTP, sRow, 2))
    txtPath = Trim(MMFTP.FTPcurdir)
    sFilePath = Trim(txtPath & "/" & sFileName)
    
    sLocalPath = gLocalPath & "\" & sFileName
    
    txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & sFileName & ":Download Start."
    
    sDownLoad = MMFTP.FTPDownloadFile(sLocalPath, sFilePath)
    
    If sDownLoad = True Then
        txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & sFileName & ":Download Succed."
    Else
        txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & sFileName & ":Download Failed."
    End If
    txtEvent.SelStart = Len(txtEvent)
        
End Sub

Private Sub File_Download(asRow As Long)
    Dim sDownLoad As Boolean
    Dim sFileName As String
    Dim sFilePath As String
    Dim sLocalPath As String
    Dim sRow As Long
    Dim i, j As Long
    
    sRow = -1
    sRow = asRow
'    For i = 1 To vasFTP.DataRowCnt
'        vasFTP.Col = 1
'        vasFTP.Row = i
'        If vasFTP.Value = 1 Then
'            sRow = i
'            Exit For
'        End If
'    Next
    If sRow = -1 Then Exit Sub
    
    sFileName = Trim(GetText(vasFTP, sRow, 2))
    If sFileName = "" Then Exit Sub
    
    txtPath = Trim(MMFTP.FTPcurdir)
    sFilePath = Trim(txtPath & "/" & sFileName)
    
    If Dir(App.Path & "\FTP", vbDirectory) = "" Then
        MkDir App.Path & "\FTP"
    End If
    
    sLocalPath = App.Path & "\FTP\" & sFileName
    
    txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & sFileName & ":Download Start."
    
    sDownLoad = MMFTP.FTPDownloadFile(sLocalPath, sFilePath)
    
    If sDownLoad = True Then
        txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & sFileName & ":Download Succed."
    Else
        txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & sFileName & ":Download Failed."
    End If
    txtEvent.SelStart = Len(txtEvent)
        
End Sub


Private Sub cmdFileList_Click()
    Local_FileList
    
End Sub

Private Sub cmdResExe_Click()
    Dim i As Integer
    
    For i = 1 To vasLFile.DataRowCnt
        vasLFile.Col = 1
        vasLFile.Row = i
        If vasLFile.Value = 1 Then
            gResFileName = Trim(GetText(vasLFile, i, 2))
            
            frmInterface.Res_Proc gLocalPath & "\" & Trim(GetText(vasLFile, i, 2))
            
        End If
    Next
    
End Sub

Private Sub Form_Load()
    Dim sConnection As Boolean
    
    sConnection = MMFTP.OpenConnection(gFTPConf.Server, gFTPConf.Port, gFTPConf.User, gFTPConf.Passwd)
    If sConnection = True Then
        txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & "FTP Server Connected."
    Else
        txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & "FTP Server not Connected."
    End If
    
    txtEvent.SelStart = Len(txtEvent.Text)
    ClearSpread vasFTP
    MMFTP.FtpScanDirectory "*.txt"
    
    If Dir(App.Path & "\FTP", vbDirectory) = "" Then
        MkDir App.Path & "\FTP"
    End If
    
    gLocalPath = App.Path & "\FTP"
    
    Local_FileList
    
    lblFtpLog.Caption = "[FTP Server] " & gFTPConf.Server & ":" & gFTPConf.Port & "   [User] " & gFTPConf.User
End Sub

Private Sub FTP_Init()
    Dim sConnection As Boolean
    MMFTP.CloseConnection
    sConnection = MMFTP.OpenConnection(gFTPConf.Server, gFTPConf.Port, gFTPConf.User, gFTPConf.Passwd)
    ClearSpread vasFTP
    MMFTP.FtpScanDirectory
    Local_FileList
    txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & "Initialize Server Connection."
    
End Sub

Private Sub Local_FileList()
    Dim sFileName As String
    Dim i As Long
    
    flbData.Refresh
    ClearSpread vasLFile
    
    flbData.Path = gLocalPath
    
  '작업중=============================================================
    For i = 0 To flbData.ListCount - 1
        SetText vasLFile, flbData.List(i), vasLFile.DataRowCnt + 1, 2
'        If InStr(1, flbData.List(i), "DATE=" & Trim(Text_Today)) > 0 And InStr(1, flbData.List(i), ".ASC") > 0 Then
'            sFileName = flbData.List(i)
'            If (Mid(sFileName, 2, 2) = "ID" Or Mid(sFileName, 1, 2) = "ID") And _
'                Mid(sFileName, InStr(1, sFileName, "DATE") + 5, 10) = Trim(Text_Today) Then
'
'                Load_Allergy sFileName
'            End If
'        End If
    Next
    
    flbData.Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    MMFTP.CloseConnection
    txtEvent = txtEvent & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]" & "FTP Server Disconnected."
    txtEvent.SelStart = Len(txtEvent.Text)
End Sub

Private Sub MDButton1_Click()
    FTP_Init
End Sub

Private Sub vasFTP_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i, j As Long
    
    If Row < 1 Or Row > vasFTP.DataRowCnt Then
        Exit Sub
    End If
    
    If Col = 1 Then
        For i = 1 To vasFTP.DataRowCnt
            If i = Row Then
            Else
                vasFTP.Col = 1
                vasFTP.Row = i
                vasFTP.Value = 0
            End If
        Next
        
    ElseIf Col = 3 Then
        File_Download Row
        Local_FileList
    End If
    
End Sub

Private Sub vasFTP_DblClick(ByVal Col As Long, ByVal Row As Long)
'    MMFTP.FTPcurdir
    Dim sChDir As String
    Dim sDirName As String
    Dim sPath As String
    
    sDirName = Trim(GetText(vasFTP, Row, 2))
    If sDirName = "" Then
        Exit Sub
    End If
    sChDir = MMFTP.FTPchdir("./" & sDirName)
    
    ClearSpread vasFTP
    MMFTP.FtpScanDirectory
    
    txtEvent.Text = txtEvent.Text & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]Root" & MMFTP.FTPcurdir
    txtEvent.SelStart = Len(txtEvent.Text)
End Sub

Private Sub cmdBackFolder_Click()
    Dim sChDir As String
    Dim sDirName As String
    Dim sPath As String
    
    sPath = MMFTP.FTPcurdir
    
    If Trim(sPath) = "" Or Trim(sPath) = "/" Then
        Exit Sub
    End If
    
    sChDir = MMFTP.FTPchdir("../")
    
    ClearSpread vasFTP
    MMFTP.FtpScanDirectory
    
    txtEvent.Text = txtEvent.Text & vbCrLf & "[" & Format(Time, "hh:mm:ss") & "]Root" & MMFTP.FTPcurdir
    txtEvent.SelStart = Len(txtEvent.Text)
End Sub

