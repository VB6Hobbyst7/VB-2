VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm428SpcStatics 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   75
   ClientWidth     =   10980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11100
   ScaleWidth      =   10980
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00EBF3ED&
      Caption         =   "조 회(&Q)"
      Height          =   510
      Left            =   6575
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00EBF3ED&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "검체상태별 건수 조회"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   945
      Left            =   75
      TabIndex        =   2
      Top             =   300
      Width           =   10770
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "취소"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   9300
         TabIndex        =   17
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   8300
         TabIndex        =   15
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "확인"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   7300
         TabIndex        =   14
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   6300
         TabIndex        =   13
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "진행"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5300
         TabIndex        =   12
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4300
         TabIndex        =   11
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3300
         TabIndex        =   10
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2300
         TabIndex        =   9
         Top             =   360
         Width           =   800
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   360
         Left            =   1365
         TabIndex        =   3
         Top             =   360
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   83558403
         UpDown          =   -1  'True
         CurrentDate     =   37509
         MinDate         =   36526
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   345
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "보 고 년 도"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblCnt 
      Height          =   6810
      Left            =   75
      TabIndex        =   1
      Tag             =   "10114"
      Top             =   1605
      Width           =   10755
      _Version        =   196608
      _ExtentX        =   18971
      _ExtentY        =   12012
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      MaxCols         =   15
      MaxRows         =   21
      OperationMode   =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frm428.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   20
      ScrollBarTrack  =   3
   End
   Begin MedControls1.LisLabel lblPro 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   1260
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "조회결과"
      LeftGab         =   100
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
      _ExtentX        =   1191
      _ExtentY        =   1191
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm428.frx":0A83
   End
End
Attribute VB_Name = "frm428SpcStatics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()
Private TotRow      As Long
Dim Selected_IDX As Integer
Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub
Private Sub cmdQuery_Click()
    Dim FRcvDt As String
    Dim TRcvDt As String
    
    FRcvDt = Format(dtpFrom.Value, "YYYY")
'    TRcvDt = Format(dtpTo.Value, "YYYYMMDD")
    
    Call SP_Setting                         '일자별로 스프레드 셋팅함
    Call SP_Clear
    Call GetWorkAccount(FRcvDt, TRcvDt)
    
End Sub

Private Sub cmdSave_Click()

    Dim strTmp As String
    
    If tblCnt.DataRowCnt = 0 Then Exit Sub

    With tblCnt
        .Row = 0: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = 14
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .DataRowCnt + 1
        tblexcel.MaxCols = 14
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.Col2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With

    DlgSave.InitDir = "C:\My Documents"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "검체별 건수집계"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)

End Sub

'Private Sub dtpTo_Change()
'    dtpFrom.Value = DateAdd("d", -6, dtpTo.Value)
'End Sub

Private Sub Form_Load()
    Call Form_Clear
    Call Workarea_Setting
    Selected_IDX = 2
    Option1(Selected_IDX).Value = True

End Sub

Private Sub Form_Clear()
    'dtpTo.Value = GetSystemDate
    dtpFrom.Value = GetSystemDate  'DateAdd("d", -6, GetSystemDate)
End Sub
Private Sub SP_Clear()
    With tblCnt
        .Row = 1: .Row2 = .MaxRows
        .Col = 2: .Col2 = 14
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
End Sub
Private Sub Workarea_Setting()
    Dim sSQL    As String
    Dim Rs      As Recordset
    Dim ii      As Long
    
    sSQL = "SELECT cdval1,field1 from " & T_LAB032 & " where " & DBW("cdindex=", LC3_WorkArea)
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        With tblCnt
'            .MaxRows = Rs.RecordCount
            Do Until Rs.EOF
                ii = ii + 1
                .Row = ii
                .Col = 1:  .Value = Rs.Fields("field1").Value & ""
                .Col = 15: .Value = Rs.Fields("cdval1").Value & ""
                Rs.MoveNext
            Loop
            TotRow = .DataRowCnt + 2
        End With
        
    End If
    Set Rs = Nothing
End Sub
Private Sub SP_Setting()
    Dim lngStart As Long
    Dim ii       As Integer
    
    'lngStart = CLng(Mid(Format(dtpFrom.Value, "YYYYMMDD"), 7))
    
    With tblCnt
        .Row = 0
        For ii = 2 To 13
            .Col = ii
'            .Value = Mid(Format(DateAdd("d", lngStart, dtpFrom.Value), "YYYYMMDD"), 7) & "일"
            .Value = Format(ii - 1, "0#") & "월"
            lngStart = lngStart + 1
        Next
    End With
End Sub


Private Function GetWorkAccount(ByVal FRcvDt As String, ByVal TRcvDt As String)
    Dim SQL         As String
    Dim Rs          As Recordset
    Dim objDic      As clsDictionary
    Dim ii          As Long
    Dim jj          As Long
    Dim sWorkarea   As String
    Dim sRCVDT      As String
    Dim lngRowCnt   As Long
    Dim lngColCnt   As Long
    Dim SelectIDX   As String
    SelectIDX = "0"
    SelectIDX = IIf(Selected_IDX = 0, "0,1,2,3,4,5,6,7", Selected_IDX)
'
    SQL = " SELECT substr(a.rcvdt,1,6) as rcvYmon ,a.workarea, Count(*) AS Cnt from " & T_LAB032 & " b, " & T_LAB201 & " a " & _
          " where substr(a.rcvdt,1,4)='" & FRcvDt & "' " & _
          " and " & DBW("b.cdindex=", LC3_WorkArea) & _
          " and a.stscd in ( " & SelectIDX & " ) and a.workarea=b.cdval1 "
          
    SQL = SQL & "GROUP BY substr(a.rcvdt,1,6), a.workarea order by substr(a.rcvdt,1,6), a.workarea"
    
    Set Rs = New Recordset
    Rs.Open SQL, DBConn
    
    If Not Rs.EOF Then
    
        Dim objPro  As jProgressBar.clsProgress
        Dim kk      As Long             'progress 진행 변수
        
        'ProgressBar 처리
        Set objPro = Nothing
        Set objPro = New jProgressBar.clsProgress
        With objPro
            .Container = Me
            .Left = lblPro.Left
            .Top = lblPro.Top
            .Width = lblPro.Width
            .Height = lblPro.Height
            .Max = Rs.RecordCount + (tblCnt.DataRowCnt * 7)
            .Message = "검색중입니다..."
            
'            .Choice = True
'            .SetMyForm Me
'            .XPos = lblPro.Left
'            .YPos = lblPro.Top
'            .XWidth = lblPro.Width
'            .YHeight = lblPro.Height
'            .Appearance = aPlate
'            .Msg = "검색중입니다..."
'            .Max = Rs.RecordCount + (tblCnt.DataRowCnt * 7)
        End With
    
        Set objDic = New clsDictionary
        
        objDic.Clear
        objDic.FieldInialize "rcvYmon,workarea", "cnt"
        objDic.Sort = False
        
        Do Until Rs.EOF
            If objDic.Exists(Mid(Format(Rs.Fields("rcvYmon").Value & ""), 5) & COL_DIV & Rs.Fields("workarea").Value & "") = False Then
                objDic.AddNew Mid(Format(Rs.Fields("rcvYmon").Value & ""), 5) & COL_DIV & Rs.Fields("workarea").Value & "", Rs.Fields("cnt").Value & ""
            End If
            kk = kk + 1
            objPro.Value = kk
            
            Rs.MoveNext
        Loop
        'col
        
        
        With tblCnt
            '화면 Display
            For ii = 1 To tblCnt.DataRowCnt
                .Row = ii
                .Col = 15
                sWorkarea = Trim(.Value)
                For jj = 2 To 13
                    .Row = 0: .Col = jj:   sRCVDT = medGetP(Trim(.Value), 1, "월")
                    .Row = ii: .Col = jj
                    If objDic.Exists(sRCVDT & COL_DIV & sWorkarea) Then
                        objDic.KeyChange sRCVDT & COL_DIV & sWorkarea
                        .Value = Val(objDic.Fields("cnt"))
                    Else
                        .Value = ""
                    End If
                    kk = kk + 1
                    objPro.Value = kk
                    
                Next jj
            Next ii
            
            'Workarea별 합계
            For ii = 1 To .DataRowCnt
                .Row = ii
                For jj = 2 To 13
                    .Col = jj
                    lngRowCnt = lngRowCnt + Val(.Value)
                    
                Next jj
                .Col = 14:
                .Value = lngRowCnt
                If .Value = 0 Then .Value = ""
                lngRowCnt = 0
                
            Next ii
            
            .Row = TotRow
            .Col = 1: .Value = " 합  계 "
            '월별 합계
            For ii = 2 To 14
                .Col = ii
                For jj = 1 To .DataRowCnt
                    .Row = jj: .Col = ii
                    lngColCnt = lngColCnt + Val(.Value)
                    Debug.Print ii - 1 & "-" & jj & "->" & Val(.Value) & "=" & lngColCnt
                Next jj
                .Row = TotRow

                .Value = lngColCnt:
                If .Value = 0 Then .Value = ""
                lngColCnt = 0
            Next ii
        End With
        Set objDic = Nothing
        Set objPro = Nothing
    Else
        MsgBox "조회된 목록이 없습니다.", vbInformation + vbOKOnly
    End If
    
    Set Rs = Nothing
     
End Function

Private Sub Option1_Click(Index As Integer)
    Dim ii As Integer
    For ii = 0 To 7
        If ii <> Index Then Option1(ii).Value = False
    Next
    Selected_IDX = Index
End Sub
Private Sub tblCnt_Click(ByVal Col As Long, ByVal Row As Long)
    Dim SQL         As String
    Dim Rs          As Recordset
    Dim strMask As String
    Dim strPtId     As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strStatus   As String
    Dim numCount    As Integer
    Static iSortOrder As Integer

    With tblCnt
        If Row > .DataRowCnt - 2 Or Row < 1 Or Col > 13 Or Col < 2 Then Exit Sub
        .Row = 0: .Col = Col:   strAccDt = Format(dtpFrom, "yyyy") & Mid(.Value, 1, 2)
        .Row = Row: .Col = 15:  strWorkArea = .Value
        .Col = Col: numCount = Val(.Value): If numCount < 1 Then Exit Sub

        SQL = " SELECT a.spcyy, a.spcno, a.ptid, a.ageday, a.sex, a.deptcd, a.orddoct, a.majdoct, a.workarea,a.accdt,a.accseq,a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, a.entdt,a.enttm,a.entid, a.spccd from " & T_LAB032 & " b, " & T_LAB201 & " a " & _
              " where substr(a.rcvdt,1,6)='" & strAccDt & "' and a.workarea = '" & strWorkArea & "'" & _
              " and " & DBW("b.cdindex=", LC3_WorkArea) & _
              " and a.stscd in ( " & Selected_IDX & " ) and a.workarea=b.cdval1 "
          
        SQL = SQL & " order by a.rcvdt, a.workarea"

        Set Rs = New Recordset
        Rs.Open SQL, DBConn
        strStatus = "환자 ID   검체번호         검체코드  DEPTCD                   " & vbCrLf
        Do Until Rs.EOF
            strStatus = strStatus & Rs.Fields("ptid").Value & " " & Rs.Fields("spcyy").Value & "" & Format(Rs.Fields("spcno").Value & "", "0#########") & "  " & Rs.Fields("spccd").Value & "           " & Rs.Fields("deptcd").Value & " " & "              " & vbCrLf
            
            Rs.MoveNext
        Loop

        Set Rs = Nothing
        Call MouseDefault
    End With
    
    If strStatus <> "" Then
        MsgBox strStatus, , strAccDt & "의 검체 상세내용 "
    End If
End Sub

