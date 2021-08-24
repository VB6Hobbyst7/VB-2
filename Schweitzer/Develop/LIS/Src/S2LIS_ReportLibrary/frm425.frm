VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm425ModifyCnt 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   12
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
      TabIndex        =   11
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   10740
      _ExtentX        =   18944
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
      Caption         =   "접수사유별건수"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   10755
      Begin VB.OptionButton optCon 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수정자별"
         Height          =   285
         Index           =   2
         Left            =   4695
         TabIndex        =   9
         Top             =   390
         Width           =   1425
      End
      Begin VB.OptionButton optCon 
         BackColor       =   &H00DBE6E6&
         Caption         =   "업무부서별"
         Height          =   285
         Index           =   1
         Left            =   3390
         TabIndex        =   8
         Top             =   375
         Width           =   1425
      End
      Begin VB.OptionButton optCon 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수정사유"
         Height          =   285
         Index           =   0
         Left            =   2265
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00EBF3ED&
         Caption         =   "조 회(&Q)"
         Height          =   510
         Left            =   7065
         Style           =   1  '그래픽
         TabIndex        =   2
         Tag             =   "0"
         Top             =   255
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpdate 
         Height          =   375
         Left            =   1380
         TabIndex        =   3
         Top             =   300
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   65142787
         UpDown          =   -1  'True
         CurrentDate     =   37509
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   165
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
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
         Caption         =   "연도 선택"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblCnt 
      Height          =   6810
      Left            =   75
      TabIndex        =   4
      Tag             =   "10114"
      Top             =   1620
      Width           =   10740
      _Version        =   196608
      _ExtentX        =   18944
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
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frm425.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   21
   End
   Begin MedControls1.LisLabel lblPro 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   1290
      Width           =   10740
      _ExtentX        =   18944
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
      Left            =   15
      Top             =   -180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   -75
      TabIndex        =   6
      Top             =   -180
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
      SpreadDesigner  =   "frm425.frx":07DC
   End
End
Attribute VB_Name = "frm425ModifyCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'*  수정사유별 건수 통계(사유별,부서별,수정자별)   *
'***************************************************

Option Explicit


Public Event FormClose()
Private TotRow As Long

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub Form_Clear()
    dtpDate.Value = GetSystemDate

End Sub

Private Sub cmdQuery_Click()
    Dim Rs          As Recordset
    Dim objPro      As jProgressBar.clsProgress
    Dim objDic      As New clsDictionary
    Dim sCondition  As String           '조회조건 변수
    Dim sRCVDT      As String
    Dim kk          As Long             'progress 진행 변수
    Dim lngRowCnt   As Long             'Row합계
    Dim lngColCnt   As Long             'Col합계
    Dim ii          As Long
    Dim jj          As Long
    
    
    objDic.Clear
    objDic.FieldInialize "rcvdt,condition", "cnt"
    objDic.Sort = False
    
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    
    With objPro
        .Container = Me
        .Left = lblPro.Left
        .Top = lblPro.Top
        .Width = lblPro.Width
        .Height = lblPro.Height
        .Message = "검색중입니다..."
'        .Choice = True
'        .SetMyForm Me
'        .XPos = lblPro.Left
'        .YPos = lblPro.Top
'        .XWidth = lblPro.Width
'        .YHeight = lblPro.Height
'        .Appearance = aPlate
'        .Msg = "검색중입니다..."
        
    End With
    Set Rs = New Recordset
    Rs.Open sQuery, DBConn
    
    If Not Rs.EOF Then

        objPro.Max = Rs.RecordCount + (tblCnt.DataRowCnt * 12)
        
        Do Until Rs.EOF
            If objDic.Exists((Format(Rs.Fields("mfydt").Value & "")) & COL_DIV & Rs.Fields("condition").Value & "") = False Then
                objDic.AddNew (Format(Rs.Fields("mfydt").Value & "")) & COL_DIV & Rs.Fields("condition").Value & "", Rs.Fields("cnt").Value & ""
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
                sCondition = Trim(.Value)
                For jj = 2 To 13
                    .Row = 0: .Col = jj:   sRCVDT = medGetP(Trim(.Value), 1, "월")
                    .Row = ii: .Col = jj
                    If objDic.Exists(sRCVDT & COL_DIV & sCondition) Then
                        objDic.KeyChange sRCVDT & COL_DIV & sCondition
                        .Value = Val(objDic.Fields("cnt"))
                    
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
                If lngRowCnt = 0 Then
                    .Value = ""
                Else
                    .Value = lngRowCnt
                End If
                
                lngRowCnt = 0
            Next ii
            
            .Row = TotRow
            .Col = 1: .Value = " 합  계 "
            '일자별 합계
            For ii = 2 To 14
                .Col = ii
                For jj = 1 To .DataRowCnt - 2
                    .Row = jj
                    lngColCnt = lngColCnt + Val(.Value)
                Next jj
                .Row = TotRow
                If lngColCnt = 0 Then
                    .Value = ""
                Else
                    .Value = lngColCnt
                End If
                
                lngColCnt = 0
                
            Next ii
        End With
        
    Else
        MsgBox "조회된 목록이 없습니다.", vbInformation + vbOKOnly
    End If
    
    Set Rs = Nothing
    Set objDic = Nothing
    Set objPro = Nothing
    
End Sub
Private Function sQuery() As String
    Dim sRCVDT  As String
    Dim ii      As Long
    
    sRCVDT = Format(dtpDate.Value, "YYYY")
    
    If optCon(0).Value Then
        sQuery = " SELECT " & FUNC_SUBSTR & " (mfydt,5,2) mfydt ,rstcd as condition,count(*) as cnt from " & T_LAB309 & _
                 " WHERE " & DBW("mfydt>=", sRCVDT & "0101") & _
                 " AND   " & DBW("mfydt<=", sRCVDT & "1231") & _
                 " Group by mfydt,rstcd"
    ElseIf optCon(1).Value Then
        sQuery = " SELECT " & FUNC_SUBSTR & " (mfydt,5,2) mfydt ,workarea as condition,count(*) as cnt from " & T_LAB309 & _
                 " WHERE " & DBW("mfydt>=", sRCVDT & "0101") & _
                 " AND   " & DBW("mfydt<=", sRCVDT & "1231") & _
                 " Group by mfydt,workarea"
    Else
        sQuery = " SELECT " & FUNC_SUBSTR & " (mfydt,5,2) mfydt ,mfyid as condition,count(*) as cnt from " & T_LAB309 & _
                 " WHERE " & DBW("mfydt>=", sRCVDT & "0101") & _
                 " AND   " & DBW("mfydt<=", sRCVDT & "1231") & _
                 " Group by mfydt,mfyid"
    End If

End Function
Private Sub Form_Load()
    Call Form_Clear
    Call Reason_Setting
End Sub
Private Sub optCon_Click(Index As Integer)
    Call medClearTable(tblCnt)
    
    Select Case Index
        Case 0: Reason_Setting
        Case 1: Workarea_Setting
        Case 2: Call Mfyid_Setting
    End Select
End Sub
'***************************************************
'*                    수정자별
'***************************************************
Private Sub Mfyid_Setting()
    Dim Rs      As Recordset
    Dim sSQL    As String
    Dim sRCVDT  As String
    Dim ii      As Long
    
    sRCVDT = Format(dtpDate.Value, "YYYY")
    
    sSQL = " SELECT distinct a.mfyid,b.empnm " & _
           " FROM " & T_COM006 & " b,s2lab309 a" & _
           " WHERE " & DBW("a.mfydt>=", sRCVDT & "0101") & _
           " AND   " & DBW("a.mfydt<=", sRCVDT & "1231") & _
           " AND a.mfyid=b.empid"
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        With tblCnt
            .Row = 0: .Col = 1: .Value = "수정자"
            Do Until Rs.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                .Col = 1:  .Value = Rs.Fields("empnm").Value & ""
                .Col = 15: .Value = Rs.Fields("mfyid").Value & ""
                Rs.MoveNext
            Loop
            TotRow = .DataRowCnt + 2
        End With
    End If
    
    Set Rs = Nothing
End Sub
'***************************************************
'*                    업무부서별
'***************************************************
Private Sub Workarea_Setting()
    Dim sSQL    As String
    Dim Rs      As Recordset
    
    
    
    sSQL = "SELECT cdval1,field1 from " & T_LAB032 & " where " & DBW("cdindex=", LC3_WorkArea)
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        With tblCnt
            .Row = 0: .Col = 1: .Value = "업무부서"
            Do Until Rs.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                .Col = 1:  .Value = Rs.Fields("field1").Value & ""
                .Col = 15: .Value = Rs.Fields("cdval1").Value & ""
                Rs.MoveNext
            Loop
            TotRow = .DataRowCnt + 2
        End With
        
    End If
    Set Rs = Nothing
End Sub
'***************************************************
'*                    수정사유
'***************************************************
Private Sub Reason_Setting()
    Dim sSQL    As String
    Dim Rs      As Recordset
    
    sSQL = "SELECT cdval1,text1 from " & T_LAB034 & " where " & DBW("cdindex=", LC4_ModifyReason)
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        With tblCnt
            .Row = 0: .Col = 1: .Value = "수정사유"
            Do Until Rs.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                .Col = 1:  .Value = Rs.Fields("text1").Value & ""
                .Col = 15: .Value = Rs.Fields("cdval1").Value & ""
                Rs.MoveNext
            Loop
            TotRow = .DataRowCnt + 2
        End With
        
    End If
    Set Rs = Nothing
End Sub

Private Sub cmdSave_Click()

    Dim strTmp As String
    
    If tblCnt.DataRowCnt = 0 Then Exit Sub

    With tblCnt
        .Row = 0: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblExcel.MaxRows = .DataRowCnt + 1
        tblExcel.MaxCols = .MaxCols
        tblExcel.Row = 1: tblExcel.Row2 = tblExcel.MaxRows
        tblExcel.Col = 1: tblExcel.Col2 = tblExcel.MaxCols
        tblExcel.BlockMode = True
        tblExcel.Clip = strTmp
        tblExcel.BlockMode = False
    End With

    DlgSave.InitDir = "C:\My Documents"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "수정사유별 건수집계"
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)

End Sub


