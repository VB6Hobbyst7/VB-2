VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBBS925 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "혈액은행 검사통계"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   135
      TabIndex        =   5
      Top             =   285
      Width           =   10740
      Begin VB.CheckBox chkPtidDupcheck 
         BackColor       =   &H00E0E0E0&
         Caption         =   "중복된 환자 표시 안함"
         Height          =   225
         Left            =   6645
         TabIndex        =   12
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9225
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "124"
         Top             =   360
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpTo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   2955
         TabIndex        =   7
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   70189059
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   8
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   70189059
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   270
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   450
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "조회기간"
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2745
         TabIndex        =   10
         Tag             =   "40304"
         Top             =   510
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8190
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   8355
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6840
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8355
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9495
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   8355
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   2850
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
      SpreadDesigner  =   "frmBBS925.frx":0000
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   5070
      Top             =   2550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6780
      Left            =   135
      TabIndex        =   11
      Tag             =   "10114"
      Top             =   1470
      Width           =   10725
      _Version        =   196608
      _ExtentX        =   18918
      _ExtentY        =   11959
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ColsFrozen      =   5
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      MaxCols         =   12
      MaxRows         =   25
      MoveActiveOnFocus=   0   'False
      OperationMode   =   3
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS925.frx":01CC
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   13
   End
End
Attribute VB_Name = "frmBBS925"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objDic As clsDictionary
Private SortTF As Boolean

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblList.DataRowCnt = 0 And tblList.DataRowCnt = 0 Then Exit Sub
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblList.MaxRows + 1
        .MaxCols = tblList.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblList.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "Blood Bank"
    DlgSave.ShowSave

    Me.MousePointer = vbHourglass
    tblexcel.SaveTabFile (DlgSave.FileName)
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    With tblList
    
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .PrintJobName = "혈액은행 검사통계 출력"
        .PrintAbortMsg = "혈액은행 검사통계 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        .PrintHeader = "/n/n/l/fb1 " & "♧ 혈액은행 검사통계 (" & Format(dtpFrom.Value, CS_DateLongFormat) & " 부터 " & _
                                                              Format(dtpTo.Value, CS_DateLongFormat) & " 까지 ) /c/fb1/n"
                                       
        .PrintFooter = " /l " & String(116, Chr(6)) & "/n/l " & HOSPITAL_MAIN & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintMarginLeft = 200
        .PrintMarginRight = 100
        .PrintShadows = False
        .PrintMarginTop = 500
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

        .GridSolid = True
    End With

'    Dim objReport   As clsBBSPrint
'    Dim ii          As Integer
'    Dim jj          As Integer
'    Dim kk          As Integer
'
'    Dim strHeader1 As String
'    Dim strHeader2 As String
'    Dim strHeader3 As String
'    Dim strBody    As String
'
'    If tblList.MaxRows = 0 Then Exit Sub
'    Set objReport = New clsBBSPrint
'
'    Me.MousePointer = 11
'
'    strHeader1 = "Blood Bank"
'
'    strHeader2 = " ♣ 조회일자 : " & Format(dtpFrom, "yyyy-mm-dd") & " ~ " & Format(dtpTo.Value, "yyyy-mm-dd")
'    strHeader2 = strHeader2 & " ♣ 출력일시 : " & Format(Getsystemdate, "YYYY-MM-DD HH:MM") & COL_DIV & "5" & COL_DIV & "1"
'
'    strHeader3 = "번호" & COL_DIV & "5" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "보고일자" & COL_DIV & "15" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "Chart No" & COL_DIV & "35" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "Name" & COL_DIV & "50" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "Sex/Age" & COL_DIV & "65" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "Ward" & COL_DIV & "80" & COL_DIV & "0"
'    jj = 80
'    With tblList
'        For ii = 6 To .MaxCols
'            .Row = 0: .Col = ii
'            If .Col <> .MaxCols Then
'                jj = jj + 20
'                strHeader3 = strHeader3 & vbTab & .Value & COL_DIV & jj & COL_DIV & "0"
'            Else
'                jj = jj + 20
'                strHeader3 = strHeader3 & vbTab & .Value & COL_DIV & jj & COL_DIV & "1"
'            End If
'        Next
'    End With
'
'    With tblList
'        For ii = 1 To .DataRowCnt
'            .Row = ii: jj = 80
'            strBody = strBody & ii & COL_DIV & "5" & COL_DIV & "0" & COL_DIV & "0"
'            .Col = 1
'            strBody = strBody & vbTab & .Value & COL_DIV & "15" & COL_DIV & "0" & COL_DIV & "0"
'            .Col = 2
'            strBody = strBody & vbTab & .Value & COL_DIV & "35" & COL_DIV & "0" & COL_DIV & "0"
'            .Col = 3
'            strBody = strBody & vbTab & .Value & COL_DIV & "50" & COL_DIV & "0" & COL_DIV & "0"
'            .Col = 4
'            strBody = strBody & vbTab & .Value & COL_DIV & "65" & COL_DIV & "0" & COL_DIV & "0"
'            .Col = 5
'            strBody = strBody & vbTab & .Value & COL_DIV & "80" & COL_DIV & "0" & COL_DIV & "0"
'
'            For kk = 6 To .MaxCols
'                .Col = kk
'                If .Col <> .MaxCols Then
'                    jj = jj + 20
'                    strBody = strBody & vbTab & .Value & COL_DIV & jj & COL_DIV & "0" & COL_DIV & "0"
'                Else
'                    jj = jj + 20
'                    strBody = strBody & vbTab & .Value & COL_DIV & jj & COL_DIV & "1" & COL_DIV & "1" & vbTab
'                End If
'            Next
'
'        Next
'    End With
'
'On Error GoTo Errors
'    strBody = Mid(strBody, 1, Len(strBody) - 1)
'    With objReport
'        .Header1 = strHeader1
'        .Header2 = strHeader2
'        .Header3 = strHeader3
'        .Body = strBody
'        Call .CallPrint
'    End With
'    Set objReport = Nothing
'    Me.MousePointer = 0
'    Exit Sub
'
'Errors:
'    MsgBox Err.Description, vbCritical, "오류"
'    Set objReport = Nothing
'    Me.MousePointer = 0
End Sub

Private Sub cmdQuery_Click()
    Dim SSQL    As String
    Dim RS      As Recordset
    Dim strTestcd As String
    Dim sFDate  As String
    Dim sTDate  As String
    Dim strPtid As String
    Dim strAge As String
    Dim objPro  As clsProgress
    Dim ii      As Long
    Dim strAccNo As String
    
    Call medClearTable(tblList)
    tblList.MaxRows = 21
    
    sFDate = Format(dtpFrom.Value, PRESENTDATE_FORMAT)
    sTDate = Format(dtpTo.Value, PRESENTDATE_FORMAT)
    
    If objDic.RecordCount < 1 Then Exit Sub
    
    Set objPro = New clsProgress
    With objPro
        .Container = Me
        .Left = tblList.Left
        .Top = tblList.Top
        .Width = tblList.Width
        .Height = .Height * 2
        .Message = "자료를 읽고 있습니다..."
    End With
    
    objDic.MoveFirst
    
    Do Until objDic.EOF
        strTestcd = strTestcd & "'" & objDic.Fields("testcd") & "',"
        objDic.MoveNext
    Loop
    strTestcd = Mid(strTestcd, 1, Len(strTestcd) - 1)
    Me.MousePointer = 11
    
    SSQL = " select a.vfydt,b.ptid,b.sex,b.ageday,b.wardid,b.deptcd,a.rstcd,a.testcd ,c." & F_PTNM & " as ptnm,d.field1 as result " & _
           " ,a.workarea,a.accdt,a.accseq " & _
           " from " & T_LAB031 & " d," & T_HIS001 & " c," & T_LAB302 & " a," & T_LAB201 & " b " & _
           " where " & DBW("a.vfydt>=", sFDate) & _
           " and   " & DBW("a.vfydt<=", sTDate) & _
           " and a.testcd in (" & strTestcd & ")" & _
           " and " & DBW("cdindex=", "C110") & _
           " and " & DBJ("d.cdval1*=a.testcd") & _
           " and " & DBJ("d.cdval2*=a.rstcd") & _
           " and a.workarea=b.workarea and a.accdt=b.accdt and a.accseq=b.accseq" & _
           " and a.ptid=c." & F_PTID & _
           " order by ptid,workarea,accdt,accseq,testcd" 'ptid,testcd,vfydt"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        objPro.Max = RS.RecordCount
        
        With tblList
            .ReDraw = False
            Do Until RS.EOF
                ii = ii + 1
                objPro.Value = ii
                
'                If strPtid <> RS.Fields("ptid").Value & "" Then
'                    If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
'                    .RowHeight(-1) = 12
'                    .Row = .DataRowCnt + 1
'                    .Col = 1: .Value = Format$(RS.Fields("vfydt").Value & "", "####-##-##")
'                    .Col = 2: .Value = RS.Fields("ptid").Value & ""
'                    .Col = 3: .Value = RS.Fields("ptnm").Value & ""
'                    .Col = 4: .Value = RS.Fields("sex").Value & "" & "/"
'                            If Val(RS.Fields("ageday").Value) < 365 Then
'                                strAge = RS.Fields("ageday").Value & "" & " D"
'                            Else
'                                strAge = CLng(Val(RS.Fields("ageday").Value & "") / 365)
'                            End If
'                            .Value = .Value & strAge
'                    .Col = 5: .Value = RS.Fields("wardid").Value & ""
'                            If .Value <> RS.Fields("deptcd").Value And RS.Fields("deptcd").Value & "" <> "" Then
'                                .Value = IIf(.Value = "", RS.Fields("deptcd").Value & "", .Value & "-" & RS.Fields("deptcd").Value & "")
'                            End If
'                End If
'2005/05/31 modify by legends
'중복된 환자를 표시할 지 여부 기능 추가
'접수번호가 같으면 같은 줄에 표시
                If chkPtidDupcheck.Value = 1 Then
                    If strPtid <> RS.Fields("ptid").Value & "" Then
                        If .DataRowCnt >= .MaxRows Then
                            .MaxRows = .MaxRows + 1
                        End If
                        .RowHeight(-1) = 12
                        
                        .Row = .DataRowCnt + 1
                        .Col = 1: .Value = Format$(RS.Fields("vfydt").Value & "", "####-##-##")
                        .Col = 2: .Value = RS.Fields("ptid").Value & ""
                        .Col = 3: .Value = RS.Fields("ptnm").Value & ""
                        .Col = 4: .Value = RS.Fields("sex").Value & "" & "/"
                                If Val(RS.Fields("ageday").Value) < 365 Then
                                    strAge = RS.Fields("ageday").Value & "" & " D"
                                Else
                                    strAge = CLng(Val(RS.Fields("ageday").Value & "") / 365)
                                End If
                                .Value = .Value & strAge
                        .Col = 5: .Value = RS.Fields("wardid").Value & ""
                                If .Value <> RS.Fields("deptcd").Value And RS.Fields("deptcd").Value & "" <> "" Then
                                    .Value = IIf(.Value = "", RS.Fields("deptcd").Value & "", .Value & "-" & RS.Fields("deptcd").Value & "")
                                End If
                    End If
                Else
                    If strAccNo <> RS.Fields("workarea").Value & "" & RS.Fields("accdt").Value & "" & RS.Fields("accseq").Value & "" Then
                        If .DataRowCnt >= .MaxRows Then
                            .MaxRows = .MaxRows + 1
                        End If
                        .RowHeight(-1) = 12
                        
                        .Row = .DataRowCnt + 1
                        .Col = 1: .Value = Format$(RS.Fields("vfydt").Value & "", "####-##-##")
                        .Col = 2: .Value = RS.Fields("ptid").Value & ""
                        .Col = 3: .Value = RS.Fields("ptnm").Value & ""
                        .Col = 4: .Value = RS.Fields("sex").Value & "" & "/"
                                If Val(RS.Fields("ageday").Value) < 365 Then
                                    strAge = RS.Fields("ageday").Value & "" & " D"
                                Else
                                    strAge = CLng(Val(RS.Fields("ageday").Value & "") / 365)
                                End If
                                .Value = .Value & strAge
                        .Col = 5: .Value = RS.Fields("wardid").Value & ""
                                If .Value <> RS.Fields("deptcd").Value And RS.Fields("deptcd").Value & "" <> "" Then
                                    .Value = IIf(.Value = "", RS.Fields("deptcd").Value & "", .Value & "-" & RS.Fields("deptcd").Value & "")
                                End If
                    End If
                End If
                
                If objDic.Exists(RS.Fields("testcd").Value & "") Then
                    objDic.KeyChange RS.Fields("testcd").Value & ""
                    .Col = objDic.Fields("col"): .Value = IIf(RS.Fields("result").Value & "" = "", RS.Fields("rstcd").Value & "", RS.Fields("result").Value & "")
                End If
                
                strAccNo = RS.Fields("workarea").Value & "" & RS.Fields("accdt").Value & "" & RS.Fields("accseq").Value & ""
                strPtid = RS.Fields("ptid").Value & ""
                
                RS.MoveNext
            Loop
            .Row = 1: .Action = ActionActiveCell
            .ReDraw = True
        End With
        
        Debug.Print tblList.DataRowCnt
    End If
    Me.MousePointer = 0
    Set RS = Nothing
End Sub

Private Sub Form_Load()
    Set objDic = New clsDictionary
    dtpFrom.Value = GetSystemDate
    dtpTo.Value = GetSystemDate
    Call SpreadSet
    Call medClearTable(tblList)
End Sub

Private Sub SpreadSet()
    Dim SSQL    As String
    Dim RS      As Recordset
    Dim ii      As Integer
    
    objDic.Clear
    objDic.FieldInialize "testcd", "testnm,col"
    
    SSQL = " select a.testcd,a.abbrnm5 from " & T_LAB001 & " a," & T_COM003 & " b " & _
           " where " & DBW("b.cdindex=", BC2_REACTION_TEST) & _
           " and a.testcd=b.cdval1"
    SSQL = SSQL & " Union  "
    SSQL = SSQL & " select b.field1 as testcd,a.abbrnm5 from " & T_LAB001 & " a," & T_COM003 & " b " & _
                  " where " & DBW("b.cdindex=", BC2_ABO_TEST) & _
                  " and b.cdval1=(select max(cdval1) from " & T_COM003 & _
                                  " where " & DBW("cdindex=", BC2_ABO_TEST) & _
                                 ")" & _
                  " and a.testcd=b.field1"
    SSQL = SSQL & " Union  "
    SSQL = SSQL & " select a.testcd,a.abbrnm5 from " & T_LAB001 & " a," & T_COM003 & " b " & _
                  " where " & DBW("b.cdindex=", BC2_ABO_TEST) & _
                  " and b.cdval1=(select max(cdval1) from " & T_COM003 & _
                                  " where " & DBW("cdindex=", BC2_ABO_TEST) & _
                                 ")" & _
                  " and a.testcd=b.field3"
    SSQL = SSQL & " order by testcd"
    
    ii = 5
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
'            Debug.Print RS.Fields("testcd").Value & ""
            If Not objDic.Exists(RS.Fields("testcd").Value & "") Then
                ii = ii + 1
                objDic.AddNew RS.Fields("testcd").Value & "", RS.Fields("abbrnm5").Value & "" & COL_DIV & ii
            End If
            RS.MoveNext
        Loop
        objDic.MoveFirst
    End If
    Set RS = Nothing
    If objDic.RecordCount > 0 Then
        With tblList
            .Row = 0
            .MaxCols = 5 + objDic.RecordCount
            Do Until objDic.EOF
                .Col = objDic.Fields("col"): .Value = objDic.Fields("testnm")
                objDic.MoveNext
            Loop
            For ii = 6 To .MaxCols
                .ColWidth(ii) = .ColWidth(6)
            Next
        End With
    End If
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then Call SpreadSort(Col)
End Sub

Private Sub SpreadSort(ByVal Col As Integer)
    With tblList
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = Col
        
        If SortTF = True Then
            .SortKeyOrder(1) = SortKeyOrderAscending
            SortTF = False
        Else
            SortTF = True
            .SortKeyOrder(1) = SortKeyOrderDescending
        End If
        
        .Col = 1:  .Col2 = .MaxCols
        .Row = 1:  .Row2 = .DataRowCnt
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub
