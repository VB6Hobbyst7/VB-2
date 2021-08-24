VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBBS961 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS961.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread tblBlood 
      Height          =   6900
      Left            =   135
      TabIndex        =   13
      Top             =   1455
      Width           =   10710
      _Version        =   196608
      _ExtentX        =   18891
      _ExtentY        =   12171
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   24
      OperationMode   =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "frmBBS961.frx":076A
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8085
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "15101"
      Top             =   8430
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6765
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "124"
      Top             =   8430
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9405
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   8430
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   225
      TabIndex        =   6
      Top             =   8325
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
      SpreadDesigner  =   "frmBBS961.frx":0E75
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   1005
      Top             =   8445
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   90
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   10770
      _ExtentX        =   18997
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
      Caption         =   "혈액출고/폐기대장"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   105
      TabIndex        =   3
      Top             =   285
      Width           =   10785
      Begin VB.ComboBox cboReturn 
         Height          =   300
         Left            =   4275
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   630
         Width           =   1230
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   2
         Tag             =   "124"
         Top             =   465
         Width           =   1320
      End
      Begin VB.ComboBox cboCenter 
         Height          =   300
         ItemData        =   "frmBBS961.frx":1020
         Left            =   1290
         List            =   "frmBBS961.frx":1022
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   225
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker dtpFMonth 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   0
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   83623939
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpTMonth 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   2910
         TabIndex        =   1
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   83623939
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   135
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   225
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
         Caption         =   "Center"
         Appearance      =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2685
         TabIndex        =   5
         Top             =   705
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmBBS961"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblCol
    tcDeliverydt = 1
    tcPtid
    tcBussdiv
    tcName
    tcSSN
    tcABO
    tcBldNo
    tcAbbrnm
    tcVOL
    tcUNIT
    tcRetcnt
    tcExpcnt
End Enum
'    tcReqcnt

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblBlood.DataRowCnt = 0 And tblBlood.DataRowCnt = 0 Then Exit Sub
    
    With tblBlood
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblBlood.MaxRows + 1
        .MaxCols = tblBlood.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblBlood.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "혈액 출고(폐기)대장"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim strTmp      As String
    Dim iTotUse     As Long
    Dim iTotRet     As Long
    Dim iTotExp     As Long
    Dim i           As Integer
    
    With tblBlood
        For i = 1 To .DataRowCnt
            .Row = i
            
            .Col = tblCol.tcUNIT: iTotUse = iTotUse + IIf(.Value = "", 0, Val(.Value))
            .Col = tblCol.tcRetcnt: iTotRet = iTotRet + IIf(.Value = "", 0, Val(.Value))
            .Col = tblCol.tcExpcnt: iTotExp = iTotExp + IIf(.Value = "", 0, Val(.Value))
        Next
        
        strTmp = " ♧ 합 계 : " & "수혈량" & "(" & iTotUse & ")" & "," & "반환량" & "(" & iTotRet & ")" & "," & "폐기량" & "(" & iTotExp & ")"
        
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .PrintJobName = "혈액출고(폐기)대장 출력"
        .PrintAbortMsg = "혈액출고(폐기)대장 출력중 입니다. "
        
        .PrintColor = False
        .PrintFirstPageNumber = 1
        
        .PrintHeader = "/n/n/l/fb1 " & "♧ 혈액출고(폐기)대장 사용량 (" & Format(dtpFMonth.Value, CS_DateLongFormat) & " 부터 " & _
                                                              Format(dtpTMonth.Value, CS_DateLongFormat) & " 까지 ) /c/fb1/n" & _
                                       " ♧ 센 터 : " & medGetP(cboCenter.Text, 1, COL_DIV) & "/n" & _
                                       strTmp & "/n/n"
                                       
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
End Sub

Private Sub cmdQuery_Click()
    Call Query
End Sub

Private Sub Form_Load()
    Call SetCenterCombo
    dtpFMonth.Value = GetSystemDate
    dtpTMonth.Value = GetSystemDate
    
    cboReturn.AddItem ""
    cboReturn.AddItem "반환"
    cboReturn.AddItem "폐기"
    cboReturn.AddItem "방사선"
    cboReturn.ListIndex = 0
    
    ClearAll
End Sub

Private Sub ClearAll()
    medClearTable tblBlood
End Sub

Private Sub SetCenterCombo()
    Dim objcom003 As clsCom003
    Dim i As Long
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter, True)
    Set objcom003 = Nothing
    
    cboCenter.ListIndex = -1
    
    For i = 0 To cboCenter.ListCount - 1
        If ObjSysInfo.BuildingCd = medGetP(cboCenter.List(i), 1, " ") Then
            cboCenter.ListIndex = i
            Exit For
        End If
    Next i
End Sub
Private Sub Query()
    Dim SSQL        As String
    Dim sFDate      As String
    Dim sTDate      As String
    Dim sCentercd   As String
    Dim RS          As Recordset
    Dim objPro As clsProgress
    
    Call ClearAll
    
    sCentercd = medGetP(cboCenter.Text, 1, " ")
    
    sFDate = Format(dtpFMonth.Value, "YYYYMMDD")
    sTDate = Format(dtpTMonth.Value, "YYYYMMDD")
    
'2005/05/30 modify by legends
'속도 저하로 인한 union 제거 후 decode 사용
    '-- 원본 ------------------------------------------------------------------------------------------------------------------
'    SSQL = " SELECT b.deliverydt,f." & F_PTID & " as ptid,e.bussdiv,f." & F_PTNM & " as ptnm," & F_SSN2("f") & " as ssn," & _
'           " a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
'           ",c.abbrnm,a.volumn,b.retfg,decode(a.stscd,'3','1',' ') as unit ,' ' as reqcnt, decode(a.stscd,'1','1',' ') retcnt, decode(a.stscd,'4','1',' ') expcnt " & _
'           " FROM " & T_HIS001 & " f," & T_LAB101 & " e," & T_LAB102 & " d," & _
'                      T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b" & _
'           " WHERE " & _
'                     "a.stscd in ('" & BBSBloodStatus.stsRETURN & "','" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "')" & _
'           " AND " & DBW("b.deliverydt>=", sFDate) & _
'           " AND " & DBW("b.deliverydt<=", sTDate)
'    If sCentercd <> "(ALL)" Then
'        SSQL = SSQL & " AND " & DBW("a.centercd=", sCentercd)
'    End If
'    SSQL = SSQL & _
'           " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd" & _
'           " AND a.compocd=c.compocd" & _
'           " AND b.workarea=d.workarea" & _
'           " AND b.accdt=d.accdt" & _
'           " AND b.accseq=d.accseq" & _
'           " AND d.ptid=e.ptid" & _
'           " AND d.orddt=e.orddt" & _
'           " AND d.ordno=e.ordno" & _
'           " AND d.ptid=f." & F_PTID & _
'           " order by b.deliverydt "
    '---------------------------------------------------------------------------------------------------------------------------
    
    '---------------------------------------------------------------------------------------------------------------------------
    ' 조회 조건 추가 반환/폐기
    ' 온승호
    ' 2012-05-31
    '---------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------
'    SSQL = " select deliverydt,ptid,decode(patsect,'I','2','1') as bussdiv, meddept,ptnm,ssn, abo, rh, bloodno, abbrnm,volumn, retfg, " & _
'             "      decode(reqcnt,' ',decode(expcnt,' ',1,''),'') unit ,decode(reqcnt,'1',1,'') as reqcnt, decode(expcnt,'1',1,'') as expcnt " & _
'             "      from (SELECT b.deliverydt,f.patno as ptid,e.patsect,e.meddept, f.patname as ptnm,f.resno1 || f.resno2 as ssn, a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
'             "        ,c.abbrnm,a.volumn,b.retfg, decode(a.stscd,'3','1',' ') as unit, decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
'             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
'             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
'             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
'             "    AND " & DBW("a.centercd=", "10") & " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
'             "    AND a.compocd=c.compocd " & _
'             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
'             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
'             "    AND d.ptid=f.patno AND e.patsect !='E' "
'
'    SSQL = SSQL & " Union All "
'
'    SSQL = SSQL & " select distinct a.deliverydt, a.ptid, decode(b.patsect,'','2','1') as bussdiv, decode(b.patsect,'',b.meddept,a.meddept) as meddept " & _
'             "        , a.ptnm, a.ssn, a.abo, a.rh, a.bloodno, a.abbrnm,a.volumn, retfg, a.unit unit,a.reqcnt reqcnt,a.expcnt expcnt " & _
'             "        from (SELECT b.deliverydt,f.patno as ptid, e.patsect,e.meddept, f.patname as ptnm, f.resno1 || f.resno2 as ssn, a.abo,a.rh, a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
'             "        ,c.abbrnm,a.volumn, b.retfg, decode(a.stscd,'3','1',' ') as unit,decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
'             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
'             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
'             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
'             "    AND " & DBW("a.centercd=", "10") & _
'             "    AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
'             "    AND a.compocd=c.compocd " & _
'             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
'             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
'             "    AND d.ptid=f.patno AND e.patsect='E' "
'
'    SSQL = SSQL & _
'             " ) a," & _
'             " (SELECT a.ptid,b.deliverydt, c.patsect,c.meddept " & _
'             "    FROM S2BBS202 A, S2bbs402 B, ORAm1.mdbldort C, oraa1.apipdlst D " & _
'             "   Where a.accdt = b.accdt AND   A.ACCSEQ = B.ACCSEQ AND A.PTID = C.PATNO and a.ptid = d.patno " & _
'             "     and a.orddt >= d.admdate and a.orddt <= d.dschdate AND A.ACCDT = C.ACCDT AND A.ACCSEQ = C.ACCSEQ " & _
'             "     AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
'             " ) b " & _
'             "   where a.deliverydt=b.deliverydt(+) and  a.ptid=b.ptid(+) ) "
'
'    SSQL = SSQL & _
'           " order by deliverydt, ptid "
    '---------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------
    
    'SSQL = " select deliverydt,ptid,decode(patsect,'I','2','1') as bussdiv, meddept,ptnm,ssn, abo, rh, bloodno, abbrnm,volumn, retfg, " & _
             "      decode(reqcnt,' ',decode(expcnt,' ',1,''),'') unit ,decode(reqcnt,'1',1,'') as reqcnt, decode(expcnt,'1',1,'') as expcnt " & _
             "      from (SELECT b.deliverydt,f.patno as ptid,e.patsect,e.meddept, f.patname as ptnm,f.resno1 || f.resno2 as ssn, a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn,b.retfg, decode(a.stscd,'3','1',' ') as unit, decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
             "    AND " & DBW("a.centercd=", "10") & " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect !='E' "
    
    ''20131214 DB 암호화 CryptIT.decrypt(resno1,'pmc1898')
    
'    SSQL = " select deliverydt,ptid,decode(patsect,'I','2','1') as bussdiv, meddept,ptnm,ssn, abo, rh, bloodno, abbrnm,volumn, retfg, " & _
'             "      decode(reqcnt,' ',decode(expcnt,' ',1,''),'') unit ,decode(reqcnt,'1',1,'') as reqcnt, decode(expcnt,'1',1,'') as expcnt " & _
'             "      from (SELECT b.deliverydt,f.patno as ptid,e.patsect,e.meddept, f.patname as ptnm,CryptIT.decrypt(f.resno1,'pmc1898') || CryptIT.decrypt(f.resno2,'pmc1898') as ssn, a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
'             "        ,c.abbrnm,a.volumn,b.retfg, decode(a.stscd,'3','1',' ') as unit, decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
'             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
'             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
'             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
'             "    AND " & DBW("a.centercd=", "10") & " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
'             "    AND a.compocd=c.compocd " & _
'             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
'             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
'             "    AND d.ptid=f.patno AND e.patsect !='E' "
    
    ''20170814 DB 암호화 crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno2) AS ssn,
    
    SSQL = " select deliverydt,ptid,decode(patsect,'I','2','1') as bussdiv, meddept,ptnm,ssn, abo, rh, bloodno, abbrnm,volumn, retfg, " & _
             "      decode(reqcnt,' ',decode(expcnt,' ',1,''),'') unit ,decode(reqcnt,'1',1,'') as reqcnt, decode(expcnt,'1',1,'') as expcnt " & _
             "      from (SELECT b.deliverydt,f.patno as ptid,e.patsect,e.meddept, f.patname as ptnm,crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno2) AS ssn, a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn,b.retfg, decode(a.stscd,'3','1',' ') as unit, decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
             "    AND " & DBW("a.centercd=", "10") & " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect !='E' "
             
    Select Case cboReturn.ListIndex
        Case 0

        Case 1
            SSQL = SSQL & " AND b.retfg = '1'"
        Case 2
            SSQL = SSQL & " AND b.expfg = '1'"
        Case 3
            SSQL = SSQL & " AND a.irrfg = '1'"
    End Select
    
    SSQL = SSQL & " Union All "
    
    ''20131214 DB 암호화 CryptIT.decrypt(resno1,'pmc1898')
    
    'SSQL = SSQL & " select distinct a.deliverydt, a.ptid, decode(b.patsect,'','2','1') as bussdiv, decode(b.patsect,'',b.meddept,a.meddept) as meddept " & _
             "        , a.ptnm, a.ssn, a.abo, a.rh, a.bloodno, a.abbrnm,a.volumn, retfg, a.unit unit,a.reqcnt reqcnt,a.expcnt expcnt " & _
             "        from (SELECT b.deliverydt,f.patno as ptid, e.patsect,e.meddept, f.patname as ptnm, f.resno1 || f.resno2 as ssn, a.abo,a.rh, a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn, b.retfg, decode(a.stscd,'3','1',' ') as unit,decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
             "    AND " & DBW("a.centercd=", "10") & _
             "    AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect='E' "
             
'    SSQL = SSQL & " select distinct a.deliverydt, a.ptid, decode(b.patsect,'','2','1') as bussdiv, decode(b.patsect,'',b.meddept,a.meddept) as meddept " & _
'             "        , a.ptnm, a.ssn, a.abo, a.rh, a.bloodno, a.abbrnm,a.volumn, retfg, a.unit unit,a.reqcnt reqcnt,a.expcnt expcnt " & _
'             "        from (SELECT b.deliverydt,f.patno as ptid, e.patsect,e.meddept, f.patname as ptnm, CryptIT.decrypt(f.resno1,'pmc1898') || CryptIT.decrypt(f.resno2,'pmc1898') as ssn, a.abo,a.rh, a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
'             "        ,c.abbrnm,a.volumn, b.retfg, decode(a.stscd,'3','1',' ') as unit,decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
'             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
'             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
'             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
'             "    AND " & DBW("a.centercd=", "10") & _
'             "    AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
'             "    AND a.compocd=c.compocd " & _
'             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
'             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
'             "    AND d.ptid=f.patno AND e.patsect='E' "
             
    SSQL = SSQL & " select distinct a.deliverydt, a.ptid, decode(b.patsect,'','2','1') as bussdiv, decode(b.patsect,'',b.meddept,a.meddept) as meddept " & _
             "        , a.ptnm, a.ssn, a.abo, a.rh, a.bloodno, a.abbrnm,a.volumn, retfg, a.unit unit,a.reqcnt reqcnt,a.expcnt expcnt " & _
             "        from (SELECT b.deliverydt,f.patno as ptid, e.patsect,e.meddept, f.patname as ptnm, crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno2) AS ssn, a.abo,a.rh, a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn, b.retfg, decode(a.stscd,'3','1',' ') as unit,decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('" & BBSBloodStatus.stsDELIVERY & "','" & BBSBloodStatus.stsEXPIRE & "') " & _
             "    AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
             "    AND " & DBW("a.centercd=", "10") & _
             "    AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect='E' "
             
    Select Case cboReturn.ListIndex
        Case 0

        Case 1
            SSQL = SSQL & " AND b.retfg = '1'"
        Case 2
            SSQL = SSQL & " AND b.expfg = '1'"
        Case 3
            SSQL = SSQL & " AND a.irrfg = '1'"
    End Select
                 

    SSQL = SSQL & _
             " ) a," & _
             " (SELECT a.ptid,b.deliverydt, c.patsect,c.meddept " & _
             "    FROM S2BBS202 A, S2bbs402 B, ORAm1.mdbldort C, oraa1.apipdlst D " & _
             "   Where a.accdt = b.accdt AND   A.ACCSEQ = B.ACCSEQ AND A.PTID = C.PATNO and a.ptid = d.patno " & _
             "     and a.orddt >= d.admdate and a.orddt <= d.dschdate AND A.ACCDT = C.ACCDT AND A.ACCSEQ = C.ACCSEQ " & _
             "     AND " & DBW("b.deliverydt>=", sFDate) & " AND " & DBW("b.deliverydt<=", sTDate) & _
             " ) b " & _
             "   where a.deliverydt=b.deliverydt(+)  and  a.ptid=b.ptid(+) ) "
             
    SSQL = SSQL & _
           " order by deliverydt, ptid "
    
    
    
    
'''''''''    SSQL = " SELECT b.deliverydt,f." & F_PTID & " as ptid,e.bussdiv,f." & F_PTNM & " as ptnm," & F_SSN2("f") & " as ssn," & _
'''''''''           " a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||a.bldno as bloodno " & _
'''''''''           ",c.abbrnm,a.volumn,'1' as unit ,' ' as reqcnt,' 'expcnt" & _
'''''''''           " FROM " & T_HIS001 & " f," & T_LAB101 & " e," & T_LAB102 & " d," & _
'''''''''                      T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b" & _
'''''''''           " WHERE " & _
'''''''''                     DBW("a.stscd=", BBSBloodStatus.stsDELIVERY) & _
'''''''''           " AND " & DBW("b.deliverydt>=", sFDate) & _
'''''''''           " AND " & DBW("b.deliverydt<=", sTDate)
'''''''''    If sCentercd <> "(ALL)" Then
'''''''''        SSQL = SSQL & " AND " & DBW("a.centercd=", sCentercd)
'''''''''    End If
'''''''''    SSQL = SSQL & _
'''''''''           " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd" & _
'''''''''           " AND a.compocd=c.compocd" & _
'''''''''           " AND b.workarea=d.workarea" & _
'''''''''           " AND b.accdt=d.accdt" & _
'''''''''           " AND b.accseq=d.accseq" & _
'''''''''           " AND d.ptid=e.ptid" & _
'''''''''           " AND d.orddt=e.orddt" & _
'''''''''           " AND d.ordno=e.ordno" & _
'''''''''           " AND d.ptid=f." & F_PTID
'''''''''
'''''''''    SSQL = SSQL & " UNION " & _
'''''''''           " SELECT b.deliverydt,f." & F_PTID & " as ptid,e.bussdiv,f." & F_PTNM & " as ptnm," & F_SSN2("f") & " as ssn," & _
'''''''''           " a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||a.bldno as bloodno ," & _
'''''''''           " c.abbrnm,a.volumn,'1' as unit  ,' ' as reqcnt,'1' expcnt" & _
'''''''''           " FROM " & T_HIS001 & " f," & T_LAB101 & " e," & T_LAB102 & " d," & _
'''''''''                      T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b" & _
'''''''''           " WHERE " & _
'''''''''                     DBW("a.stscd=", BBSBloodStatus.stsEXPIRE) & _
'''''''''           " AND " & DBW("b.deliverydt>=", sFDate) & _
'''''''''           " AND " & DBW("b.deliverydt<=", sTDate)
'''''''''    If sCentercd <> "(ALL)" Then
'''''''''        SSQL = SSQL & " AND " & DBW("a.centercd=", sCentercd)
'''''''''    End If
'''''''''    SSQL = SSQL & _
'''''''''           " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd" & _
'''''''''           " AND a.compocd=c.compocd" & _
'''''''''           " AND b.workarea=d.workarea" & _
'''''''''           " AND b.accdt=d.accdt" & _
'''''''''           " AND b.accseq=d.accseq" & _
'''''''''           " AND d.ptid=e.ptid" & _
'''''''''           " AND d.orddt=e.orddt" & _
'''''''''           " AND d.ordno=e.ordno" & _
'''''''''           " AND d.ptid=f." & F_PTID & _
'''''''''           " ORDER BY deliverydt"
                   
    Screen.MousePointer = vbHourglass
                   
    Set objPro = New clsProgress
    With objPro
        .Container = Me
        .Left = tblBlood.Left
        .Top = tblBlood.Top
        .Width = tblBlood.Width
        .Height = .Height * 2
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn, , , adCmdText
    
    If Not RS.EOF Then
        objPro.Message = "자료를 읽고 있습니다..."
        objPro.Max = RS.RecordCount
        With tblBlood
            Do Until RS.EOF
                objPro.Value = objPro.Value + 1
                
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = tblCol.tcDeliverydt: .Value = Format(RS.Fields("deliverydt").Value & "", "####-##-##")
                .Col = tblCol.tcPtid: .Value = RS.Fields("ptid").Value & ""
                .Col = tblCol.tcBussdiv
                    Select Case RS.Fields("bussdiv").Value & ""
                        Case "1": .Value = "외래"
                        Case "2": .Value = "입원"
                        Case Else: .Value = ""
                    End Select
                .Col = tblCol.tcName: .Value = RS.Fields("ptnm").Value & ""
                .Col = tblCol.tcSSN
                If Len(RS.Fields("ssn").Value & "") > 6 Then
                    .Value = Mid(RS.Fields("ssn").Value & "", 1, 6) & "-" & Mid(RS.Fields("ssn").Value & "", 7)
                Else
                    .Value = RS.Fields("ssn").Value & ""
                End If
                .Col = tblCol.tcABO:    .Value = RS.Fields("abo").Value & "" & RS.Fields("rh").Value & ""
                .Col = tblCol.tcBldNo:  .Value = RS.Fields("bloodno").Value & ""
                .Col = tblCol.tcAbbrnm: .Value = RS.Fields("abbrnm").Value & ""
                .Col = tblCol.tcVOL:    .Value = RS.Fields("volumn").Value & "" & "cc"
                If RS.Fields("retfg").Value & "" = "1" Then
                    .Col = tblCol.tcRetcnt: .Value = "1"
                Else
                    .Col = tblCol.tcUNIT:   .Value = RS.Fields("unit").Value & ""
                    .Col = tblCol.tcExpcnt: .Value = RS.Fields("expcnt").Value & ""
                End If
                
                If RS.Fields("expcnt").Value & "" = "1" Then '폐기된 혈액
                    .Row = .Row: .Row2 = .Row
                    .Col = 1: .Col2 = .MaxCols
                    .BlockMode = True
                    .ForeColor = vbBlue
                    .BlockMode = False
                Else
                    .Row = .Row: .Row2 = .Row
                    .Col = 1: .Col2 = .MaxCols
                    .BlockMode = True
                    .ForeColor = vbBlack
                    .BlockMode = False
                End If
                
                RS.MoveNext
            Loop
        End With
    End If
    
    Set RS = Nothing
    Set objPro = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub SampleTest()
    Dim strSQL      As String
    
    '20131214 DB 암호화 CryptIT.decrypt(resno1,'pmc1898')
    
    'strSQL = " select deliverydt,ptid,decode(patsect,'I','입원','외래'),meddept,ptnm,ssn, abo, rh, bloodno, abbrnm,volumn, decode(reqcnt,' ',decode(expcnt,' ',1,''),'') unit ,decode(reqcnt,'1',1,''), decode(expcnt,'1',1,'') " & _
             "        from (SELECT b.deliverydt,f.patno as ptid,e.patsect,e.meddept, f.patname as ptnm,f.resno1 || f.resno2 as ssn, a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn,decode(a.stscd,'3','1',' ') as unit, decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('3','4') AND  b.deliverydt >= '20070701' AND  b.deliverydt <= '20070930' " & _
             "    AND a.centercd = '10' AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect !='E' "
             
    strSQL = " select deliverydt,ptid,decode(patsect,'I','입원','외래'),meddept,ptnm,ssn, abo, rh, bloodno, abbrnm,volumn, decode(reqcnt,' ',decode(expcnt,' ',1,''),'') unit ,decode(reqcnt,'1',1,''), decode(expcnt,'1',1,'') " & _
             "        from (SELECT b.deliverydt,f.patno as ptid,e.patsect,e.meddept, f.patname as ptnm, CryptIT.decrypt(f.resno1,'pmc1898') || CryptIT.decrypt(f.resno2,'pmc1898') as ssn, a.abo,a.rh,a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn,decode(a.stscd,'3','1',' ') as unit, decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('3','4') AND  b.deliverydt >= '20070701' AND  b.deliverydt <= '20070930' " & _
             "    AND a.centercd = '10' AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect !='E' "
    
    strSQL = strSQL & " Union All "
    
    'strSQL = " select distinct a.deliverydt, a.ptid, decode(b.patsect,'',a.patsect,'I'),decode(b.patsect,'',a.meddept,b.meddept) " & _
             "        , a.ptnm, a.ssn, a.abo, a.rh, a.bloodno, a.abbrnm,a.volumn, a.unit unit,a.reqcnt reqcnt,a.expcnt expcnt " & _
             "        from (SELECT b.deliverydt,f.patno as ptid, e.patsect,e.meddept, f.patname as ptnm, f.resno1 || f.resno2 as ssn, a.abo,a.rh, a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn,decode(a.stscd,'3','1',' ') as unit,decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('3','4') AND  b.deliverydt >= '20070701' AND  b.deliverydt <= '20070930' AND  a.centercd = '10' " & _
             "    AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect='E' "
             
    '20131214 DB 암호화 CryptIT.decrypt(resno1,'pmc1898')
    strSQL = " select distinct a.deliverydt, a.ptid, decode(b.patsect,'',a.patsect,'I'),decode(b.patsect,'',a.meddept,b.meddept) " & _
             "        , a.ptnm, a.ssn, a.abo, a.rh, a.bloodno, a.abbrnm,a.volumn, a.unit unit,a.reqcnt reqcnt,a.expcnt expcnt " & _
             "        from (SELECT b.deliverydt,f.patno as ptid, e.patsect,e.meddept, f.patname as ptnm, CryptIT.decrypt(f.resno1,'pmc1898') || CryptIT.decrypt(f.resno2,'pmc1898') as ssn, a.abo,a.rh, a.bldsrc||'-'||a.bldyy||'-'||trim(to_char(a.bldno,'000000')) as bloodno " & _
             "        ,c.abbrnm,a.volumn,decode(a.stscd,'3','1',' ') as unit,decode(b.retfg,'1','1',' ') as reqcnt,decode(b.expfg,'1','1',' ') expcnt " & _
             "   FROM ORAA1.APPATBAT f,ORAm1.mdbldort e,ORAS1.S2ORD102_V d,s2bbs006 c,s2bbs401 a,s2bbs402 b " & _
             "  WHERE a.stscd in ('3','4') AND  b.deliverydt >= '20070701' AND  b.deliverydt <= '20070930' AND  a.centercd = '10' " & _
             "    AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
             "    AND a.compocd=c.compocd " & _
             "    AND b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq " & _
             "    AND d.ptid=e.patno AND d.orddt=e.orddate AND d.ordno=e.ordseqno " & _
             "    AND d.ptid=f.patno AND e.patsect='E' "
             
    strSQL = strSQL & _
             " ) a," & _
             " (SELECT a.ptid,b.deliverydt, c.patsect,c.meddept " & _
             "    FROM S2BBS202 A, S2bbs402 B, ORAm1.mdbldort C, oraa1.apipdlst D " & _
             "   Where a.accdt = b.accdt AND   A.ACCSEQ = B.ACCSEQ AND A.PTID = C.PATNO and a.ptid = d.patno " & _
             "     and a.orddt >= d.admdate and a.orddt <= d.dschdate AND A.ACCDT = C.ACCDT AND A.ACCSEQ = C.ACCSEQ " & _
             "     AND A.ACCDT = '2007' AND B.DELIVERYDT >= '20070701' AND B.DELIVERYDT <= '20070930' " & _
             " ) b " & _
             "   where a.deliverydt=b.deliverydt(+) and  a.ptid=b.ptid(+) ) "
             
End Sub
