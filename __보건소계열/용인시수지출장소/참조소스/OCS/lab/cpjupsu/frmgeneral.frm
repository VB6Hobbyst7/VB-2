VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmGeneral 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "접수환자 접수Data 확인 Form"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   960
   ClientWidth     =   11685
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2130
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3360
      _ExtentY        =   3757
      _StockProps     =   15
      Caption         =   "접수일자조건"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36356
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36356
      End
   End
   Begin FPSpreadADO.fpSpread ssGeneral 
      Height          =   3615
      Left            =   2745
      TabIndex        =   0
      Top             =   45
      Width           =   8790
      _Version        =   196608
      _ExtentX        =   15505
      _ExtentY        =   6376
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
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
      GrayAreaBackColor=   8421504
      MaxCols         =   22
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmGeneral.frx":0000
      Appearance      =   1
      TextTip         =   1
      ScrollBarTrack  =   1
   End
   Begin MSForms.CommandButton cmdLabel 
      Height          =   600
      Left            =   90
      TabIndex        =   1
      Top             =   2250
      Width           =   1905
      Caption         =   "Call BarCode"
      PicturePosition =   327683
      Size            =   "3360;1058"
      Picture         =   "frmGeneral.frx":4144
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdLabel_Click()
    'frmLabel.Show vbModal
    
    ssGeneral.Row = ssGeneral.ActiveRow
    ssGeneral.Col = 4: GLabelPtno = ssGeneral.Text
    ssGeneral.Col = 3: GLabelJeobsuDt = ssGeneral.Text
    
    frmBarCode.Show vbModal
    

End Sub

Private Sub Form_Load()
    Dim sFrDate     As String
    Dim sToDate     As String
    
    '/ - ssGeneral Column -------------------------------------------
    '/  1. Button          11. JeobsuT1 : JeobsuT2         21. gbEr
    '/  2. RowID           12. Indate
    '/  3. JeobsuDt        13. RoomCode
    '/  4. Ptno            14. DeptCode
    '/  5. Sname           15. Gbio
    '/  6. Sex             16. DrCode
    '/  7. AgeYY           17. Bi
    '/  8. AgeMM           18. OrderDt
    '/  9. SLipno1         19. Orderno
    '/ 10. SLipno2         20. Bi
    '/---------------------------------------------------------------
    
    dtFrDate.Value = Format(frmMain.dtFromDate.Value, "yyyy-MM-dd")
    dtToDate.Value = Format(frmMain.dtToDate.Value, "yyyy-MM-dd")
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    GoSub Get_General_Data
    Exit Sub
    
Get_General_Data:
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID RwID,"
    StrSql = StrSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate,"
    StrSql = StrSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,"
    StrSql = StrSql & "        b.Sname, c.Codenm SLname"
    StrSql = StrSql & " FROM   TWEXAM_General a,"
    StrSql = StrSql & "        TWBAS_Patient  b,"
    StrSql = StrSql & "        TWEXAM_Specode c "
    StrSql = StrSql & " WHERE  a.JeobsuDt  >=  TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.JeobsuDt  <=  TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.Ptno       =  b.Ptno(+)"
    StrSql = StrSql & " AND    c.Codegu     = '12'"
    StrSql = StrSql & " AND    TO_NUMBER(c.Codeky)     =  a.Slipno1"
    StrSql = StrSql & " ORDER  BY a.JeobsuDt, a.Ptno"
    
    ssGeneral.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Return
    ssGeneral.MaxRows = adoSet.RecordCount
    ssGeneral.RowHeight(-1) = 11
    
    Do Until adoSet.EOF
        ssGeneral.Row = ssGeneral.DataRowCnt + 1
        ssGeneral.Col = 2:  ssGeneral.Text = adoSet.Fields("RwID").Value & ""
        ssGeneral.Col = 3:  ssGeneral.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssGeneral.Col = 4:  ssGeneral.Text = adoSet.Fields("Ptno").Value & ""
        ssGeneral.Col = 5:  ssGeneral.Text = adoSet.Fields("Sname").Value & ""
        ssGeneral.Col = 6:  ssGeneral.Text = adoSet.Fields("Sex").Value & ""
        ssGeneral.Col = 7:  ssGeneral.Text = adoSet.Fields("AgeYY").Value & ""
        ssGeneral.Col = 8:  ssGeneral.Text = adoSet.Fields("AgeMM").Value & ""
        ssGeneral.Col = 9:  ssGeneral.Text = adoSet.Fields("SLipno1").Value & ""
        ssGeneral.Col = 10: ssGeneral.Text = adoSet.Fields("SLname").Value & ""
        ssGeneral.Col = 11: ssGeneral.Text = adoSet.Fields("SLipno2").Value & ""
        ssGeneral.Col = 12: ssGeneral.Text = Format(adoSet.Fields("JeobsuT1").Value, "00") & ":" & _
                                             Format(adoSet.Fields("JeobsuT2").Value, "00")
        ssGeneral.Col = 13: ssGeneral.Text = adoSet.Fields("Indate").Value & ""
        ssGeneral.Col = 14: ssGeneral.Text = adoSet.Fields("RoomCode").Value & ""
        ssGeneral.Col = 15: ssGeneral.Text = adoSet.Fields("DeptCode").Value & ""
        ssGeneral.Col = 16: ssGeneral.Text = adoSet.Fields("GBio").Value & ""
        ssGeneral.Col = 17: ssGeneral.Text = adoSet.Fields("Drcode").Value & ""
        ssGeneral.Col = 18: ssGeneral.Text = adoSet.Fields("Bi").Value & ""
        ssGeneral.Col = 19: ssGeneral.Text = adoSet.Fields("OrderDt").Value & ""
        ssGeneral.Col = 20: ssGeneral.Text = adoSet.Fields("Orderno").Value & ""
        ssGeneral.Col = 21: ssGeneral.Text = adoSet.Fields("Bi").Value & ""
        ssGeneral.Col = 22: ssGeneral.Text = adoSet.Fields("gbEr").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    GoSub frmMain_TextBox_Clear
    GoSub Spread_ssEnrol_Clear
    Exit Sub

frmMain_TextBox_Clear:
    frmMain.txtPtno.Text = ""
    frmMain.txtSname.Text = ""
    frmMain.txtSex.Text = ""
    frmMain.txtBirthDate.Text = ""
    frmMain.txtJumin1.Text = ""
    frmMain.txtJumin2.Text = ""
    frmMain.txtAgeYY.Text = ""
    
    Return
    
    
Spread_ssEnrol_Clear:
    frmMain.ssEnrol.ReDraw = False
    frmMain.ssEnrol.MaxRows = 0
    frmMain.ssEnrol.MaxRows = 100
    frmMain.ssEnrol.RowHeight(-1) = 11
    frmMain.ssEnrol.ReDraw = True
    Return

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssGeneral_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sJeobsuDt       As String
    Dim nSLipno1        As Integer
    Dim nSLipno2        As Integer
    
    ssGeneral.Row = Row
    ssGeneral.Col = 3:  sJeobsuDt = ssGeneral.Text
    ssGeneral.Col = 9:  nSLipno1 = Val(ssGeneral.Text)
    ssGeneral.Col = 11: nSLipno2 = Val(ssGeneral.Text)
    
    GoSub Hand_Flag_Set
    GoSub Spread_ssEnrol_Clear
    GoSub Get_General_Sub_Data
    
    
    ssGeneral.Row = Row
    ssGeneral.Col = 4: frmMain.txtPtno.Text = ssGeneral.Text
    ssGeneral.Col = 5: frmMain.txtSname.Text = ssGeneral.Text
    ssGeneral.Col = 6: frmMain.txtSex.Text = ssGeneral.Text
    ssGeneral.Col = 7: frmMain.txtAgeYY.Text = ssGeneral.Text
    
    StrSql = ""
    StrSql = StrSql & " SELECT Jumin1, Jumin2, TO_CHAR(BirthDate, 'YYYY-MM-DD') BirthDate"
    StrSql = StrSql & " From   TWBAS_Patient "
    StrSql = StrSql & " WHERE  Ptno = '" & frmMain.txtPtno.Text & "'"
    If adoSetOpen(StrSql, adoSet) Then
        frmMain.txtJumin1.Text = adoSet.Fields("Jumin1").Value & ""
        frmMain.txtJumin2.Text = adoSet.Fields("Jumin2").Value & ""
        frmMain.txtBirthDate.Text = adoSet.Fields("BirthDate").Value & ""
        Call adoSetClose(adoSet)
    End If
    
    frmMain.txtAgeYY.Text = SetAge_Check(frmMain.txtJumin1.Text, frmMain.txtJumin2.Text)
    
    Exit Sub
    

    
Hand_Flag_Set:
    ssGeneral.Row = Row
    ssGeneral.Col = 1
    If ssGeneral.CellType = CellTypeButton Then
        ssGeneral.TypeButtonPicture = LoadPicture("c:\twhis\src60\ocs\lab\data\fingerr.bmp")
        ssGeneral.Row = Row
        ssGeneral.Row2 = Row
        ssGeneral.Col = 2
        ssGeneral.Col2 = ssGeneral.DataColCnt
        ssGeneral.BlockMode = True
        ssGeneral.ForeColor = RGB(192, 0, 220)
        ssGeneral.BlockMode = False
    End If
    Return

Spread_ssEnrol_Clear:
    frmMain.ssEnrol.ReDraw = False
    frmMain.ssEnrol.MaxRows = 0
    frmMain.ssEnrol.MaxRows = 100
    frmMain.ssEnrol.RowHeight(-1) = 11
    frmMain.ssEnrol.ReDraw = True
    Return
    
Get_General_Sub_Data:
    Dim sCompareText        As String
    Dim sCompRoutine        As String
    Dim strSpace3           As String
    
    strSpace3 = "   "
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID RwID, b.Codenm SLname, c.itemNM, "
    StrSql = StrSql & "        d.JeobsuT1, d.JeobsuT2, d.GeomchCD, d.GeomsaGu,"
    StrSql = StrSql & "        d.CmDoctor, d.RoomCode, d.Deptcode, d.Gbio, d.DrCode, d.GbCh, d.GbEr,"
    StrSql = StrSql & "        TO_CHAR(d.Indate,  'YYYY-MM-DD') Indate,"
    StrSql = StrSql & "        TO_CHAR(d.OrderDt, 'YYYY-MM-DD') OrderDt"
    StrSql = StrSql & " FROM   TWEXAM_General_Sub a,"
    StrSql = StrSql & "        TWEXAM_Specode     b,"
    StrSql = StrSql & "        TWEXAM_ItemML      c,"
    StrSql = StrSql & "        TWEXAM_General     d "
    StrSql = StrSql & " WHERE  a.JeobsuDt = TO_DATE( '" & sJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.SLipno1  =  " & nSLipno1
    StrSql = StrSql & " AND    a.SLipno2  =  " & nSLipno2
    StrSql = StrSql & " AND    b.Codegu   = '12'"
    StrSql = StrSql & " AND    a.SLipno1  = TO_Number(b.Codeky)"
    StrSql = StrSql & " AND    a.itemCD   = c.Codeky(+)"
    StrSql = StrSql & " AND    a.JeobsuDt = d.JeobsuDt(+)"
    StrSql = StrSql & " AND    a.SLipno1  = d.SLipno1(+)"
    StrSql = StrSql & " AND    a.SLipno2  = d.SLipno2(+)"
    StrSql = StrSql & " ORDER  BY a.SLipno1, a.ITEMCD"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        frmMain.ssEnrol.Row = frmMain.ssEnrol.DataRowCnt + 1
        
        If adoSet.Fields("itemCd").Value <> adoSet.Fields("RoutinCD") Then
            If sCompRoutine <> adoSet.Fields("RoutinCD") Then
                GoSub Routine_Expaned
                frmMain.ssEnrol.Row = frmMain.ssEnrol.DataRowCnt + 1
            End If
        End If
        
        If sCompareText <> adoSet.Fields("SLipno1").Value Then
            frmMain.ssEnrol.Col = 3:  frmMain.ssEnrol.Text = adoSet.Fields("SLname").Value & ""
        End If
        
        
        frmMain.ssEnrol.Col = 2:  frmMain.ssEnrol.Text = adoSet.Fields("SLipno1").Value & ""
        frmMain.ssEnrol.Col = 4:  frmMain.ssEnrol.Text = adoSet.Fields("iTemCd").Value & ""
        If adoSet.Fields("itemCd").Value <> adoSet.Fields("RoutinCD") Then
            frmMain.ssEnrol.Col = 5:  frmMain.ssEnrol.Text = strSpace3 & adoSet.Fields("iTemNM").Value & ""
        Else
            frmMain.ssEnrol.Col = 5:  frmMain.ssEnrol.Text = adoSet.Fields("iTemNM").Value & ""
        End If
        
        frmMain.ssEnrol.Col = 6:  frmMain.ssEnrol.Text = adoSet.Fields("SLipno2").Value & ""
        frmMain.ssEnrol.Col = 7:  frmMain.ssEnrol.Text = ""
        frmMain.ssEnrol.Col = 8:  frmMain.ssEnrol.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                                         adoSet.Fields("JeobsuT2").Value
        frmMain.ssEnrol.Col = 9:  frmMain.ssEnrol.Text = adoSet.Fields("SLipno1").Value & ""
        frmMain.ssEnrol.Col = 10: frmMain.ssEnrol.Text = adoSet.Fields("GeomchCD").Value & ""
        frmMain.ssEnrol.Col = 11: frmMain.ssEnrol.Text = adoSet.Fields("GeomsaGu").Value & ""
        frmMain.ssEnrol.Col = 12: frmMain.ssEnrol.Text = adoSet.Fields("OrderDt").Value & ""
        frmMain.ssEnrol.Col = 13: frmMain.ssEnrol.Text = adoSet.Fields("Orderno").Value & ""
        frmMain.ssEnrol.Col = 14: frmMain.ssEnrol.Text = adoSet.Fields("CmDoctor").Value & ""
        frmMain.ssEnrol.Col = 15: frmMain.ssEnrol.Text = adoSet.Fields("Indate").Value & ""
        frmMain.ssEnrol.Col = 16: frmMain.ssEnrol.Text = adoSet.Fields("RoomCode").Value & ""
        frmMain.ssEnrol.Col = 17: frmMain.ssEnrol.Text = adoSet.Fields("DeptCode").Value & ""
        frmMain.ssEnrol.Col = 18: frmMain.ssEnrol.Text = adoSet.Fields("Gbio").Value & ""
        frmMain.ssEnrol.Col = 19: frmMain.ssEnrol.Text = adoSet.Fields("Drcode").Value & ""
        frmMain.ssEnrol.Col = 22: frmMain.ssEnrol.Text = adoSet.Fields("Bi").Value & ""
        frmMain.ssEnrol.Col = 23: frmMain.ssEnrol.Text = adoSet.Fields("GbEr").Value & ""
        frmMain.ssEnrol.Col = 24: frmMain.ssEnrol.Text = adoSet.Fields("GbCh").Value & ""
        frmMain.ssEnrol.Col = 25: frmMain.ssEnrol.Text = sJeobsuDt
        frmMain.ssEnrol.Col = 28: frmMain.ssEnrol.Text = adoSet.Fields("RoutinCd").Value & ""
        
        sCompareText = adoSet.Fields("SLipno1").Value & ""
        sCompRoutine = adoSet.Fields("RoutinCD").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

Routine_Expaned:
    Dim adoRcd      As ADODB.Recordset
    Dim sRcode      As String
    
    sRcode = adoSet.Fields("RoutinCD").Value & ""
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Routine"
    StrSql = StrSql & " WHERE  RoutinCD = '" & sRcode & "'"
    If False = adoSetOpen(StrSql, adoRcd) Then Return
    frmMain.ssEnrol.Row = frmMain.ssEnrol.DataRowCnt + 1
    
    frmMain.ssEnrol.Col = 2:  frmMain.ssEnrol.Text = adoSet.Fields("SLipno1").Value & ""
    frmMain.ssEnrol.Col = 3:  frmMain.ssEnrol.Text = adoSet.Fields("SLname").Value & ""

    'frmMain.ssEnrol.Col = 4: frmMain.ssEnrol.Text = adoRcd.Fields("RoutinCD").Value & ""
    
    frmMain.ssEnrol.Col = 4: frmMain.ssEnrol.Text = adoSet.Fields("SLipno1").Value & ""
    frmMain.ssEnrol.Col = 5: frmMain.ssEnrol.Text = adoRcd.Fields("RoutinNM").Value & ""
    sCompareText = adoSet.Fields("SLipno1").Value & ""
    Call adoSetClose(adoRcd)
    Return
End Sub

Private Sub ssGeneral_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If Col > 0 Then
            GoSub General_Sort_Sub
        End If
    End If
    Exit Sub
    
    
General_Sort_Sub:
    ssGeneral.Row = 1
    ssGeneral.Row2 = ssGeneral.DataRowCnt
    ssGeneral.Col = 1
    ssGeneral.Col2 = ssGeneral.DataColCnt
    ssGeneral.SortBy = SortByRow
    ssGeneral.SortKey(1) = Col
    
    If ssGeneral.SortKeyOrder(1) = SortKeyOrderAscending Then
        ssGeneral.SortKeyOrder(1) = SortKeyOrderDescending
    Else
        ssGeneral.SortKeyOrder(1) = SortKeyOrderAscending
    End If
    ssGeneral.Action = ActionSort
    
    Return
    
End Sub

Private Sub ssGeneral_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    ssGeneral.ReDraw = False
    ssGeneral.Row = Row
    ssGeneral.Row2 = Row
    ssGeneral.Col = 2
    ssGeneral.Col2 = ssGeneral.DataColCnt
    ssGeneral.BlockMode = True
    ssGeneral.ForeColor = RGB(0, 0, 0)
    ssGeneral.BlockMode = False
    ssGeneral.ReDraw = True

    ssGeneral.Row = Row
    ssGeneral.Col = 1
    ssGeneral.TypeButtonPicture = LoadPicture("")

End Sub
