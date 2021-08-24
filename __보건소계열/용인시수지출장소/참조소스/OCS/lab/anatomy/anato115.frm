VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anato_Dyeing_Persons 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Æ¯¼ö°Ë»çÈ¯ÀÚÁ¶È¸"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   1665
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8325
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin FPSpread.vaSpread ssResult 
      Height          =   6870
      Left            =   60
      TabIndex        =   0
      Top             =   1815
      Width           =   10155
      _Version        =   196608
      _ExtentX        =   17912
      _ExtentY        =   12118
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   11
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421376
      MaxCols         =   12
      MaxRows         =   600
      Protect         =   0   'False
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "ANATO115.frx":0000
      VisibleCols     =   12
      VisibleRows     =   500
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'À§ ¸ÂÃã
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12060
      _Version        =   65536
      _ExtentX        =   21272
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Æ¯  ¼ö  °Ë  »ç  È¯  ÀÚ  Á¶  È¸"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   16.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1908
      Left            =   10296
      TabIndex        =   2
      Top             =   1080
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   3360
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1440
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24379395
         CurrentDate     =   36311
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24379395
         CurrentDate     =   36311
      End
      Begin VB.Label Label2 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00808000&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "Á¢¼öÀÏÀÚ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   1200
         Width           =   210
      End
   End
   Begin Threed.SSCommand cmdView 
      Height          =   900
      Left            =   10296
      TabIndex        =   7
      Top             =   3312
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2984
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "Á¶ È¸"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      AutoSize        =   1
      Picture         =   "ANATO115.frx":20AD
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   900
      Left            =   10296
      TabIndex        =   6
      Top             =   4320
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2984
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "Á¾ ·á"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      AutoSize        =   1
      Picture         =   "ANATO115.frx":24FF
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   48
      TabIndex        =   10
      Top             =   1008
      Width           =   9948
      _Version        =   65536
      _ExtentX        =   17547
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "Æ¯¼ö°Ë»çÇ×¸ñ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   1
         Left            =   324
         TabIndex        =   11
         Top             =   312
         Width           =   1548
         _Version        =   65536
         _ExtentX        =   2730
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "Æ¯¼ö¿°»ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   2
         Left            =   1881
         TabIndex        =   12
         Top             =   312
         Width           =   1548
         _Version        =   65536
         _ExtentX        =   2730
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "¸é¿ª¿°»ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   3
         Left            =   3438
         TabIndex        =   13
         Top             =   312
         Width           =   1548
         _Version        =   65536
         _ExtentX        =   2730
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "¸é¿ªÇü±¤¿°»ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   4
         Left            =   4995
         TabIndex        =   14
         Top             =   312
         Width           =   1548
         _Version        =   65536
         _ExtentX        =   2730
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "È¿¼Ò¿°»ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   5
         Left            =   6552
         TabIndex        =   15
         Top             =   312
         Width           =   1548
         _Version        =   65536
         _ExtentX        =   2730
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "ÀüÀÚÇö¹Ì°æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   225
         Index           =   6
         Left            =   8115
         TabIndex        =   16
         Top             =   315
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Flow Cytometry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Anato_Dyeing_Persons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdView_Click()
    
    Dim rs                  As ADODB.Recordset
    
    Dim LsPtNo              As String * 8
    Dim LsStatus            As String * 1
    Dim LsCodeKy            As String
    Dim LsDrCode            As String * 6
    Dim LsDeptCode          As String * 4
    Dim LiReccnt            As Integer
    Dim i                   As Integer
    Dim LsRet
    
    Dim SpecialS            As String
    Dim SpecialE            As String
    
'    Call SSInitialize_H(ssResult)

    ssResult.Col = 1:      ssResult.Col2 = ssResult.DataColCnt
    ssResult.Row = 1:      ssResult.Row2 = ssResult.DataRowCnt
    
    ssResult.BlockMode = True
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.RowHeight(i) = 12
    Next i
    
    ssResult.Action = SS_ACTION_CLEAR_TEXT
    ssResult.ForeColor = RGB(0, 0, 0)
    ssResult.BlockMode = False
    
    ssResult.Col = 1:      ssResult.Row = 1
    ssResult.Action = SS_ACTION_ACTIVE_CELL
    
    gSFrDate = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.* , "
    strSQL = strSQL & "        TO_CHAR(a.Jdate, 'YYYY-MM-DD') Jdate1,"
    strSQL = strSQL & "        b.Deptnamek, c.Drname"
    strSQL = strSQL & " FROM   TWANAT_DIAG  a,"
    strSQL = strSQL & "        TWBAS_DEPT   b,"
    strSQL = strSQL & "        TWBAS_DOCTOR c "
    strSQL = strSQL & " WHERE  a.Jdate   >= TO_DATE('" & gSFrDate & "','YYYY-MM-DD')"
    strSQL = strSQL & " AND    a.Jdate   <= TO_DATE('" & gSToDate & "','YYYY-MM-DD')"
    Select Case GCodegu
           Case "83" '55"
                    ssResult.Row = 0: ssResult.Col = 7:   ssResult.Text = "Æ¯ ¼ö ¿° »ö"
                    SpecialS = "853001"
                    SpecialE = "853999"
                    
                    For i = 1 To 30
                        Select Case i
                               Case 1
                                    strSQL = strSQL & " AND  (  a.SPECIAL" & Format(i, "00") & " BETWEEN '853001' AND '853999' "
                               Case 2 To 29
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '853001' AND '853999' "
                               Case 30
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '853001' AND '853999' )"
                        End Select
                    Next i
           
           Case "87" '56"
                    ssResult.Row = 0: ssResult.Col = 7:   ssResult.Text = "¸é ¿ª ¿° »ö"
                    SpecialS = "857001"
                    SpecialE = "857999"
           
                    For i = 1 To 30
                        Select Case i
                               Case 1
                                    strSQL = strSQL & " AND  (  a.SPECIAL" & Format(i, "00") & " BETWEEN '857001' AND '857999' "
                               Case 2 To 29
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '857001' AND '857999' "
                               Case 30
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '857001' AND '857999' )"
                        End Select
                    Next i
           
           Case "84" '57"
                    ssResult.Row = 0: ssResult.Col = 7:   ssResult.Text = "¸é ¿ª Çü ±¤ ¿°»ö"
                    SpecialS = "854001"
                    SpecialE = "854999"
                    
                    For i = 1 To 30
                        Select Case i
                               Case 1
                                    strSQL = strSQL & " AND  (  a.SPECIAL" & Format(i, "00") & " BETWEEN '854001' AND '854999' "
                               Case 2 To 29
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '854001' AND '854999' "
                               Case 30
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '854001' AND '854999' )"
                        End Select
                    Next i
           
           Case "86" '58"
                    ssResult.Row = 0: ssResult.Col = 7:   ssResult.Text = "È¿ ¼Ò ¿° »ö"
                    SpecialS = "856001"
                    SpecialE = "856999"
                    
                    For i = 1 To 30
                        Select Case i
                               Case 1
                                    strSQL = strSQL & " AND  (  a.SPECIAL" & Format(i, "00") & " BETWEEN '856001' AND '856999' "
                               Case 2 To 29
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '856001' AND '856999' "
                               Case 30
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '856001' AND '856999' )"
                        End Select
                    Next i
           
           Case "85"
                    ssResult.Row = 0: ssResult.Col = 7:   ssResult.Text = "ÀüÀÚÇö¹Ì°æ"
                    SpecialS = "855001"
                    SpecialE = "855001"
                    
                    For i = 1 To 30
                        Select Case i
                               Case 1
                                    strSQL = strSQL & " AND  (  a.SPECIAL" & Format(i, "00") & " BETWEEN '855001' AND '855999' "
                               Case 2 To 29
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '855001' AND '855999' "
                               Case 30
                                    strSQL = strSQL & "  OR    a.SPECIAL" & Format(i, "00") & " BETWEEN '855001' AND '855999' )"
                        End Select
                    Next i
           
           Case "82"
                    strSQL = strSQL & " AND    a.Flow = 'Y' "
                    ssResult.Row = 0: ssResult.Col = 7:   ssResult.Text = "Flow Cytometry"
    End Select
    strSQL = strSQL & " AND    a.Deptcode = b.Deptcode(+)"
    strSQL = strSQL & " AND    a.DrCode   = c.DrCode(+)"
    strSQL = strSQL & " ORDER  BY JDATE1, a.CLASS, a.DATEYY, a.SEQNUM"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Dim strSpecial_Code         As String
    Dim strSpecial              As String
    Dim strSpecial_V            As String
    
    Do Until rs.EOF
        
        strSpecial = ""
        
        If (rs.Fields("flow").Value & "") <> "" And GCodegu = "82" Then
        
            ssResult.MaxRows = ssResult.DataRowCnt + 1
            ssResult.Row = ssResult.DataRowCnt + 1
            ssResult.Col = 1:  ssResult.Text = ssResult.Row
            ssResult.Col = 2:  ssResult.Text = rs.Fields("Class").Value & "-" & _
                                               rs.Fields("DateYY").Value & "-" & _
                                               rs.Fields("SeqNum").Value & ""
            ssResult.Col = 3:  ssResult.Text = rs.Fields("Ptno").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("Sname").Value & ""
            ssResult.Col = 5:  ssResult.Text = IIf(rs.Fields("Sex").Value & "" = "M", "³²", "¿©")
            ssResult.Col = 6:  ssResult.Text = rs.Fields("AgeYY").Value & ""
            
            ssResult.Col = 7:  ssResult.Text = rs.Fields("flow").Value & ""
            ssResult.Col = 8:  ssResult.Text = rs.Fields("Jdate").Value & ""
            ssResult.Col = 9:  ssResult.Text = rs.Fields("RoomCode").Value & ""
            ssResult.Col = 10: ssResult.Text = rs.Fields("Deptnamek").Value & ""
            
            ssResult.Col = 11: ssResult.Text = rs.Fields("Drname").Value & ""
        
        ElseIf GCodegu = "85" Then
            
            ssResult.MaxRows = ssResult.DataRowCnt + 1
            ssResult.Row = ssResult.DataRowCnt + 1
            ssResult.Col = 1:  ssResult.Text = ssResult.Row
            ssResult.Col = 2:  ssResult.Text = rs.Fields("Class").Value & "-" & _
                                               rs.Fields("DateYY").Value & "-" & _
                                               rs.Fields("SeqNum").Value & ""
            ssResult.Col = 3:  ssResult.Text = rs.Fields("Ptno").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("Sname").Value & ""
            ssResult.Col = 5:  ssResult.Text = IIf(rs.Fields("Sex").Value & "" = "M", "³²", "¿©")
            ssResult.Col = 6:  ssResult.Text = rs.Fields("AgeYY").Value & ""
            
            ssResult.Col = 7:  ssResult.Text = "Y" 'rs.Fields("flow").Value & ""
            ssResult.Col = 8:  ssResult.Text = rs.Fields("Jdate").Value & ""
            ssResult.Col = 9:  ssResult.Text = rs.Fields("RoomCode").Value & ""
            ssResult.Col = 10: ssResult.Text = rs.Fields("Deptnamek").Value & ""
            
            ssResult.Col = 11: ssResult.Text = rs.Fields("Drname").Value & ""
        
        Else
            For i = 1 To 30
                strSpecial_V = "Special" & Format(i, "00")
                
                strSpecial_Code = rs.Fields(strSpecial_V).Value & ""
                strSpecial = Special_Load(strSpecial_Code)
            
                If strSpecial <> "" Then
                    If i = 1 Then
                        ssResult.MaxRows = ssResult.DataRowCnt + 1
                        ssResult.Row = ssResult.DataRowCnt + 1
                        ssResult.Col = 1:  ssResult.Text = ssResult.Row
                        ssResult.Col = 2:  ssResult.Text = rs.Fields("Class").Value & "-" & _
                                                           rs.Fields("DateYY").Value & "-" & _
                                                           rs.Fields("SeqNum").Value & ""
                        ssResult.Col = 3:  ssResult.Text = rs.Fields("Ptno").Value & ""
                        ssResult.Col = 4:  ssResult.Text = rs.Fields("Sname").Value & ""
                        ssResult.Col = 5:  ssResult.Text = IIf(rs.Fields("Sex").Value & "" = "M", "³²", "¿©")
                        ssResult.Col = 6:  ssResult.Text = rs.Fields("AgeYY").Value & ""
                        
                        If strSpecial_Code >= SpecialS And strSpecial_Code <= SpecialE Then
                            ssResult.Col = 7:  ssResult.Text = ssResult.Text & strSpecial & vbCr
                            ssResult.RowHeight(ssResult.Row) = ssResult.MaxTextRowHeight(ssResult.Row)
                        End If
                        ssResult.Col = 8:  ssResult.Text = rs.Fields("Jdate").Value & ""
                        ssResult.Col = 9:  ssResult.Text = rs.Fields("RoomCode").Value & ""
                        ssResult.Col = 10: ssResult.Text = rs.Fields("Deptnamek").Value & ""
                        ssResult.Col = 11: ssResult.Text = rs.Fields("Drname").Value & ""
                    Else
                        If strSpecial_Code >= SpecialS And strSpecial_Code <= SpecialE Then
                            ssResult.Col = 7:  ssResult.Text = ssResult.Text & strSpecial & vbCr
                            ssResult.RowHeight(ssResult.Row) = ssResult.MaxTextRowHeight(ssResult.Row)
                        End If
                    End If
                Else
                    Exit For
                End If
            
            Next i
        
        End If
        
        rs.MoveNext
    Loop
    
    ssResult.MaxRows = ssResult.DataRowCnt + 1
        
    AdoCloseSet rs
    
    
End Sub


Private Sub Form_Load()
    
    dtFromJeobsu.Value = Format(CDate(Dual_Date_Get("yyyy-MM-dd")) - 7, "yyyy-MM-dd")
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    SSOption1(1).Value = True
    
End Sub


Private Sub SSOption1_Click(Index As Integer, Value As Integer)

    Select Case Index
           Case 1
                GCodegu = "83"      'Æ¯¼ö¿°»ö
           Case 2
                GCodegu = "87"      '¸é¿ª¿°»ö
           Case 3
                GCodegu = "84"      '¸é¿ªÇü±¤¿°»ö
           Case 4
                GCodegu = "86"      'È¿¼Ò¸é¿ª
           Case 5
                GCodegu = "85"      'Electroscope
           Case 6
                GCodegu = "82"      'Flow Cytometry
    End Select


End Sub
