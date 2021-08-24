VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmitemadd 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "추가검사항목"
   ClientHeight    =   6810
   ClientLeft      =   15
   ClientTop       =   1680
   ClientWidth     =   6315
   Icon            =   "frmItemAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin FPSpreadADO.fpSpread ssItemList 
      Height          =   6675
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9551
      _ExtentY        =   11774
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
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
      MaxCols         =   13
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SpreadDesigner  =   "frmItemAdd.frx":030A
      UserResize      =   0
      Appearance      =   2
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   6660
      Left            =   5535
      ScaleHeight     =   6600
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   90
      Width           =   675
      Begin Threed.SSCommand cmdItemAdd 
         Height          =   900
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1588
         _StockProps     =   78
         Caption         =   "추가"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "frmItemAdd.frx":1D33
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   600
         Left            =   0
         TabIndex        =   1
         Top             =   900
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   8.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "frmItemAdd.frx":204D
      End
   End
End
Attribute VB_Name = "frmitemadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub LIST_PROCESS()
    Dim i           As Integer
    Dim LiPos       As Integer
    Dim LiLen       As Integer
    Dim LiSlipNo1   As Integer
    Dim adoList     As ADODB.Recordset
    
    
    LiSlipNo1 = Val(Left(frmResult.cmbSLip, 2))
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT CODEKY, ITEMNM, SUGACD, YAGEO, MINCHAM, MAXCHAM, "
    gStrSql = gStrSql & "        DELTAQC, DELTAMIN, DELTAMAX,DANWI "
    gStrSql = gStrSql & " FROM  TWEXAM_ITEMML "
    gStrSql = gStrSql & " WHERE SubSTR(Codeky, 1,2) = '" & LiSlipNo1 & "'"
    gStrSql = gStrSql & " ORDER BY Codeky "
    
   
    If False = adoSetOpen(gStrSql, adoList) Then Exit Sub
    
    Do Until adoList.EOF
        ssItemList.Row = ssItemList.DataRowCnt + 1
        ssItemList.Col = 2:     ssItemList.Text = adoList.Fields("SUGACD").Value & ""
        ssItemList.Col = 3:     ssItemList.Text = adoList.Fields("ITEMNM").Value & ""
        ssItemList.Col = 4:     ssItemList.Text = adoList.Fields("CODEKY").Value & ""
        ssItemList.Col = 5:     ssItemList.Text = adoList.Fields("MINCHAM").Value & ""
        ssItemList.Col = 6:     ssItemList.Text = adoList.Fields("MAXCHAM").Value & ""
        ssItemList.Col = 7:     ssItemList.Text = adoList.Fields("DANWI").Value & ""
        ssItemList.Col = 8:     ssItemList.Text = adoList.Fields("DELTAQC").Value & ""
        ssItemList.Col = 9:     ssItemList.Text = adoList.Fields("DELTAMIN").Value & ""
        ssItemList.Col = 10:    ssItemList.Text = adoList.Fields("DELTAMAX").Value & ""
        ssItemList.Col = 13:    ssItemList.Text = adoList.Fields("YAGEO").Value & ""
        
        adoList.MoveNext
    Loop
    Call adoSetClose(adoList)
    

End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub


Private Sub cmdItemAdd_Click()

    Dim i                As Integer
    Dim j                As Integer
    Dim LbExist          As Boolean
    Dim LbAddOk          As Boolean
    Dim LsItemCD         As String
    
    
    GoSub SUB_SLIP1_PROC
    
    
    Exit Sub
    
'/======================================================================/
SUB_SLIP1_PROC:

    For i = 1 To ssItemList.DataRowCnt
        ssItemList.Row = i
        
        ssItemList.Col = 1
        If ssItemList.Text = "1" Then
            ssItemList.Col = 4: LsItemCD = Trim$(ssItemList.Text)
            LbExist = False                     '   Not Exist
            
            '기존의 자료있는지 확인
'            For j = 1 To frmResult.sprSLip.ssSlip.DataRowCnt
'                frmResult.sprSLip.ssSlip.Col = 12
'                frmResult.sprSLip.ssSlip.Row = j
'                If Trim(frmResult.sprSLip.ssSlip.Text) = Trim(ssItemList.Text) Then
'                    LbExist = True              '   Exist
'                    Exit For
'                End If
'            Next j
 
            If LbExist = False Then
            
                GoSub SUB_ITEMADD_PROC
                
                If LbAddOk = True Then
                    ssItemList.Row = i
                    frmResult.sprSLip.ssSlip.Row = frmResult.sprSLip.ssSlip.DataRowCnt + 1
                    frmResult.sprSLip.ssSlip.Col = 1:   ssItemList.Col = 3:  frmResult.sprSLip.ssSlip.Text = Trim(ssItemList.Text)  ' Item Name
                    frmResult.sprSLip.ssSlip.Col = 12:  ssItemList.Col = 4:  frmResult.sprSLip.ssSlip.Text = Trim(ssItemList.Text)  ' Item code
                    frmResult.sprSLip.ssSlip.Col = 4:   ssItemList.Col = 5:  frmResult.sprSLip.ssSlip.Text = Trim(ssItemList.Text)  ' Min
                    frmResult.sprSLip.ssSlip.Col = 5:   ssItemList.Col = 6:  frmResult.sprSLip.ssSlip.Text = Trim(ssItemList.Text)  ' Max
                    frmResult.sprSLip.ssSlip.Col = 6:   ssItemList.Col = 7:  frmResult.sprSLip.ssSlip.Text = Trim(ssItemList.Text)  ' Unit
                    frmResult.sprSLip.ssSlip.Col = 13:  ssItemList.Col = 8:  frmResult.sprSLip.ssSlip.Text = Trim(ssItemList.Text)  ' Delta QC
                    frmResult.sprSLip.ssSlip.Col = 14:  ssItemList.Col = 9:  frmResult.sprSLip.ssSlip.Text = ssItemList.Text        ' Delta Min
                    frmResult.sprSLip.ssSlip.Col = 15:  ssItemList.Col = 10: frmResult.sprSLip.ssSlip.Text = ssItemList.Text        ' Delta Max
                    frmResult.sprSLip.ssSlip.Col = 16:  ssItemList.Col = 11: frmResult.sprSLip.ssSlip.Text = "0"                    ' OrderNo
                    frmResult.sprSLip.ssSlip.Col = 12:  ssItemList.Col = 12: frmResult.sprSLip.ssSlip.Text = ssItemList.Text        ' Rowid
                End If
            End If
        End If
    Next i
    
    Unload Me
            
    Return
    
    
'/======================================================================/
        
SUB_ITEMADD_PROC:
    Dim LsJeobSuDt       As String
    Dim LiSlipNo1        As Integer
    Dim LiSlipNo2        As Integer
    Dim LsPtNo           As String
    Dim LsSex            As String
    Dim LiAgeYY          As Integer
    Dim LiAgeMM          As Integer
    Dim LsBi             As String
    Dim LsGbJeobSu       As String
    Dim LsMatchno        As Integer
    Dim LsDaySeq         As Integer
    
    Dim sRowID           As String
    
    

    
    frmResult.sprSLip.Row = 1
    frmResult.sprSLip.Col = 12
    sRowID = frmResult.sprSLip.Text
    
    
    
    gStrSql = "  SELECT a.*, TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt FROM TWEXAM_GENERAL_SUB a WHERE ROWID = '" & sRowID & "'"
    If adoSetOpen(gStrSql, adoSet) Then
        LsJeobSuDt = adoSet.Fields("JEOBSUDT").Value & ""
        LiSlipNo1 = Val(adoSet.Fields("SLIPNO1").Value & "")
        LiSlipNo2 = Val(adoSet.Fields("SLIPNO2").Value & "")
        
        LsBi = adoSet.Fields("BI").Value & ""
        LsPtNo = adoSet.Fields("PTNO").Value & ""
        LsSex = adoSet.Fields("SEX").Value & ""
        LiAgeYY = Val(adoSet.Fields("AGEYY").Value & "")
        LiAgeMM = Val(adoSet.Fields("AGEMM").Value & "")
        LsGbJeobSu = adoSet.Fields("GBJEOBSU").Value & ""
        LsMatchno = Val(adoSet.Fields("Matchno").Value & "")
        LsDaySeq = Val(adoSet.Fields("DaySeq").Value & "")
        Call adoSetClose(adoSet)
    End If
    
    
    gStrSql = ""
    gStrSql = gStrSql & " INSERT INTO TWEXAM_General_Sub"
    gStrSql = gStrSql & "       (  Jeobsudt,    Slipno1,    Slipno2,    Routincd,    Codeky1,   "
    gStrSql = gStrSql & "          Itemcd,      Verify,     Result1,    Result2,     Result3,   "
    gStrSql = gStrSql & "          Result4,     Result5,    PtNo,       Sex,         AgeYY,     "
    gStrSql = gStrSql & "          Agemm,       Bi,         GbHost,     GbJeobsu,    OrderNo,"
    gStrSql = gStrSql & "          DaySeq,      Matchno) "
    gStrSql = gStrSql & " VALUES( TO_DATE('" & LsJeobSuDt & "','yyyy-MM-dd'),"
    gStrSql = gStrSql & "          " & LiSlipNo1 & ","
    gStrSql = gStrSql & "          " & LiSlipNo2 & ","
    gStrSql = gStrSql & "         '" & Trim(LsItemCD) & "',"                      'Routincd
    gStrSql = gStrSql & "          " & LiSlipNo1 & ","
    gStrSql = gStrSql & "         '" & Trim(LsItemCD) & "',"
    gStrSql = gStrSql & "         'N',"
    gStrSql = gStrSql & "         ' ',' ',' ',' ',' ',"      'Result
    gStrSql = gStrSql & "         '" & LsPtNo & "',"
    gStrSql = gStrSql & "         '" & LsSex & "',"
    gStrSql = gStrSql & "          " & LiAgeYY & ","
    gStrSql = gStrSql & "          " & LiAgeMM & ","
    gStrSql = gStrSql & "         '" & LsBi & "',"
    gStrSql = gStrSql & "         '2',"
    gStrSql = gStrSql & "         '" & Trim(LsGbJeobSu) & "',"
    gStrSql = gStrSql & "          0,"
    gStrSql = gStrSql & "          " & LsDaySeq & ","
    gStrSql = gStrSql & "          " & LsMatchno & ")"
    
    adoConnect.BeginTrans
    If adoExec(gStrSql) Then
        adoConnect.CommitTrans
        
        gStrSql = ""
        gStrSql = gStrSql & " SELECT Orderno, RowID RWID"
        gStrSql = gStrSql & " FROM   TWEXAM_General_Sub"
        gStrSql = gStrSql & " WHERE  JeobsuDt = TO_DATE('" & LsJeobSuDt & "','yyyy-MM-dd')"
        gStrSql = gStrSql & " AND    SLipno1  =  " & LiSlipNo1 & ""
        gStrSql = gStrSql & " AND    SLipno1  =  " & LiSlipNo2 & ""
        gStrSql = gStrSql & " AND    iTemCd   = '" & LsItemCD & "'"
        gStrSql = gStrSql & " AND    Verify   <> 'Y'"
        If adoSetOpen(strSql, adoSet) Then
            ssItemList.Row = i
            ssItemList.Col = 11:    ssItemList.Text = adoSet.Fields("ORDERNO").Value & ""
            ssItemList.Col = 12:    ssItemList.Text = adoSet.Fields("RWID").Value & ""
            Call adoSetClose(adoSet)
            LbAddOk = True
        End If
    Else
        adoConnect.RollbackTrans
        LbAddOk = False
    End If
        
    Return
    
    
    
End Sub

Private Sub OptSort_Click(Index As Integer)

    Call LIST_PROCESS
    
End Sub

Private Sub Form_Load()

    Call LIST_PROCESS
    
End Sub

Private Sub ssItemList_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    
    KeyAscii = 0
    
    SendKeys "{tab}"
    
End Sub


