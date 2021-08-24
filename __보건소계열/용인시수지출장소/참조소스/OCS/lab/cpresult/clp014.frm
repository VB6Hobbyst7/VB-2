VERSION 4.00
Begin VB.Form clp_item_add 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "추가검사항목"
   ClientHeight    =   6630
   ClientLeft      =   2625
   ClientTop       =   1530
   ClientWidth     =   6195
   Height          =   7035
   Icon            =   "CLP014.frx":0000
   Left            =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Top             =   1185
   Width           =   6315
   Begin VB.OptionButton OptSort 
      Caption         =   "&2. 검사코드순"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   30
      Width           =   1725
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "&1. 검사명순"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   1365
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   6255
      Left            =   5460
      ScaleHeight     =   6195
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   330
      Width           =   675
      Begin Threed.SSCommand cmdItemAdd 
         Height          =   900
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   600
         _version        =   65536
         _extentx        =   1058
         _extenty        =   1588
         _stockprops     =   78
         caption         =   "추가"
         forecolor       =   8388608
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         roundedcorners  =   0   'False
         autosize        =   1
         picture         =   "CLP014.frx":030A
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   600
         Left            =   0
         TabIndex        =   2
         Top             =   900
         Width           =   600
         _version        =   65536
         _extentx        =   1058
         _extenty        =   1058
         _stockprops     =   78
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "돋움체"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         roundedcorners  =   0   'False
         autosize        =   1
         picture         =   "CLP014.frx":0624
      End
   End
   Begin FPSpread.vaSpread ssItemList 
      Height          =   6255
      Left            =   60
      TabIndex        =   0
      Top             =   330
      Width           =   5385
      _version        =   131077
      _extentx        =   9499
      _extenty        =   11033
      _stockprops     =   64
      backcolorstyle  =   1
      colheaderdisplay=   0
      displayrowheaders=   0   'False
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "굴림체"
         charset         =   129
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      maxcols         =   13
      maxrows         =   300
      rowheaderdisplay=   0
      scrollbars      =   2
      shadowcolor     =   12632256
      shadowdark      =   8421504
      shadowtext      =   0
      spreaddesigner  =   "CLP014.frx":093E
      userresize      =   1
      visiblecols     =   500
      visiblerows     =   500
   End
End
Attribute VB_Name = "clp_item_add"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Sub LIST_PROCESS()
    Dim i           As Integer
    Dim LiPos       As Integer
    Dim LiLen       As Integer
    Dim LiSlipNo1   As Integer
    
    LiPos = InStr(clp_result_mgr.lblSlipNo, "-")
    LiLen = Len(clp_result_mgr.lblSlipNo)
    LiSlipNo1 = Val(Mid(clp_result_mgr.lblSlipNo, 1, LiPos - 1))
    
    Result = dosql("Open Scope")
    
    GStrSql = " FOR ALL SELECT CODEKY, ITEMNM, SUGACD, YAGEO, MINCHAM, MAXCHAM, DELTAQC, DELTAMIN, DELTAMAX, "
    GStrSql = GStrSql & " DANWI "
    GStrSql = GStrSql & " FROM  TWEXAM_ITEMML "
    GStrSql = GStrSql & " WHERE SubSTR(Codeky, 1,2) = '" & LiSlipNo1 & "'"
    
    If OptSort(0).Value Then
        GStrSql = GStrSql & " ORDER BY ItemNm "
    ElseIf OptSort(1).Value Then
        GStrSql = GStrSql & " ORDER BY Codeky "
    End If
   
    Result = dosql(GStrSql)
    
    If rowindicator = 0 Then Exit Sub
    
    For i = 0 To rowindicator - 1
        ssItemList.Row = i + 1
        ssItemList.Col = 2:     ssItemList.Text = GlueGetString("SUGACD", i)
        ssItemList.Col = 3:     ssItemList.Text = GlueGetString("ITEMNM", i)
        ssItemList.Col = 4:     ssItemList.Text = GlueGetString("CODEKY", i)
        ssItemList.Col = 5:     ssItemList.Text = GlueGetString("MINCHAM", i)
        ssItemList.Col = 6:     ssItemList.Text = GlueGetString("MAXCHAM", i)
        ssItemList.Col = 7:     ssItemList.Text = GlueGetString("DANWI", i)
        ssItemList.Col = 8:     ssItemList.Text = GlueGetString("DELTAQC", i)
        ssItemList.Col = 9:     ssItemList.Text = GlueGetString("DELTAMIN", i)
        ssItemList.Col = 10:    ssItemList.Text = GlueGetString("DELTAMAX", i)
        ssItemList.Col = 13:    ssItemList.Text = GlueGetString("YAGEO", i)
    Next i
    
    Result = dosql("Close Scope")
 

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
    
    If GsSlipGubun = "SLIP2" Then
        GoSub SUB_SLIP2_PROC
    Else
        GoSub SUB_SLIP1_PROC
    End If
    
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
'            For j = 1 To clp_result_slip1.ssSlip.DataRowCnt
'                clp_result_slip1.ssSlip.Col = 12
'                clp_result_slip1.ssSlip.Row = j
'                If Trim(clp_result_slip1.ssSlip.Text) = Trim(ssItemList.Text) Then
'                    LbExist = True              '   Exist
'                    Exit For
'                End If
'            Next j
 
            If LbExist = False Then
            
                GoSub SUB_ITEMADD_PROC
                
                If LbAddOk = True Then
                    ssItemList.Row = i
                    clp_result_slip1.ssSlip.Row = clp_result_slip1.ssSlip.DataRowCnt + 1
                    clp_result_slip1.ssSlip.Col = 1:   ssItemList.Col = 3:  clp_result_slip1.ssSlip.Text = Trim(ssItemList.Text)  ' Item Name
                    clp_result_slip1.ssSlip.Col = 12:  ssItemList.Col = 4:  clp_result_slip1.ssSlip.Text = Trim(ssItemList.Text)  ' Item code
                    clp_result_slip1.ssSlip.Col = 4:   ssItemList.Col = 5:  clp_result_slip1.ssSlip.Text = Trim(ssItemList.Text)  ' Min
                    clp_result_slip1.ssSlip.Col = 5:   ssItemList.Col = 6:  clp_result_slip1.ssSlip.Text = Trim(ssItemList.Text)  ' Max
                    clp_result_slip1.ssSlip.Col = 6:   ssItemList.Col = 7:  clp_result_slip1.ssSlip.Text = Trim(ssItemList.Text)  ' Unit
                    clp_result_slip1.ssSlip.Col = 14:  ssItemList.Col = 8:  clp_result_slip1.ssSlip.Text = Trim(ssItemList.Text)  ' Delta QC
                    clp_result_slip1.ssSlip.Col = 15:  ssItemList.Col = 9:  clp_result_slip1.ssSlip.Text = ssItemList.Text        ' Delta Min
                    clp_result_slip1.ssSlip.Col = 16:  ssItemList.Col = 10: clp_result_slip1.ssSlip.Text = ssItemList.Text        ' Delta Max
                    clp_result_slip1.ssSlip.Col = 17:  ssItemList.Col = 11: clp_result_slip1.ssSlip.Text = "0"                    ' OrderNo
                    clp_result_slip1.ssSlip.Col = 13:  ssItemList.Col = 12: clp_result_slip1.ssSlip.Text = ssItemList.Text        ' Rowid
                End If
            End If
        End If
    Next i
    
    Unload Me
            
    Return
    
'/======================================================================/

SUB_SLIP2_PROC:

    For i = 1 To ssItemList.DataRowCnt
        ssItemList.Row = i
        
        ssItemList.Col = 1
        If ssItemList.Text = "1" Then
            ssItemList.Col = 4: LsItemCD = Trim$(ssItemList.Text)
            
            GoSub SUB_ITEMADD_PROC
            
            If LbAddOk = True Then
                clp_result_slip2.ssSlip.Row = clp_result_slip2.ssSlip.DataRowCnt + 1
                clp_result_slip2.ssSlip.Col = 1:  ssItemList.Col = 3:  clp_result_slip2.ssSlip.Text = Trim(ssItemList.Text)  ' Item Name
                clp_result_slip2.ssSlip.Col = 7:  ssItemList.Col = 4:  clp_result_slip2.ssSlip.Text = Trim(ssItemList.Text)  ' Item code
                clp_result_slip2.ssSlip.Col = 8:  ssItemList.Col = 11: clp_result_slip2.ssSlip.Text = "0"               ' OrderNo
                clp_result_slip2.ssSlip.Col = 9:  ssItemList.Col = 12: clp_result_slip2.ssSlip.Text = ssItemList.Text   ' Rowid
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
    
    Result = dosql("Open Scope")
    
    If GsSlipGubun = "SLIP1" Then
        clp_result_slip1.ssSlip.Row = 1
        clp_result_slip1.ssSlip.Col = 13
        GlueSetString "cROWID", 0, Trim$(clp_result_slip1.ssSlip.Text)
        
    ElseIf GsSlipGubun = "SLIP2" Then
        clp_result_slip2.ssSlip.Row = 1
        clp_result_slip2.ssSlip.Col = 9
        GlueSetString "cROWID", 0, Trim$(clp_result_slip2.ssSlip.Text)
        
    Else
        Exit Sub
    End If
    
    GStrSql = " FOR 1 SELECT * FROM TWEXAM_GENERAL_SUB "
    GStrSql = GStrSql & " WHERE ROWID  = :cROWID: "
    
    Result = dosql(GStrSql)
    
    If rowindicator = 1 Then
        LsJeobSuDt = GlueGetString("JEOBSUDT", 0)
        LiSlipNo1 = GlueGetNumber("SLIPNO1", 0)
        LiSlipNo2 = GlueGetNumber("SLIPNO2", 0)
        
        LsBi = GlueGetString("BI", 0)
        LsPtNo = GlueGetString("PTNO", 0)
        LsSex = GlueGetString("SEX", 0)
        LiAgeYY = GlueGetNumber("AGEYY", 0)
        LiAgeMM = GlueGetNumber("AGEMM", 0)
        LsGbJeobSu = GlueGetString("GBJEOBSU", 0)
    End If
    
    GlueSetString "cJEOBSUDT", 0, Trim(LsJeobSuDt)
    GlueSetnumber "cSLIPNO1", 0, LiSlipNo1
    GlueSetnumber "cCODEKY1", 0, LiSlipNo1
    GlueSetnumber "cSLIPNO2", 0, LiSlipNo2
    GlueSetString "cITEMCD", 0, Trim(LsItemCD)
    GlueSetString "cBI", 0, Trim(LsBi)
    GlueSetString "cPTNO", 0, Trim(LsPtNo)
    GlueSetString "cSEX", 0, Trim(LsSex)
    GlueSetnumber "cAGEYY", 0, LiAgeYY
    GlueSetnumber "cAGEMM", 0, LiAgeMM
    GlueSetnumber "cORDERNO", 0, 0
    GlueSetString "cGBJEOBSU", 0, Trim(LsGbJeobSu)
    GlueSetString "cRoutinCd", 0, " "
    GlueSetString "cGbHost", 0, "2"
    GlueSetString "cVERIFY", 0, "N"
    GlueSetString "cRESULT1", 0, " "
    GlueSetString "cRESULT2", 0, " "
    GlueSetString "cRESULT3", 0, " "
    GlueSetString "cRESULT4", 0, " "
    GlueSetString "cRESULT5", 0, " "
    GlueSetString "cRCODE1", 0, " "
    GlueSetString "cRCODE2", 0, " "
    GlueSetString "cRCODE3", 0, " "
    GlueSetString "cRCODE4", 0, " "
    GlueSetString "cRCODE5", 0, " "
    
    GStrSql = "INSERT INTO TWEXAM_GENERAL_SUB "
    GStrSql = GStrSql & "       (  Jeobsudt,    Slipno1,    Slipno2,    Routincd,    Codeky1,   "
    GStrSql = GStrSql & "          Itemcd,      Verify,     Result1,    Result2,     Result3,   "
    GStrSql = GStrSql & "          Result4,     Result5,    PtNo,       Sex,         AgeYY,     "
    GStrSql = GStrSql & "          Agemm,       Bi,         GbHost,     GbJeobsu,    OrderNo  ) "
    GStrSql = GStrSql & "VALUES (:cJeobsudt:,  :cSlipno1:, :cSlipno2:, :cRoutincd:, :cCodeky1:, "
    GStrSql = GStrSql & "        :cItemcd:,    :cVerify:,  :cResult1:, :cResult2:,  :cResult3:, "
    GStrSql = GStrSql & "        :cResult4:,   :cResult5:, :cPtNo:,    :cSex:,      :cAgeYY:,   "
    GStrSql = GStrSql & "        :cAgemm:,     :cBi:,      :cGbHost:,  :cGbJeobsu:, :cORDERNO: ) "
          
    Result = dosql(GStrSql)
    
    If Result = 0 Then
        Result = dosql("COMMIT")
        
        GStrSql = " FOR 1    SELECT ORDERNO, ROWID  FROM TWEXAM_GENERAL_SUB  "
        GStrSql = GStrSql & " WHERE JEOBSUDT = '" & LsJeobSuDt & "'           "
        GStrSql = GStrSql & " AND   SLIPNO1  = '" & LiSlipNo1 & "'            "
        GStrSql = GStrSql & " AND   SLIPNO2  = '" & LiSlipNo2 & "'            "
        GStrSql = GStrSql & " AND   ITEMCD   = '" & LsItemCD & "'             "
        GStrSql = GStrSql & " AND   VERIFY   <> 'Y'                           "
    
        Result = dosql(GStrSql)
    
        If rowindicator = 1 Then
            ssItemList.Row = i
            ssItemList.Col = 11:    ssItemList.Text = GlueGetNumber("ORDERNO", 0)
            ssItemList.Col = 12:    ssItemList.Text = GlueGetString("ROWID", 0)
            LbAddOk = True
        End If
    Else
       Result = dosql("ROLLBACK")
       LbAddOk = False
    End If
    
    Result = dosql("Close Scope")
     
    Return
    
End Sub

Private Sub Form_Load()

    OptSort(0).Value = True
    
End Sub


Private Sub OptSort_Click(Index As Integer)

    Call LIST_PROCESS
    
End Sub

Private Sub ssItemList_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    
    KeyAscii = 0
    
    SendKeys "{tab}"
    
End Sub


