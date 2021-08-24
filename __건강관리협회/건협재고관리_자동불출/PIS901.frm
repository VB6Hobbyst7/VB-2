VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form PIS901 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검체RACK정보"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   14880
   WindowState     =   2  '최대화
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   9705
      Left            =   30
      Top             =   30
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   17119
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TextPosition    =   0
      TextCustomMargin=   4
      GroupBoxStyle   =   0
      TextBarColor1   =   12757903
      TextBarStyle    =   3
      TextBarColor2   =   11767328
      TextBarSymbol   =   0   'False
      TextBarSymbolColor=   16777215
      TextBarHeightMargin=   10
      MouseCursor     =   0
      TextBarMouseCursor=   0
      IconandTextMargin=   4
      BodyColor       =   16777215
      Enabled         =   -1  'True
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   11070
         TabIndex        =   3
         Top             =   9240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "저 장"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS901.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDelete 
         Height          =   375
         Left            =   12300
         TabIndex        =   2
         Top             =   9240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "삭 제"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS901.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   1
         Top             =   9240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "닫 기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS901.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   9075
         Left            =   60
         TabIndex        =   0
         Tag             =   "20001"
         Top             =   60
         Width           =   14670
         _Version        =   524288
         _ExtentX        =   25876
         _ExtentY        =   16007
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   12
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS901.frx":4746
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psDataDisplay()
Dim sRow As Integer

    If gWorkArea Then
        gSql = "SELECT A.RACKNO,A.ROWCNT,A.COLCNT,A.USEFG,B.SAVEDT,B.EXPIRYDT,B.DEPOTCD,C.DEPOTNM,B.SAVEFLOOR,B.EMPID,D.EMPNM,B.WRTDT           " & vbNewLine & _
                "    , (SELECT COUNT(*) FROM S2PIS902 Z WHERE A.RACKNO=Z.RACKNO AND Z.STATUS='0') AS SPCCNT                                     " & vbNewLine & _
                "  FROM S2PIS091 A LEFT JOIN S2PIS901 B ON A.RACKNO=B.RACKNO                                                                    " & vbNewLine & _
                "                  LEFT JOIN S2PIS092 C ON B.DEPOTCD=C.DEPOTCD                                                                  " & vbNewLine & _
                "                  LEFT JOIN S2COM006 D ON B.EMPID=D.EMPID                                                                      " & vbNewLine & _
                " ORDER BY A.RACKNO"
    Else
        gSql = "SELECT A.RACKNO,A.ROWCNT,A.COLCNT,A.USEFG,B.SAVEDT,B.EXPIRYDT,B.DEPOTCD,C.DEPOTNM,B.SAVEFLOOR,B.EMPID,D.USER_NM AS EMPNM,B.WRTDT" & vbNewLine & _
                "    , (SELECT COUNT(*) FROM S2PIS902 Z WHERE A.RACKNO=Z.RACKNO AND Z.STATUS='0') AS SPCCNT                                     " & vbNewLine & _
                "  FROM S2PIS091 A LEFT JOIN S2PIS901 B ON A.RACKNO=B.RACKNO                                                                    " & vbNewLine & _
                "                  LEFT JOIN S2PIS092 C ON B.DEPOTCD=C.DEPOTCD                                                                  " & vbNewLine & _
                "                  LEFT JOIN " & gKahpUserTable & " D ON B.EMPID=D.USERID                                                       " & vbNewLine & _
                " ORDER BY A.RACKNO"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            spList.Row = -1
            spList.Col = 2
            spList.TypeMaxEditLen = .Fields("RACKNO").DefinedSize
            
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 100, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("RACKNO").Value
                    spList.SetText 3, sRow, Val("" & .Fields("ROWCNT").Value)
                    spList.SetText 4, sRow, Val("" & .Fields("COLCNT").Value)
                    spList.SetText 5, sRow, "" & .Fields("USEFG").Value
                    spList.SetText 6, sRow, Format("" & .Fields("SAVEDT").Value, "####-##-##")
                    spList.SetText 7, sRow, Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                    If Val("" & .Fields("SPCCNT").Value) > 0 Then
                        spList.SetText 8, sRow, "" & .Fields("SPCCNT").Value
                    End If
                    spList.SetText 9, sRow, "" & .Fields("DEPOTNM").Value
                    spList.SetText 10, sRow, "" & .Fields("SAVEFLOOR").Value
                    spList.SetText 11, sRow, "" & .Fields("EMPNM").Value
                    spList.SetText 12, sRow, "" & .Fields("WRTDT").Value
                    
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With

End Sub

Private Sub psDataProcess(ByVal brJob As Boolean)
Dim cPis091 As clsPis091, sRow As Long, sGetVal As Variant, sReturn As Boolean, sSaveDt As String

    MousePointer = vbHourglass
    Set cPis091 = New clsPis091
    
    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sGetVal:      cPis091.rackno = Trim(sGetVal)
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(cPis091.rackno) > 0 Then
                .GetText 3, sRow, sGetVal:      cPis091.rowcnt = Trim(sGetVal)
                .GetText 4, sRow, sGetVal:      cPis091.colcnt = Trim(sGetVal)
                .GetText 5, sRow, sGetVal:      cPis091.usefg = Trim(sGetVal)
                cPis091.empid = gUserId
                
                If brJob Then
                    sReturn = cPis091.cfSave
                Else
                    .GetText 6, sRow, sGetVal:      sSaveDt = Trim(sGetVal)
                    If Len(sSaveDt) > 0 Then
                        MsgBox "보관처리된 RACK을 삭제 할 수 없습니다.!", vbCritical
                        sReturn = False
                        Exit For
                    Else
                        sReturn = cPis091.cfDelete
                    End If
                End If
                
                If sReturn Then
                    .SetText 1, sRow, ""
                Else
                    Exit For
                End If
            End If
        Next sRow
    End With
    
    If sReturn Then Call psDataDisplay
    MousePointer = vbDefault

End Sub

Private Sub cmdClose_Click()

     Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Call psDataProcess(False)
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdSave_Click()

    MousePointer = vbHourglass
    Call psDataProcess(True)
    MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    With spList
        .Row = SpreadHeader
        .Col = 2
        .Text = "▒ RACK Master ▒"
        .FontBold = True
        .Col = 6
        .Text = "▒ RACK 저장정보 ▒"
        .FontBold = True
    End With
    
    Call psDataDisplay

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Top = (Me.ScaleHeight - grpMain.Height) / 2
    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sGetVal As Variant, sUseFg As Integer, sSaveDt As String

    With spList
        If Row > 0 And Col > 1 Then
            .SetText 1, Row, 1
            
            If Col = 5 Then
                .GetText 5, Row, sGetVal:       sUseFg = Val(sGetVal)
                .GetText 6, Row, sGetVal:       sSaveDt = Trim(sGetVal)
                If sUseFg > 0 And Len(sSaveDt) > 0 Then
                    MsgBox "보관처리된 RACK을 사용중지할 수 없습니다.!", vbCritical
                    .SetText 5, Row, ""
                End If
            End If
        End If
    End With

End Sub
