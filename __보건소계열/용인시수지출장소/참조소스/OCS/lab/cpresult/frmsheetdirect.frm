VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSheetDirect 
   Caption         =   "InDrect, Direct Antiglobulin & Du Test"
   ClientHeight    =   7890
   ClientLeft      =   195
   ClientTop       =   750
   ClientWidth     =   11760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11760
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11025
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetDirect.frx":0000
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      ButtonWidth     =   1270
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Description     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Verify"
      Height          =   225
      Left            =   450
      TabIndex        =   2
      Top             =   2070
      Width           =   1410
   End
   Begin VB.OptionButton Option2 
      Caption         =   "NoVerify"
      Height          =   225
      Left            =   450
      TabIndex        =   1
      Top             =   2340
      Width           =   1410
   End
   Begin VB.OptionButton Option3 
      Caption         =   "ALL"
      Height          =   225
      Left            =   450
      TabIndex        =   0
      Top             =   2610
      Value           =   -1  'True
      Width           =   1410
   End
   Begin FPSpreadADO.fpSpread sprDirect 
      Height          =   6315
      Left            =   2070
      TabIndex        =   3
      Top             =   1170
      Width           =   9555
      _Version        =   196608
      _ExtentX        =   16854
      _ExtentY        =   11139
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      ScrollBars      =   2
      SpreadDesigner  =   "frmSheetDirect.frx":031C
      UserResize      =   1
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   450
      TabIndex        =   4
      Top             =   1530
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24510467
      CurrentDate     =   36566
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   450
      TabIndex        =   5
      Top             =   1170
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24510467
      CurrentDate     =   36566
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   510
      Left            =   405
      TabIndex        =   10
      Top             =   3735
      Width           =   1455
      Caption         =   "출력"
      PicturePosition =   327683
      Size            =   "2566;900"
      Picture         =   "frmSheetDirect.frx":1F1E
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   510
      Left            =   405
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2566;900"
      Picture         =   "frmSheetDirect.frx":27F8
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdChoise 
      Height          =   420
      Left            =   2385
      TabIndex        =   7
      Top             =   720
      Width           =   1770
      Caption         =   "전체선택"
      PicturePosition =   327683
      Size            =   "3122;741"
      Picture         =   "frmSheetDirect.frx":30DA
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "접수일자:"
      Height          =   240
      Left            =   135
      TabIndex        =   6
      Top             =   810
      Width           =   870
   End
End
Attribute VB_Name = "frmSheetDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChoise_Click()
    
    If cmdChoise.Caption = "전체선택" Then
        For i = 1 To sprDirect.DataRowCnt
            sprDirect.Row = i
            sprDirect.Col = 1
            sprDirect.Value = True
        Next
        cmdChoise.Caption = "전체해제"
    Else
        For i = 1 To sprDirect.DataRowCnt
            sprDirect.Row = i
            sprDirect.Col = 1
            sprDirect.Value = False
        Next
        cmdChoise.Caption = "전체선택"
    End If

End Sub

Private Sub cmdPrint_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    
    
    If sprDirect.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    sprDirect.RowHeight(-1) = 22
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "Item별 접수내역 LIST" & " - " & "(AntiGlobulin & Du TEST)"
    
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFrDate.Value, "yyyy-MM-dd") & " / " & _
                                               Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprDirect.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                            strFont(1) + "/n" + sBarLine + strFont(1)
    sprDirect.PrintFooter = "/f2" & "/l" & sBarLine & _
                            "/n" & Space(80) & "Page : " & "/p" & " of " & sprDirect.PrintPageCount
    sprDirect.PrintMarginLeft = 0
    sprDirect.PrintMarginRight = 0
    sprDirect.PrintMarginTop = 0
    sprDirect.PrintMarginBottom = 0
    sprDirect.PrintColHeaders = True
    sprDirect.PrintRowHeaders = True
    sprDirect.PrintBorder = False
    sprDirect.PrintColor = False
    sprDirect.PrintGrid = True
    sprDirect.PrintShadows = True
    sprDirect.PrintUseDataMax = False
    sprDirect.Row = 1
    sprDirect.Row2 = sprDirect.DataRowCnt
    sprDirect.Col = 2
    sprDirect.Col2 = sprDirect.MaxCols
    sprDirect.PrintOrientation = 1
    sprDirect.PrintOrientation = PrintOrientationPortrait
    sprDirect.PrintType = PrintTypeCellRange
    sprDirect.Action = ActionPrint
    
    sprDirect.RowHeight(-1) = 12
    

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate             As String
    Dim sToDate             As String

    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprDirect.ReDraw = False
    sprDirect.MaxRows = 0
    sprDirect.MaxRows = 300
    sprDirect.RowHeight(-1) = 11
    sprDirect.ReDraw = True
    
    For i = 1 To sprDirect.DataRowCnt
        sprDirect.Row = i
        sprDirect.Col = 1
        sprDirect.Value = False
    Next
    cmdChoise.Caption = "전체선택"
    
    
    GoSub Get_ABO_Data
    
    Exit Sub
    
    

Get_ABO_Data:
    strSql = ""
    strSql = strSql & " SELECT JeobsuDt1, Labno, Ptno,  Sname, Sx, Age, DeptCode, "
    strSql = strSql & "        MAX(Decode(ItemCode, '510105', Result11, '')) S1,"
    strSql = strSql & "        MAX(Decode(ItemCode, '510104', Result11, '')) S2,"
    strSql = strSql & "        MAX(Decode(ItemCode, '510127', Result11, '')) S3 "
    strSql = strSql & " FROM(  SELECT a.*, TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt1,c.DeptCode,"
    strSql = strSql & "               a.SLipno2 Labno, RTRIM(a.ItemCd) ItemCode,b.Sname, b.Sex sx, b.AgeYY age,"
    strSql = strSql & "               NVL(LTRIM(RTRIM(a.Result1)), '..') RESULT11"
    strSql = strSql & "        FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "               TWEXAM_IDNOMST     b,"
    strSql = strSql & "               TWEXAM_General     c"
    strSql = strSql & "        WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & "        AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    
    If Option1.Value = True Then
        strSql = strSql & "    AND    a.Verify    = 'Y'": End If
    If Option2.Value = True Then
        strSql = strSql & "    AND    a.Verify    = 'N'": End If
    
    strSql = strSql & "        AND    a.ITemCD   IN ('510105','510104','510127')"
    strSql = strSql & "        AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & "        AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "        AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "        AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "        AND    c.GBCh      = 'Y')"
    strSql = strSql & " GROUP BY JeobsuDt1, Labno, Ptno,  Sname, Sx, Age, DeptCode"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprDirect.Row = sprDirect.DataRowCnt + 1
        sprDirect.Col = 1:  sprDirect.Value = False
        sprDirect.Col = 2:  sprDirect.Text = adoSet.Fields("JeobsuDt1").Value & ""
        sprDirect.Col = 3:  sprDirect.Text = adoSet.Fields("Labno").Value & ""
        sprDirect.Col = 4:  sprDirect.Text = adoSet.Fields("Ptno").Value & ""
        sprDirect.Col = 5:  sprDirect.Text = Trim(adoSet.Fields("Sname").Value & "")
        sprDirect.Col = 6:  sprDirect.Text = adoSet.Fields("Sx").Value & ""
        sprDirect.Col = 7:  sprDirect.Text = adoSet.Fields("Age").Value & ""
        sprDirect.Col = 8:  sprDirect.Text = adoSet.Fields("DeptCode").Value & ""
        
        For i = 1 To 3
            sprDirect.Col = 8 + i
            Select Case Trim(adoSet.Fields("S" & i).Value & "")
                Case "..":    sprDirect.Text = ""
                Case "":      sprDirect.Text = "---------------"
                              sprDirect.Lock = True
                Case Else:    sprDirect.Text = Trim(adoSet.Fields("S" & i).Value & "")
            End Select
        Next
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Call SpreadSetClear(sprDirect)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
