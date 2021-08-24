VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmExResult 
   Caption         =   "외부의뢰검사 결과입력"
   ClientHeight    =   7785
   ClientLeft      =   660
   ClientTop       =   2655
   ClientWidth     =   11505
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
   ScaleHeight     =   7785
   ScaleWidth      =   11505
   WindowState     =   2  '최대화
   Begin Threed.SSCommand cmdSelect 
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   675
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "▼ 모두선택"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin VB.TextBox txtRemark 
      BackColor       =   &H00800000&
      ForeColor       =   &H80000005&
      Height          =   465
      Left            =   5985
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   7
      Top             =   6840
      Width           =   5775
   End
   Begin FPSpreadADO.fpSpread sprExResult 
      Height          =   5730
      Left            =   90
      TabIndex        =   4
      Top             =   1080
      Width           =   11715
      _Version        =   196608
      _ExtentX        =   20664
      _ExtentY        =   10107
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   11
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
      MaxCols         =   15
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmExResult.frx":0000
      Appearance      =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3015
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
            Picture         =   "frmExResult.frx":0F1E
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   635
      ButtonWidth     =   1270
      ButtonHeight    =   582
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
            Description     =   "Exit of Screen"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   6570
      TabIndex        =   1
      Top             =   630
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36528
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   5085
      TabIndex        =   0
      Top             =   630
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36528
   End
   Begin MSForms.CommandButton cmdSet 
      Height          =   465
      Left            =   4500
      TabIndex        =   8
      Top             =   6840
      Width           =   1455
      Caption         =   "Setting"
      PicturePosition =   327683
      Size            =   "2566;820"
      Picture         =   "frmExResult.frx":123A
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsert 
      Height          =   465
      Left            =   9540
      TabIndex        =   6
      Top             =   585
      Width           =   1545
      Caption         =   "결과입력"
      PicturePosition =   327683
      Size            =   "2725;820"
      Picture         =   "frmExResult.frx":1554
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   465
      Left            =   8145
      TabIndex        =   5
      Top             =   585
      Width           =   1410
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2487;820"
      Picture         =   "frmExResult.frx":2D16
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "접수일자:"
      Height          =   240
      Left            =   4050
      TabIndex        =   3
      Top             =   675
      Width           =   960
   End
End
Attribute VB_Name = "frmExResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInsert_Click()
    Dim iRowCount       As Integer
    Dim sRowID          As String
    Dim sJeobsuDt       As String
    Dim sSLipno1        As String
    Dim sSLipno2        As String
    Dim sPtno           As String
    Dim sResult1        As String
    Dim iExeCount       As Integer
    Dim sGeomsaCM       As String
    
    
    
    
    GoSub Check_InsertRow
    GoSub Main_Process
    Exit Sub
    
    
Check_InsertRow:
    iRowCount = 0
    For i = 1 To Me.sprExResult.DataRowCnt
        sprExResult.Row = i
        sprExResult.Col = 1
        If sprExResult.Value = True Then
            iRowCount = iRowCount + 1
        End If
    Next
    
    If iRowCount = 0 Then
        MsgBox "결과입력이 Check된 행이 없습니다!..........."
        Exit Sub
    End If
    Return
    

Main_Process:
    iExeCount = 0
    For i = 1 To Me.sprExResult.DataRowCnt
        sprExResult.Row = i
        sprExResult.Col = 1
        If sprExResult.Value = True Then
            sprExResult.Col = 2:  sRowID = sprExResult.Text
            sprExResult.Col = 3:  sJeobsuDt = sprExResult.Text
            sprExResult.Col = 4:  sSLipno1 = sprExResult.Text
            sprExResult.Col = 5:  sSLipno2 = sprExResult.Text
            sprExResult.Col = 6:  sPtno = sprExResult.Text
            sprExResult.Col = 13: sResult1 = sprExResult.Text
            sprExResult.Col = 15: sGeomsaCM = Trim(sprExResult.Text)
            GoSub Update_Result_General_Sub
            GoSub Update_Result_General
            iExeCount = iExeCount + 1
        End If
    Next
    
    MsgBox iExeCount & " 개 Data 의 결과를 입력하였습니다!......."
    
    Return
    
Update_Result_General_Sub:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General_Sub"
    strSql = strSql & " SET    Result1 = '" & Quot_Conv(sResult1) & "',"
    strSql = strSql & "        Verify  = 'Y'"
    strSql = strSql & " WHERE  RowId   = '" & sRowID & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    
    
Update_Result_General:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    Status   = 'C',"         '검사완료
    strSql = strSql & "        GeomsaCm = '" & Quot_Conv(sGeomsaCM) & "'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    SLipno2  =  " & Val(sSLipno2)
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    
    
End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    
    
    
    GoSub Set_Initial
    GoSub Main_Query
    For i = 1 To Me.sprExResult.DataRowCnt
        Me.sprExResult.Row = i
        Me.sprExResult.Col = 12
        GoSub RET_Setting_Init1
    Next
    
    Exit Sub
    
    


Set_Initial:
    Call SpreadSetClear(sprExResult)
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    Return
    
    
Main_Query:
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
    strSql = strSql & "        g.Result1, g.Result4, g.Result5, g.Verify, h.GeomchCd,"
    strSql = strSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, i.OldCode, p.Sname,"
    strSql = strSql & "        g.RowID SubRowID, h.RowID GeneralRowID, "
    strSql = strSql & "        p.Sex, p.AgeYY, h.GeomsaCM"
    strSql = strSql & "  FROM  TWEXAM_General_Sub g,"
    strSql = strSql & "        TWEXAM_General     h,"
    strSql = strSql & "        TWEXAM_ItemML      i,"
    strSql = strSql & "        TWEXAM_IDNOMST     p"
    strSql = strSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    strSql = strSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    'strSql = strSql & "  AND   g.Verify    = 'N'"
    'strSql = strSql & "  AND   i.GeomsaGb  = 'W'"
    strSql = strSql & "  AND    g.Codegu    = 'W'"
    strSql = strSql & "  AND   g.ItemCd    = I.CodeKy"
    'strSql = strSql & "  AND   g.Result4   = '1'"         'Print(Sheet) 발행 Check
    strSql = strSql & "  AND   g.JeobsuDt  = h.JeobsuDt(+)"
    strSql = strSql & "  AND   g.SLipno1   = h.SLipno1(+)"
    strSql = strSql & "  AND   g.SLipno2   = h.SLipno2(+)"
    strSql = strSql & "  AND   g.Ptno      = p.Ptno(+)"
    strSql = strSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprExResult.Row = sprExResult.DataRowCnt + 1
        sprExResult.Col = 1:  sprExResult.Value = False
        sprExResult.Col = 2:  sprExResult.Text = adoSet.Fields("SubRowID").Value & ""
        sprExResult.Col = 3:  sprExResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprExResult.Col = 4:  sprExResult.Text = adoSet.Fields("SLipno1").Value & ""
        sprExResult.Col = 5:  sprExResult.Text = adoSet.Fields("SLipno2").Value & ""
        sprExResult.Col = 6:  sprExResult.Text = adoSet.Fields("Ptno").Value & ""
        sprExResult.Col = 7:  sprExResult.Text = adoSet.Fields("Sname").Value & ""
        sprExResult.Col = 8:  sprExResult.Text = adoSet.Fields("Sex").Value & ""
        sprExResult.Col = 9:  sprExResult.Text = adoSet.Fields("ageyy").Value & ""
        sprExResult.Col = 10: sprExResult.Text = adoSet.Fields("ItemCd").Value & ""
        sprExResult.Col = 11: sprExResult.Text = adoSet.Fields("OldCode").Value & ""
        sprExResult.Col = 12: sprExResult.Text = adoSet.Fields("ItemNM").Value & ""
        sprExResult.Col = 13: sprExResult.Text = adoSet.Fields("Result1").Value & ""
        
        sprExResult.Col = 15: sprExResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        
        If Trim(sprExResult.Text) = "" Then
            sprExResult.Col = 14
            sprExResult.CellType = CellTypeStaticText
            sprExResult.Col = 14: sprExResult.Text = ""
        Else
            sprExResult.Col = 14
            sprExResult.CellType = CellTypeButton
            sprExResult.TypeButtonText = "R"
        End If
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    

RET_Setting_Init1:
    Dim sResult     As String
    Dim sItemCd     As String
    
    
    sprExResult.Col = 13:
    sprExResult.CellType = SS_CELL_TYPE_EDIT
    sprExResult.TypeHAlign = SS_CELL_H_ALIGN_LEFT
    sprExResult.TypeEditCharCase = SS_CELL_EDIT_CASE_NO_CASE
    sprExResult.TypeEditMultiLine = False
    sprExResult.TypeEditLen = 50

    sprExResult.Col = 10: sItemCd = sprExResult.Text
    sResult = Get_Result_Text(sItemCd)
    If Trim(sResult) <> "" Then
        sprExResult.Col = 13
        sprExResult.CellType = CellTypeComboBox
        sprExResult.TypeComboBoxList = sResult
        sprExResult.TypeComboBoxEditable = True
    End If
    
    Return

End Sub

Private Sub cmdSelect_Click()
    If cmdSelect.Caption = "▼ 모두선택" Then
        For i = 1 To Me.sprExResult.DataRowCnt
            sprExResult.Row = i
            sprExResult.Col = 1
            sprExResult.Value = True
            cmdSelect.Caption = "▼ 모두해제"
        Next
    Else
        For i = 1 To Me.sprExResult.DataRowCnt
            sprExResult.Row = i
            sprExResult.Col = 1
            sprExResult.Value = False
            cmdSelect.Caption = "▼ 모두선택"
        Next
    End If
    
End Sub

Private Sub cmdSet_Click()
    
    sprExResult.Row = sprExResult.ActiveRow
    sprExResult.Col = 15
    sprExResult.Text = txtRemark.Text
    
    
    sprExResult.Col = 14
    If Trim(txtRemark.Text) = "" Then
        sprExResult.CellType = CellTypeStaticText
        sprExResult.Text = ""
    Else
        sprExResult.CellType = CellTypeButton
        sprExResult.TypeButtonText = "R"
    End If
    
    sprExResult.Row = sprExResult.ActiveRow
    sprExResult.Col = 1
    sprExResult.Value = True
    txtRemark.Text = ""
    
End Sub

Private Sub Form_Load()
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub

Private Sub sprExResult_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Col = 14 Then
        sprExResult.Row = Row
        sprExResult.Col = 15
        txtRemark.Text = sprExResult.Text
        sprExResult.Action = ActionActiveCell
        'sprExResult.ActiveRow = Row
    End If
    
    
End Sub

Private Sub sprExResult_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col = 14 Then
        sprExResult.Row = Row
        sprExResult.Col = Col
        If sprExResult.CellType <> CellTypeButton Then
            txtRemark.Text = ""
        End If
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
