VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmNormalEdit 
   Caption         =   "참조치 Data 관리"
   ClientHeight    =   3750
   ClientLeft      =   2370
   ClientTop       =   2565
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7215
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   60
      Width           =   6795
      _Version        =   65536
      _ExtentX        =   11986
      _ExtentY        =   1085
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtItemCD 
         Enabled         =   0   'False
         Height          =   330
         Left            =   225
         TabIndex        =   5
         Top             =   75
         Width           =   1020
      End
      Begin VB.TextBox txtItemName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1245
         TabIndex        =   4
         Top             =   75
         Width           =   2565
      End
      Begin MSForms.CommandButton cmdQryOk 
         Height          =   465
         Left            =   3855
         TabIndex        =   6
         Top             =   75
         Width           =   1470
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2593;820"
         Picture         =   "frmNormalEdit.frx":0000
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread ssRefData 
      Height          =   2475
      Left            =   180
      TabIndex        =   0
      Top             =   1125
      Width           =   6900
      _Version        =   196608
      _ExtentX        =   12171
      _ExtentY        =   4366
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
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
      MaxCols         =   11
      NoBorder        =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmNormalEdit.frx":08DA
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "연령구분이 없을시에는 0세 ~ 999세로 Setting"
      Height          =   240
      Left            =   225
      TabIndex        =   7
      Top             =   810
      Width           =   3795
   End
   Begin MSForms.CommandButton cmdRef 
      Height          =   375
      Left            =   5565
      TabIndex        =   2
      Top             =   720
      Width           =   1425
      Caption         =   "Refrash"
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmNormalEdit.frx":47D1
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsert 
      Height          =   375
      Left            =   4095
      TabIndex        =   1
      Top             =   720
      Width           =   1470
      Caption         =   "입력확인"
      PicturePosition =   327683
      Size            =   "2593;661"
      Picture         =   "frmNormalEdit.frx":5623
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmNormalEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInsert_Click()
    Dim sAppdate        As String
    Dim sAppGubun       As String
    Dim sItemCD         As String
    Dim nAgeMin         As Integer
    Dim nAgeMax         As Integer
    Dim sMmin           As String
    Dim sMmax           As String
    Dim sFmin           As String
    Dim sFmax           As String
    Dim sRowid          As String
    Dim bDelCheck       As Boolean
    
    For I = 1 To ssRefData.DataRowCnt
        ssRefData.Row = I
        GoSub SpreadData_Move
        If bDelCheck = True Then              'Col =1: Del Check
            If Trim(sRowid) <> "" Then        'Col =2: RowID Check
                GoSub RefData_Delete
            End If
        Else
            If Trim(sRowid) = "" Then
                GoSub RefData_Insert
            Else
                GoSub RefData_Update
            End If
        End If
    Next
    Call cmdRef_Click
    Exit Sub
    
    
RefData_Delete:
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_RefData"
    strSql = strSql & " WHERE  ROWID  =  '" & sRowid & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    

SpreadData_Move:
    ssRefData.Row = I
    ssRefData.Col = 1
    If ssRefData.Value = True Then
        bDelCheck = True
    Else
        bDelCheck = False
    End If
    ssRefData.Col = 2:  sRowid = Trim(ssRefData.Text)
    ssRefData.Col = 3:  sAppdate = ssRefData.Text
    ssRefData.Col = 4:  sItemCD = Trim(ssRefData.Text)
    ssRefData.Col = 5:  sAppGubun = Trim(ssRefData.Text)
    ssRefData.Col = 6:  nAgeMin = Val(ssRefData.Text)
    ssRefData.Col = 7:  nAgeMax = Val(ssRefData.Text)
    ssRefData.Col = 8:  sMmin = ssRefData.Text
    ssRefData.Col = 9:  sMmax = ssRefData.Text
    ssRefData.Col = 10: sFmin = ssRefData.Text
    ssRefData.Col = 11: sFmax = ssRefData.Text
    Return
    
'/___________________________________________________________________
RefData_Insert:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_RefData"
    strSql = strSql & "        (iTemCode, appDate, appGubun, ageMin, ageMax, "
    strSql = strSql & "         M_min,    M_max,   F_min,    F_max)"
    strSql = strSql & " VALUES ('" & txtItemCD.Text & "',"
    strSql = strSql & "              TO_DATE('" & sAppdate & "','YYYY-MM-DD'),"
    strSql = strSql & "         '" & sAppGubun & "',"
    strSql = strSql & "          " & nAgeMin & ","
    strSql = strSql & "          " & nAgeMax & ","
    strSql = strSql & "         '" & sMmin & "',"
    strSql = strSql & "         '" & sMmax & "',"
    strSql = strSql & "         '" & sFmin & "',"
    strSql = strSql & "         '" & sFmax & "')"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    
RefData_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_RefData"
    strSql = strSql & " SET    appDate  = TO_DATE('" & sAppdate & "','yyyy-MM-dd'),"
    strSql = strSql & "        appGubun = '" & sAppGubun & "',"
    strSql = strSql & "        ageMin   =  " & nAgeMin & ","
    strSql = strSql & "        ageMax   =  " & nAgeMax & ","
    strSql = strSql & "        M_min    = '" & sMmin & "',"
    strSql = strSql & "        M_max    = '" & sMmax & "',"
    strSql = strSql & "        F_min    = '" & sFmin & "',"
    strSql = strSql & "        F_max    = '" & sFmax & "'"
    strSql = strSql & " WHERE  RowID    = '" & sRowid & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

End Sub

Private Sub cmdQryOk_Click()

    strSql = " SELECT ItemNm FROM  TWEXAM_iTemML WHERE Codeky = '" & txtItemCD.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "해당 코드가 없습니다!"
        txtItemCD.Text = ""
        txtItemName.Text = ""
        Exit Sub
    End If
    txtItemName.Text = adoSet.Fields("iTemNM").Value & ""
    Call adoSetClose(adoSet)
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & Me.txtItemCD.Text & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC, a.AppGubun ASC, a.AgeMin ASC"
    
       
    Call Spread_Set_Clear(ssRefData)
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 11: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdRef_Click()
    
    Call cmdQryOk_Click
    
End Sub

Private Sub Form_Load()
    
    ssRefData.RowHeight(-1) = 11
    
    txtItemCD.Text = Trim(frmItemCode.txtSlipno.Text) & Trim(frmItemCode.txtItemCode.Text)
    Call cmdQryOk_Click
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssRefData_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    
    If Col = 6 Then
        ssRefData.Row = Row
        ssRefData.Col = 2
        If Trim(ssRefData.Text) = "" Then Exit Sub
        ssRefData.Col = Col
        If Trim(ssRefData.Text) = "" Then
            ssRefData.Text = "0"
        End If
    ElseIf Col = 7 Then
        ssRefData.Row = Row
        ssRefData.Col = 2
        If Trim(ssRefData.Text) = "" Then Exit Sub
        ssRefData.Col = Col
        If Trim(ssRefData.Text) = "" Then
            ssRefData.Text = "999"
        End If
    End If
    
End Sub
