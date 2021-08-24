VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmExadd 
   Caption         =   "외부코드 등록"
   ClientHeight    =   7890
   ClientLeft      =   1005
   ClientTop       =   1425
   ClientWidth     =   11565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11565
   WindowState     =   2  '최대화
   Begin Threed.SSCommand cmdInsert 
      Height          =   1410
      Left            =   8145
      TabIndex        =   7
      Top             =   1305
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   2487
      _StockProps     =   78
      Caption         =   "입력확인"
      BevelWidth      =   1
      Outline         =   0   'False
      Picture         =   "frmExadd.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "조회조건"
      Height          =   825
      Left            =   540
      TabIndex        =   1
      Top             =   270
      Width           =   8745
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   1080
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   360
         Width           =   2625
      End
      Begin VB.OptionButton Option1 
         Caption         =   "전체List"
         Height          =   285
         Left            =   4050
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton Option2 
         Caption         =   "외부검사"
         Height          =   285
         Left            =   5265
         TabIndex        =   2
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "검사종목"
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   405
         Width           =   825
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   420
         Left            =   6750
         TabIndex        =   5
         Top             =   270
         Width           =   1635
         Caption         =   "조회확인"
         Size            =   "2884;741"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread spriTemLst 
      Height          =   6135
      Left            =   585
      TabIndex        =   0
      Top             =   1305
      Width           =   7485
      _Version        =   196608
      _ExtentX        =   13203
      _ExtentY        =   10821
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
      MaxCols         =   6
      ScrollBars      =   2
      SpreadDesigner  =   "frmExadd.frx":17C2
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   1410
      Left            =   8235
      TabIndex        =   8
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmExadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdInsert_Click()
    Dim sRowid          As String
    Dim sGeomsaGb       As String
    Dim sOLdCode        As String
    Dim sCodeGu         As String
    
    
    For i = 1 To Me.spriTemLst.DataRowCnt
        spriTemLst.Row = i
        spriTemLst.Col = 1: sRowid = spriTemLst.Text
        spriTemLst.Col = 4: sGeomsaGb = spriTemLst.Text
        spriTemLst.Col = 5: sOLdCode = spriTemLst.Text
        spriTemLst.Col = 6: sCodeGu = spriTemLst.Text
        
        strSql = ""
        strSql = strSql & " UPDATE TWEXAM_ITEMML"
        strSql = strSql & " SET    GeomsaGb = '" & sGeomsaGb & "',"
        strSql = strSql & "        OLdCode  = '" & sOLdCode & "',"
        strSql = strSql & "        Codegu   = '" & sCodeGu & "'"
        strSql = strSql & " WHERE  ROWID    = '" & sRowid & "'"
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
    Next
    MsgBox "입력 완료 되었습니다!.."
    
    
    
End Sub

Private Sub cmdQuery_Click()
    
        
    Call Spread_Set_Clear(spriTemLst)
    GoSub get_ItemList
    Exit Sub
    
get_ItemList:
    strSql = ""
    strSql = strSql & " SELECT a.RowID RWID, a.Codeky, a.Itemnm, a.GeomsaGB, "
    strSql = strSql & "        a.OLdCode, a.Codegu"
    strSql = strSql & " FROM   TWEXAM_ITEMML a"
    strSql = strSql & " WHERE  CODEKY Like '" & Left(cmbSLip.Text, 2) & "%'"
    
    If Option2.Value = True Then
        strSql = strSql & " AND  a.GeomsaGb = 'W'"       '외부검사만......(J: 자체검사  W:외부검사)
    End If
    
    strSql = strSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    spriTemLst.MaxRows = 0
    spriTemLst.MaxRows = adoSet.RecordCount
    spriTemLst.RowHeight(-1) = 10.5
    Do Until adoSet.EOF
        spriTemLst.Row = spriTemLst.DataRowCnt + 1
        spriTemLst.Col = 1: spriTemLst.Text = adoSet.Fields("RWID").Value & ""
        spriTemLst.Col = 2: spriTemLst.Text = adoSet.Fields("Codeky").Value & ""
        spriTemLst.Col = 3: spriTemLst.Text = adoSet.Fields("ItemNM").Value & ""
        spriTemLst.Col = 4: spriTemLst.Text = adoSet.Fields("GeomsaGB").Value & ""
        spriTemLst.Col = 5: spriTemLst.Text = adoSet.Fields("OldCode").Value & ""
        spriTemLst.Col = 6: spriTemLst.Text = adoSet.Fields("Codegu").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Return

End Sub

Private Sub Form_Load()
    
    
    GoSub Set_Specode12
    
    Exit Sub
    
Set_Specode12:
    cmbSLip.Clear
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '52'"
    strSql = strSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                             adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
