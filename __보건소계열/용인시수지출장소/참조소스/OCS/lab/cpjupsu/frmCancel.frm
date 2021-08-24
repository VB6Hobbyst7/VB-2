VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCancel 
   Caption         =   "접수취소화면"
   ClientHeight    =   7575
   ClientLeft      =   450
   ClientTop       =   570
   ClientWidth     =   11055
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
   ScaleHeight     =   7575
   ScaleWidth      =   11055
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Left            =   810
      TabIndex        =   6
      Top             =   945
      Width           =   10140
      _Version        =   65536
      _ExtentX        =   17886
      _ExtentY        =   979
      _StockProps     =   15
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
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtPtno 
         Height          =   330
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   0
         Top             =   135
         Width           =   1230
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox txtSexage 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   330
         Left            =   1260
         TabIndex        =   9
         Top             =   135
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24707075
         CurrentDate     =   36537
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   465
         Left            =   7290
         TabIndex        =   1
         Top             =   45
         Width           =   1590
         Caption         =   "Order조회"
         PicturePosition =   327683
         Size            =   "2805;820"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label labPtno 
         Caption         =   "등록번호:"
         Height          =   240
         Left            =   3060
         TabIndex        =   13
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자:"
         Height          =   240
         Left            =   225
         TabIndex        =   12
         Top             =   180
         Width           =   915
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   810
      TabIndex        =   5
      Top             =   495
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   "접수 취소작업 화면"
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
      BorderWidth     =   1
      BevelInner      =   1
   End
   Begin FPSpreadADO.fpSpread sprDt 
      Height          =   5460
      Left            =   6480
      TabIndex        =   4
      Top             =   1575
      Width           =   4470
      _Version        =   196608
      _ExtentX        =   7885
      _ExtentY        =   9631
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
      MaxCols         =   3
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmCancel.frx":0000
      UserResize      =   1
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread sprOrder 
      Height          =   4695
      Left            =   810
      TabIndex        =   3
      Top             =   1575
      Width           =   5640
      _Version        =   196608
      _ExtentX        =   9948
      _ExtentY        =   8281
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   6
      ScrollBars      =   2
      SpreadDesigner  =   "frmCancel.frx":0B4E
      UserResize      =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11160
      Top             =   495
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
            Picture         =   "frmCancel.frx":2470
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
      Width           =   11055
      _ExtentX        =   19500
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
            Description     =   "Exit of CancelScreen"
            Object.ToolTipText     =   "Exit of CancelScreen"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   510
      Left            =   4725
      TabIndex        =   8
      Top             =   6525
      Width           =   1590
      Caption         =   "Clear"
      PicturePosition =   327683
      Size            =   "2805;900"
      Picture         =   "frmCancel.frx":278C
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   510
      Left            =   3105
      TabIndex        =   7
      Top             =   6525
      Width           =   1635
      Caption         =   "취소확인"
      PicturePosition =   327683
      Size            =   "2884;900"
      Picture         =   "frmCancel.frx":3F1E
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function IS_NOTResultData(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer) As Integer
    Dim adoIS       As ADODB.Recordset
    Dim strIS       As String
    
        
    strIS = ""
    strIS = strIS & " SELECT *"
    strIS = strIS & " FROM   TWEXAM_GENERAL_SUB"
    strIS = strIS & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strIS = strIS & " AND    SLipno1  = " & iSLno1
    strIS = strIS & " AND    SLipno2  = " & iSLno2
    strIS = strIS & " AND    Verify   = 'Y'"
    
    If False = adoSetOpen(strIS, adoIS) Then
        IS_NOTResultData = False
    Else
        IS_NOTResultData = True
        Call adoSetClose(adoIS)
    End If
    
End Function

Private Sub CmdCancel_Click()
    Dim sJeobsuDt           As String
    Dim sSLipno1            As String
    Dim sSLipno2            As String
    Dim sOrderno            As String
    Dim iCheckCount         As Integer
    Dim iJobcount           As Integer
    Dim sTableName          As String
    Dim sStat               As String * 1
    
    sTableName = ""
    iJobcount = 0
    
    
    For i = 1 To sprOrder.DataRowCnt
        sprOrder.Row = i
        sprOrder.Col = 1
        If sprOrder.Value = True Then iCheckCount = iCheckCount + 1
    Next
    
    If iCheckCount = 0 Then
        MsgBox "접수 취소할 선택된 Data가 없습니다!..."
        Exit Sub
    End If
    
    
    For i = 1 To sprOrder.DataRowCnt
        sprOrder.Row = i
        sprOrder.Col = 2: sJeobsuDt = sprOrder.Text
        sprOrder.Col = 3: sSLipno1 = sprOrder.Text
        sprOrder.Col = 4: sSLipno2 = sprOrder.Text
        sprOrder.Col = 6: sOrderno = sprOrder.Text
        
        sprOrder.Col = 1
        If sprOrder.Value = True Then
            sStat = Trim(Get_Status(sJeobsuDt, Val(sSLipno1), Val(sSLipno2)))
            If Trim(sStat) = "R" Or Trim(sStat) = "" Then
                iCheckCount = 0
                GoSub GENERAL_SUB_DATA_DELETE
                GoSub GENERAL_DATA_DELETE
                GoSub EXAM_Order_Update
            Else
                MsgBox "결과가 입력되어 있습니다!!!. 검사Part 에 확인하세요" & _
                       "검사종목은? " & sSLipno1
                Exit Sub
            End If
        End If
    Next
    
    If iJobcount > 2 Then
        MsgBox "작업을 끝마쳤습니다!..........."
    Else
        MsgBox "어떠한 오류로 인하여 작업을 완료하지 못하였습니다!.." & vbCrLf & _
               "등록번호 : " & txtPtno.Text & "의 Data 를 전산실에 문의하세요" & vbCrLf & _
               sTableName, vbOKOnly, iCheckCount
        Exit Sub
    End If
    
    
    Exit Sub
    


GENERAL_SUB_DATA_DELETE:
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    SLipno2  = " & Val(sSLipno2)
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        iJobcount = iJobcount + 1
        sTableName = "TWEXAM_General_Sub"
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
GENERAL_DATA_DELETE:
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TWEXAM_GENERAL"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    SLipno2  = " & Val(sSLipno2)
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        iJobcount = iJobcount + 1
        sTableName = sTableName & vbCrLf & "TWEXAM_General"
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    

EXAM_Order_Update:
    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_Order"
    strSql = strSql & " SET   COLLDATE   = NULL,"
    strSql = strSql & "       COLLHH     = NULL,"
    strSql = strSql & "       COLLMM     = NULL,"
    strSql = strSql & "       COLLID     = NULL,"
    strSql = strSql & "       JEOBSU_LAB = NULL,"
    strSql = strSql & "       JEOBSUYN   = NULL,"
    strSql = strSql & "       GeomsaGu   = 'R', "
    strSql = strSql & "       GBCH       = NULL,"
    strSql = strSql & "       GBDate     = NULL,"
    strSql = strSql & "       Matchno    = NULL"
    strSql = strSql & " WHERE CoLLDate   = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND   SLipno1    = " & Val(sSLipno1)
    strSql = strSql & " AND   Ptno       = '" & txtPtno.Text & "'"
    strSql = strSql & " AND   Orderno    = " & Val(sOrderno)
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        iJobcount = iJobcount + 1
        sTableName = sTableName & vbCrLf & "TW_MIS_EXAM.TWEXAM_Order"
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    

End Sub

Private Sub cmdClear_Click()

    dtDate.Value = Dual_Date_Get("yyyy-MM-dd")
    txtPtno.Text = ""
    txtSname.Text = ""
    txtSexage.Text = ""
    Call Spread_Set_Clear(sprOrder)
    Call Spread_Set_Clear(sprDt)
    
    
    
End Sub

Private Sub cmdQuery_Click()
    Dim sToDate         As String
    
    If Trim(txtPtno.Text) = "" Then
        MsgBox "접수취소할 등록번호를 입력하세요!."
        Exit Sub
    End If
    
    Call Spread_Set_Clear(sprOrder)
    Call Spread_Set_Clear(sprDt)
    
    
    
    sToDate = Format(dtDate.Value, "yyyy-MM-dd")
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        a.SLipno1, a.SLipno2, b.Codenm, a.Orderno"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode     b"
    strSql = strSql & " WHERE  a.JeobsuDt =  TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.PTNO     = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    b.codeky   = a.Slipno1"
    strSql = strSql & " AND    b.Codegu   = '12'"
    strSql = strSql & " GROUP  BY JeobsuDt, SLipno1, SLipno2, b.Codenm, a.Orderno"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1: sprOrder.Value = False
        sprOrder.Col = 2: sprOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprOrder.Col = 3: sprOrder.Text = adoSet.Fields("SLipno1").Value & ""
        sprOrder.Col = 4: sprOrder.Text = adoSet.Fields("SLipno2").Value & ""
        sprOrder.Col = 5: sprOrder.Text = adoSet.Fields("Codenm").Value & ""
        sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("Orderno").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub Form_Load()
    
    dtDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
End Sub

Private Sub sprOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sJeobsuDt           As String
    Dim sSLipno1            As String
    Dim sSLipno2            As String
    Dim sOrderno            As String
    
    
    
    If Row = 0 Then Exit Sub
    
    Call Spread_Set_Clear(sprDt)
    
    sprOrder.Row = Row
    sprOrder.Col = 2: sJeobsuDt = sprOrder.Text
    sprOrder.Col = 3: sSLipno1 = sprOrder.Text
    sprOrder.Col = 4: sSLipno2 = sprOrder.Text
    sprOrder.Col = 6: sOrderno = sprOrder.Text

    strSql = ""
    strSql = strSql & " SELECT DISTINCT a.RoutinCD Code, b.RoutinNM ItemName, a.Verify"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine     b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    a.SLipno2  = " & Val(sSLipno2)
    strSql = strSql & " AND    a.RoutinCd = b.RoutinCD"
    strSql = strSql & " UNION ALL"
    strSql = strSql & " SELECT a.ItemCd Code, b.ItemNM ItemName, a.Verify"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    a.SLipno2  = " & Val(sSLipno2)
    strSql = strSql & " AND    a.ItemCd   = b.Codeky"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprDt.Row = sprDt.DataRowCnt + 1
        sprDt.Col = 1: sprDt.Text = adoSet.Fields("Code").Value & ""
        sprDt.Col = 2: sprDt.Text = adoSet.Fields("ItemName").Value & ""
        sprDt.Col = 3: sprDt.Text = adoSet.Fields("Verify").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub

Private Sub txtPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Trim(txtPtno.Text) = "" Then Exit Sub
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        txtSname.Text = ""
        txtSexage.Text = ""
        GoSub Get_Patient_Data
        cmdQuery.SetFocus
    End If
    Exit Sub
    

Get_Patient_Data:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_IDNOMST"
    strSql = strSql & " WHERE  Ptno = '" & txtPtno.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSexage.Text = adoSet.Fields("Sex").Value & "" & "/" & _
                     adoSet.Fields("ageYY").Value & ""
    Call adoSetClose(adoSet)
    
    Return
    
End Sub
