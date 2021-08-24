VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm341Common1 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "Lis341.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EBF3ED&
      Height          =   690
      Left            =   3105
      ScaleHeight     =   630
      ScaleWidth      =   7320
      TabIndex        =   28
      Top             =   7800
      Width           =   7380
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00DBE6E6&
         Caption         =   "지움(&C)"
         Height          =   510
         Left            =   4665
         Style           =   1  '그래픽
         TabIndex        =   32
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   2010
         Style           =   1  '그래픽
         TabIndex        =   31
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   3345
         Style           =   1  '그래픽
         TabIndex        =   30
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DBE6E6&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   6000
         Style           =   1  '그래픽
         TabIndex        =   29
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.ListBox lstSKey1 
      BackColor       =   &H00F7FFFF&
      Height          =   1860
      Left            =   240
      TabIndex        =   26
      Top             =   1140
      Width           =   2865
   End
   Begin VB.CommandButton cmdDelChild 
      BackColor       =   &H00F4F0F2&
      Caption         =   "하위키 까지 제거"
      Height          =   420
      Left            =   765
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   3015
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7095
      Left            =   3120
      TabIndex        =   12
      Top             =   660
      Width           =   7380
      Begin VB.CommandButton cmdPopup 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2940
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis341.frx":030A
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   465
         Width           =   300
      End
      Begin VB.CommandButton cmdPopup 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2925
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis341.frx":0894
         Style           =   1  '그래픽
         TabIndex        =   34
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdPopup 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4200
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis341.frx":0E1E
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   1860
         Width           =   300
      End
      Begin VB.TextBox txtFIndex 
         BackColor       =   &H00F7FFFF&
         Height          =   330
         Left            =   300
         TabIndex        =   0
         Top             =   465
         Width           =   2610
      End
      Begin VB.TextBox txtSIndex 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   300
         TabIndex        =   1
         Top             =   1080
         Width           =   2625
      End
      Begin VB.CheckBox chkKeyLock 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Index Lock Mode"
         Height          =   180
         Left            =   5415
         TabIndex        =   10
         Top             =   1230
         Value           =   1  '확인
         Width           =   1785
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Index           =   4
         Left            =   3630
         TabIndex        =   6
         Top             =   3840
         Width           =   3315
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Index           =   3
         Left            =   270
         TabIndex        =   5
         Top             =   3840
         Width           =   3315
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   1860
         Width           =   3930
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   2520
         Width           =   6675
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Index           =   2
         Left            =   270
         TabIndex        =   4
         Top             =   3180
         Width           =   6675
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   1230
         Index           =   5
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4455
         Width           =   6675
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   1035
         Index           =   6
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   5955
         Width           =   6660
      End
      Begin MedControls1.LisLabel lblCaption 
         Height          =   300
         Index           =   0
         Left            =   3270
         TabIndex        =   36
         Top             =   480
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCaption 
         Height          =   300
         Index           =   1
         Left            =   3240
         TabIndex        =   37
         Top             =   1095
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCaption 
         Height          =   300
         Index           =   2
         Left            =   4530
         TabIndex        =   38
         Top             =   1860
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "First Index"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   21
         Top             =   240
         Width           =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   105
         X2              =   7275
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   90
         X2              =   7275
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Secondary Index"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   20
         Top             =   855
         Width           =   1440
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   2
         Left            =   315
         TabIndex        =   19
         Top             =   1635
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   5
         Left            =   315
         TabIndex        =   18
         Top             =   3630
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   6
         Left            =   3675
         TabIndex        =   17
         Top             =   3615
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   3
         Left            =   315
         TabIndex        =   16
         Top             =   2295
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   4
         Left            =   315
         TabIndex        =   15
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   7
         Left            =   315
         TabIndex        =   14
         Top             =   4230
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Label1"
         Height          =   180
         Index           =   8
         Left            =   315
         TabIndex        =   13
         Top             =   5730
         Width           =   555
      End
   End
   Begin VB.ListBox lstSKey2 
      BackColor       =   &H00F7FFF7&
      Height          =   4560
      Left            =   240
      TabIndex        =   9
      Top             =   3855
      Width           =   2835
   End
   Begin VB.Label lblLastSKey1 
      BackStyle       =   0  '투명
      Height          =   195
      Left            =   3180
      TabIndex        =   27
      Top             =   8100
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFIndx 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "First Index"
      Height          =   225
      Left            =   345
      TabIndex        =   24
      Top             =   825
      Width           =   2760
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   1
      Left            =   240
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   2865
   End
   Begin VB.Label lblRName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00613636&
      Height          =   315
      Left            =   420
      TabIndex        =   23
      Top             =   200
      Width           =   4095
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00D191A2&
      BorderWidth     =   3
      FillColor       =   &H00F7F0F0&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   180
      Shape           =   4  '둥근 사각형
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00613636&
      Height          =   495
      Left            =   180
      TabIndex        =   22
      Top             =   210
      Width           =   4095
   End
   Begin VB.Label lblSIndx 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Second Index"
      Height          =   225
      Left            =   345
      TabIndex        =   11
      Top             =   3570
      Width           =   2760
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00EFFFEE&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   0
      Left            =   240
      Shape           =   4  '둥근 사각형
      Top             =   3465
      Width           =   2835
   End
End
Attribute VB_Name = "frm341Common1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cLableCount = 7

Private objsSQL As clsLISSqlCodeMaster
Private objProbar As clsProgress
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private mvarRKey As String
Private mvarTable As String
Private ChangeFlag As Boolean

Public Property Get RTable() As String
    RTable = mvarTable
End Property

Public Property Let RTable(ByVal vNewValue As String)
    mvarTable = vNewValue
End Property

Public Property Get Rkey() As String
    Rkey = mvarRKey
End Property

Public Property Let Rkey(ByVal vNewValue As String)
    mvarRKey = vNewValue
    Dim ii As Integer
    
    For ii = 0 To 2
        cmdPopup(ii).Visible = False: lblCaption(ii).Visible = False: lblCaption(ii).Caption = ""
    Next
    Select Case Rkey
        Case LC2_Detail, LC2_RelTest, LC2_Calculation, LC2_DoctTest, lc2_warning
            cmdPopup(0).Visible = True: cmdPopup(1).Visible = True
            lblCaption(0).Visible = True: lblCaption(1).Visible = True
        Case LC2_MultiSpc, LC2_Panel
            cmdPopup(0).Visible = True: cmdPopup(2).Visible = True
            lblCaption(0).Visible = True: lblCaption(2).Visible = True
        Case lc2_itemresult, LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, _
             LC2_TempletText3
            cmdPopup(0).Visible = True: lblCaption(0).Visible = True
    End Select
    
End Property

Public Property Get RName() As String
    RName = lblRName
End Property

Public Property Let RName(ByVal vNewValue As String)
    lblRName = vNewValue
End Property

Private Sub chkKeyLock_Click()
    txtFIndex.Locked = chkKeyLock.Value
    txtSIndex.Locked = chkKeyLock.Value
End Sub

Private Sub cmdDelChild_Click()
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer

    If lstSKey1.ListIndex < 0 Then
        MsgBox "First Index가 지정되지 않았습니다. 확인 후 처리 바랍니다."
        Exit Sub
    End If

    sMsg = lblRName & " Table에서 (" & lstSKey1.List(lstSKey1.ListIndex) & ") 와 그 하위 Data를 모두 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2

    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteAllChild
    Else
        Exit Sub
    End If
End Sub

Private Sub DeleteAllChild()
    Dim SSQL(1) As String

    Set objsSQL = New clsLISSqlCodeMaster

    SSQL(0) = objsSQL.DelALLComCdMST1(Rkey, Trim(txtFIndex))
    
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
        
        Call LoadFirstIndex
        If lstSKey1.ListCount < 1 Then
            Set objsSQL = Nothing
            Exit Sub
        End If
        lstSKey1.ListIndex = 0
        Call LoadSKey2
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objsSQL = Nothing
    
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    ChangeFlag = False
    chkKeyLock.Value = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objsSQL = Nothing
    Set objProbar = Nothing
    Set objCodeList = Nothing
End Sub

Private Sub lblRName_Change()
    LoadFieldNm
    LoadFirstIndex
    ChangeFlag = False
End Sub

Private Sub LoadFieldNm()
    Dim SSQL As String
    Dim dsSKey As Recordset
    Dim strFields As String
    Dim strTmp As String
    Dim i As Integer
    
    Set objsSQL = New clsLISSqlCodeMaster
    
    SSQL = objsSQL.GetComCdIndex("LC1", Rkey)
    
    With objsSQL
        Set dsSKey = New Recordset
        dsSKey.Open SSQL, DBConn
    End With
    
    If Not dsSKey.EOF Then
        strFields = "" & dsSKey.Fields("text1").Value
    Else
        strFields = ""
    End If
    
    Set dsSKey = Nothing
    Set objsSQL = Nothing
    
    For i = 0 To lblCap.Count - 1
        strTmp = medShift(strFields, ";")
        lblCap(i).Caption = medShift(strTmp, ":")
        If i > 1 Then
            If Trim(lblCap(i)) = "" Then
                lblCap(i).Visible = False
                txtVal(i - 2).Visible = False
                txtVal(i - 2).MaxLength = 0
            Else
                lblCap(i).Visible = True
                txtVal(i - 2).Visible = True
                txtVal(i - 2).MaxLength = Val(strTmp)
            End If
        End If
    Next
    
    lblFIndx = lblCap(0)
    lblSIndx = lblCap(1)
End Sub

Private Sub LoadFirstIndex()
    Dim i As Integer, SSQL As String
    Dim dsSKey As Recordset
    Dim itmFound As ListItem
    Dim strTmp As String
    
    Set objsSQL = New clsLISSqlCodeMaster
    
    Set objProbar = New clsProgress
    With objProbar
        .Container = MainFrm.stsbar
'        .Value = 1
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
    
    With objsSQL
        Set dsSKey = New Recordset
        dsSKey.Open .GetComCdMST1(Rkey), DBConn
    End With
    
    lstSKey1.Clear: lstSKey2.Clear
    Call ClearScreen
    txtFIndex.Locked = False
    txtSIndex.Locked = False
    
    If dsSKey.RecordCount < 1 Then
        Set dsSKey = Nothing
        Set objsSQL = Nothing
        Set objProbar = Nothing
        Exit Sub
    End If
    
    objProbar.Max = dsSKey.RecordCount
    objProbar.DisplayMessage = False
    For i = 1 To dsSKey.RecordCount
        strTmp = medListFind(lstSKey1, "" & dsSKey.Fields("cdval1").Value)
        If strTmp = "-1" Then
            Select Case Rkey
                Case LC2_MultiSpc
                    lstSKey1.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("spcnm").Value & ""
                Case LC2_Detail, LC2_Panel, lc2_itemresult, LC2_Calculation, LC2_RelTest, LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, LC2_TempletText3, lc2_warning:
                    lstSKey1.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("testnm").Value & ""
                Case Else
                    lstSKey1.AddItem "" & dsSKey.Fields("cdval1").Value
            End Select
        End If
        dsSKey.MoveNext
        objProbar.Value = i
    Next i
    
    Set dsSKey = Nothing
    Set objsSQL = Nothing
    Set objProbar = Nothing
End Sub

Private Sub cmdClear_Click()
'   If Not ConfirmExit Then Exit Sub
    lstSKey2.ListIndex = -1
    lstSKey2.Clear
    ClearScreen
    txtFIndex.Locked = False
    txtSIndex.Locked = False
    txtFIndex.SetFocus
    ChangeFlag = False
    lblCaption(0).Caption = "": lblCaption(1).Caption = "": lblCaption(2).Caption = ""
End Sub

Private Sub ClearScreen()
    Dim i As Integer

    txtFIndex = ""
    txtSIndex = ""
    
    For i = 1 To cLableCount
        txtVal(i - 1) = ""
    Next i
    
End Sub

Private Sub LostFcousClear(ByVal sIndex As String)
    Dim i As Integer

    If sIndex = "0" Then
        txtSIndex.Text = ""
        lblCaption(1).Caption = ""
    End If
    For i = 1 To cLableCount
        If sIndex = "0" Or sIndex = "1" Then
            txtVal(i - 1) = ""
        Else
            If i - 1 <> 0 Then txtVal(i - 1) = ""
        End If
    Next i
    
'    lstSKey2.ListIndex = -1
'    ClearScreen
    txtFIndex.Locked = False
    txtSIndex.Locked = False
'    txtFIndex.SetFocus
    ChangeFlag = False
: lblCaption(2).Caption = ""
    
End Sub
Private Sub cmdDelete_Click()
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer

    If Trim(txtFIndex) = "" Or Trim(txtSIndex) = "" Then
        MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
        Exit Sub
    End If

    sMsg = lblRName & " Table에서 (" & txtFIndex & " - " & txtSIndex & ") Key와 Data를 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2

    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteData
        ChangeFlag = False
    Else
        Exit Sub
    End If

    cmdClear_Click
End Sub

Private Sub DeleteData()
    Dim i As Integer, sRName As String
    Dim SSQL(1) As String
    
    Set objsSQL = New clsLISSqlCodeMaster
    
    SSQL(0) = objsSQL.DelComCdMST1(Rkey, Trim(txtFIndex.Text), Trim(txtSIndex.Text))
    
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
        
        sRName = Trim(txtFIndex.Text)
        Call LoadFirstIndex
        If lstSKey1.ListCount < 1 Then
            Set objsSQL = Nothing
            Exit Sub
        End If
        
        For i = 1 To lstSKey1.ListCount
            If lstSKey1.List(i) = sRName Then
                lstSKey1.ListIndex = 1
                Exit For
            End If
        Next i
        If lstSKey1.ListCount > 0 And lstSKey1.ListIndex < 0 Then lstSKey1.ListIndex = 0
        Call LoadSKey2
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objsSQL = Nothing
End Sub

Private Sub cmdExit_Click()
'   If Not ConfirmExit Then Exit Sub
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer, sIndexKey As String, sFlag As String
    Dim SSQL As String, dsChk As Recordset
    
    If txtVal(0).Visible Then
        If Trim(txtFIndex.Text) = "" Or Trim(txtSIndex.Text) = "" Or Trim(txtVal(0).Text) = "" Then
            MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
            Exit Sub
        End If
    Else
        If Trim(txtFIndex.Text) = "" Or Trim(txtSIndex.Text) = "" Then
            MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
            Exit Sub
        End If
    End If
    
    sFlag = "I"
    
    Set objsSQL = New clsLISSqlCodeMaster
        
    With objsSQL
        Set dsChk = New Recordset
        dsChk.Open .GetComCdMST1(Rkey, Trim(txtFIndex.Text), Trim(txtSIndex.Text)), DBConn
    End With
            
    If dsChk.RecordCount = 1 Then sFlag = "U"
    
    Set dsChk = Nothing
    Set objsSQL = Nothing

    Select Case sFlag
        Case "I"
            Call CommonInsert
        Case "U"
            Call CommonUpdate
        Case Else: MsgBox "시스템에 오류가 있습니다."
    End Select

    ChangeFlag = False
    cmdClear_Click
End Sub

Private Sub CommonInsert()
    Dim i As Integer, sRName As String
    Dim sData As String
    Dim SSQL(1) As String

    Set objsSQL = New clsLISSqlCodeMaster
    
    SSQL(0) = objsSQL.SetComCdMST1(False, Rkey, Trim(txtFIndex.Text), Trim(txtSIndex.Text), Trim(txtVal(0).Text), _
                                  Trim(txtVal(1).Text), Trim(txtVal(2).Text), Trim(txtVal(3).Text), _
                                  Trim(txtVal(4).Text), RTrim(txtVal(5).Text), RTrim(txtVal(6).Text))
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
    
        lblLastSKey1 = txtFIndex
        
        sRName = Trim(txtFIndex.Text)
        Call LoadFirstIndex
        
        If lstSKey1.ListCount < 1 Then
            Set objsSQL = Nothing
            Exit Sub
        End If
        
        For i = 1 To lstSKey1.ListCount
            If lstSKey1.List(i) = sRName Then
                lstSKey1.ListIndex = i
                Exit For
            End If
        Next i
        If lstSKey1.ListCount > 0 And lstSKey1.ListIndex < 0 Then lstSKey1.ListIndex = 0
        Call LoadSKey2
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objsSQL = Nothing
End Sub

Private Sub CommonUpdate()
    Dim sData As String
    Dim SSQL(1) As String
    
    Set objsSQL = New clsLISSqlCodeMaster
        
    SSQL(0) = objsSQL.SetComCdMST1(True, Rkey, Trim(txtFIndex.Text), Trim(txtSIndex.Text), Trim(txtVal(0).Text), _
                                Trim(txtVal(1).Text), Trim(txtVal(2).Text), Trim(txtVal(3).Text), _
                                Trim(txtVal(4).Text), RTrim(txtVal(5).Text), RTrim(txtVal(6).Text))
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
        lblLastSKey1 = txtFIndex
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objsSQL = Nothing
End Sub

Private Sub lstSKey1_Click()
    If lstSKey1.ListIndex > -1 Then
        lblLastSKey1 = medGetP(lstSKey1.List(lstSKey1.ListIndex), 1, Chr$(9))
    End If
End Sub

Private Sub lstSKey1_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Not ConfirmExit Then Exit Sub
    If KeyCode = vbKeyReturn Then LoadSKey2
    txtFIndex = medGetP(lstSKey1.List(lstSKey1.ListIndex), 1, Chr$(9))
    ChangeFlag = False
End Sub

Private Sub lstSKey1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not ConfirmExit Then Exit Sub
    lblCaption(1).Caption = "": lblCaption(2).Caption = ""
    If Button = vbLeftButton Then LoadSKey2
    txtFIndex = medGetP(lstSKey1.List(lstSKey1.ListIndex), 1, Chr$(9))
    lblCaption(0).Caption = medGetP(lstSKey1.List(lstSKey1.ListIndex), 2, Chr$(9))
    ChangeFlag = False
End Sub

Private Sub LoadSKey2()
    Dim i As Integer
    Dim SSQL As String
    Dim dsSKey As Recordset
    
    If lstSKey1.ListIndex < 0 Then Exit Sub
    
    Set objsSQL = New clsLISSqlCodeMaster
    With objsSQL
        Set dsSKey = New Recordset
        dsSKey.Open .GetComCdMST1(Rkey, medGetP(lstSKey1.List(lstSKey1.ListIndex), 1, Chr$(9))), DBConn
    End With
    
    Call ClearScreen
    
    lstSKey2.Clear
    
    If dsSKey.RecordCount < 1 Then
        Set dsSKey = Nothing
        Set objsSQL = Nothing
        Exit Sub
    End If
    
    For i = 1 To dsSKey.RecordCount
        Select Case Rkey
            Case LC2_MultiSpc
                lstSKey2.AddItem "" & dsSKey.Fields("cdval2").Value & vbTab & dsSKey.Fields("spcnm").Value & ""
            Case lc2_itemresult, LC2_Calculation, LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, LC2_TempletText3
                lstSKey2.AddItem "" & dsSKey.Fields("cdval2").Value & vbTab & dsSKey.Fields("field1").Value & ""
            Case LC2_Panel
                lstSKey2.AddItem "" & dsSKey.Fields("cdval2").Value & vbTab & dsSKey.Fields("field1").Value & "" & vbTab & dsSKey.Fields("testnm").Value & ""
            Case LC2_Detail, LC2_RelTest
                lstSKey2.AddItem "" & dsSKey.Fields("cdval2").Value & vbTab & dsSKey.Fields("testnm").Value & ""
            Case lc2_warning
                lstSKey2.AddItem "" & dsSKey.Fields("cdval2").Value & vbTab & dsSKey.Fields("rstnm").Value & ""
            Case Else
                lstSKey2.AddItem "" & dsSKey.Fields("cdval2").Value
        End Select
        
        dsSKey.MoveNext
    Next i
    
    txtFIndex.Locked = chkKeyLock.Value
    txtSIndex.Locked = False

    Set dsSKey = Nothing
    Set objsSQL = Nothing
End Sub

Private Sub lstSKey2_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Not ConfirmExit Then Exit Sub
    If KeyCode = vbKeyReturn Then LoadData
    ChangeFlag = False
End Sub

Private Sub lstSKey2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Not ConfirmExit Then Exit Sub
    If Button = vbLeftButton Then LoadData
    ChangeFlag = False
End Sub

Private Sub LoadData()
    Dim SSQL As String
    Dim dsInfo As Recordset
    Dim strLastSkey1 As String
    
    If lstSKey2.ListIndex < 0 Then Exit Sub
    strLastSkey1 = medGetP(lstSKey1.List(lstSKey1.ListIndex), 1, Chr$(9))
    
    If strLastSkey1 = "" Then
        strLastSkey1 = lblLastSKey1
    End If
    
    Set objsSQL = New clsLISSqlCodeMaster
    
    With objsSQL
        SSQL = .GetComCdMST1(Rkey, strLastSkey1, medGetP(lstSKey2.List(lstSKey2.ListIndex), 1, Chr$(9)))
        Set dsInfo = New Recordset
        dsInfo.Open SSQL, DBConn
    End With
    
    Call ClearScreen
    
    If dsInfo.RecordCount < 1 Then
        MsgBox "등록되어 있지 않은 ID입니다." & vbNewLine & SSQL
        Exit Sub
    End If
    
    txtFIndex.Text = "" & dsInfo.Fields("cdval1").Value
    txtSIndex.Text = "" & dsInfo.Fields("cdval2").Value:
    Select Case Rkey
        Case lc2_warning
            lblCaption(1).Caption = "" & dsInfo.Fields("rstnm").Value
        Case Else
            lblCaption(1).Caption = GetDetailCodeName("1", txtSIndex.Text)
    End Select
    txtVal(0).Text = "" & dsInfo.Fields("field1").Value: lblCaption(2).Caption = GetDetailCodeName("2", txtVal(0).Text)
    txtVal(1).Text = "" & dsInfo.Fields("field2").Value
    txtVal(2).Text = "" & dsInfo.Fields("field3").Value
    txtVal(3).Text = "" & dsInfo.Fields("field4").Value
    txtVal(4).Text = "" & dsInfo.Fields("field5").Value
    txtVal(5).Text = "" & dsInfo.Fields("text1").Value
    txtVal(6).Text = "" & dsInfo.Fields("text2").Value
    
    ' 데이타 읽고 나서 키를 바꿀수 없게..
    txtFIndex.Locked = chkKeyLock.Value
    txtSIndex.Locked = chkKeyLock.Value
    
    Set dsInfo = Nothing
    Set objsSQL = Nothing
End Sub

Private Sub txtFIndex_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If Not ConfirmExit Then
'       KeyAscii = 0
'       Exit Sub
'    End If
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtFIndex_LostFocus()
    lstSKey1.ListIndex = medListFind(lstSKey1, txtFIndex.Text)
    If txtFIndex.Text = medGetP(lstSKey1.List(lstSKey1.ListIndex), 1, vbTab) Then
       Call lstSKey1_MouseUp(1, 0, 0, 0)
       ChangeFlag = False
    Else
        Dim RS      As Recordset
        Dim SSQL    As String
        
        Call LostFcousClear("0")
        Set RS = New Recordset
        
        SSQL = GetPopUpListSQL(UCase(txtFIndex.Text))
        If SSQL = "" Then Exit Sub
        RS.Open SSQL, DBConn
        If Not RS.EOF Then
            txtFIndex.Text = UCase(txtFIndex.Text)
            
            If Rkey = LC2_DoctTest Then
                lblCaption(0).Caption = RS.Fields("empnm").Value & ""
            Else
                lblCaption(0).Caption = RS.Fields("testnm").Value & ""
            End If
        Else
            txtFIndex.Text = ""
            lblCaption(0).Caption = ""
        End If
        
        Set RS = Nothing
    End If
End Sub

Private Sub txtSIndex_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If Not ConfirmExit Then
'       KeyAscii = 0
'       Exit Sub
'    End If
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
   
End Sub

Private Sub txtSIndex_LostFocus()
    
    If medListFind(lstSKey1, txtFIndex.Text) >= 0 Then
        lstSKey2.ListIndex = medListFind(lstSKey2, txtSIndex.Text)
                
        If lstSKey2.ListIndex >= 0 Then
            If txtSIndex.Text <> medGetP(lstSKey2.List(lstSKey2.ListIndex), 1, vbTab) Then
                GoTo NoData
            End If
        End If
        
        If lstSKey2.ListIndex <> -1 Then
            Call LoadData
            ChangeFlag = False
        Else
NoData:
            Dim RS      As Recordset
            Dim SSQL    As String
            
            Call LostFcousClear("1")
            If txtSIndex.Text = "" Then Exit Sub
            SSQL = GetPopUpListSQL1(UCase(txtSIndex.Text))
            If SSQL = "" Then Exit Sub
            Set RS = New Recordset
            RS.Open SSQL, DBConn
            If Not RS.EOF Then
                txtSIndex.Text = UCase(txtSIndex.Text)
                lblCaption(1).Caption = RS.Fields("testnm").Value & ""
                
                If Rkey = LC2_DoctTest Then
                    txtVal(0).Text = RS.Fields("testnm").Value & ""
                End If
            Else
'                txtSIndex.Text = ""
'                lblCaption(1).Caption = ""
            End If
            
            Set RS = Nothing
        End If
    End If
End Sub


Public Sub Raise_lstSKey1_MouseUp()
    lstSKey1.ListIndex = medListFind(lstSKey1, Trim(txtFIndex.Text))
    Call lstSKey1_MouseUp(1, 0, 0, 0)
End Sub

Private Sub txtVal_Change(Index As Integer)
    ChangeFlag = True
End Sub
Private Sub txtVal_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 1 And mvarRKey = lc2_itemresult Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub txtVal_LostFocus(Index As Integer)
    If Index <> 0 Then Exit Sub
    
    Dim RS      As Recordset
    Dim SSQL    As String
    
    If Rkey = lc2_warning Then
        Exit Sub
    End If
    
    Call LostFcousClear("2")
    SSQL = GetPopUpListSQL2(UCase(txtVal(0).Text))
    If SSQL = "" Then Exit Sub
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        txtVal(0).Text = UCase(txtVal(0).Text)
        lblCaption(2).Caption = RS.Fields("testnm").Value & ""
    Else
        txtFIndex.Text = ""
        lblCaption(2).Caption = ""
    End If
    
    Set RS = Nothing
End Sub

Private Function GetDetailCodeName(ByVal sIndex As String, Optional ByVal sCode As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    GetDetailCodeName = sCode
    
    If sIndex = "0" Then
        SSQL = GetPopUpListSQL(sCode)
    ElseIf sIndex = "1" Then
        SSQL = GetPopUpListSQL1(sCode)
    ElseIf sIndex = "2" Then
        SSQL = GetPopUpListSQL2(sCode)
    End If
    If SSQL = "" Then Exit Function
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        GetDetailCodeName = RS.Fields("testnm").Value & ""
    End If
    
    Set RS = Nothing
End Function
'POPUP1 찾기
Private Function GetPopUpListSQL(Optional ByVal sCode As String = "") As String
    Dim SSQL As String

    Select Case Rkey
        Case LC2_Detail
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            SSQL = SSQL & " and " & DBW("a.panelfg=", "D")
            
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sCode)
            
            SSQL = SSQL & " order by testcd"
            
        Case LC2_RelTest, LC2_Calculation, lc2_itemresult, _
             LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, LC2_TempletText3, lc2_warning
            '검사항목 아이템 찾기
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            Select Case Rkey
                Case LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, LC2_TempletText3
                    SSQL = SSQL & " and " & DBW("a.testdiv=", "1")
            End Select
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sCode)
        Case LC2_MultiSpc
            SSQL = " select cdval1 as testcd ,field4 as testnm from " & T_LAB032 & _
                   " where " & _
                           DBW("cdindex", LC3_Specimen, 2) & _
                   " and " & DBW("field1=", "Y")
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("cdval1=", sCode)
        Case LC2_Panel
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            SSQL = SSQL & " and " & DBW("a.panelfg=", "G")
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sCode)
            SSQL = SSQL & " order by testcd"
        Case LC2_DoctTest
            '-- Group ID 는 일단 Fix ㅡ.ㅡ ...
            SSQL = " select a.empid, b.empnm " & _
                   "   from " & T_COM010 & " a, " & T_COM006 & " b " & _
                   "  where a.groupid = 'G004' "
            
            If sCode <> "" Then
                SSQL = SSQL & "    and a.empid = " & DBS(sCode)
            End If
            
            SSQL = SSQL & "    and b.empid = a.empid " & _
                          "  order by a.empid "
        
    End Select
    GetPopUpListSQL = SSQL
End Function

'POPUP2 찾기
Private Function GetPopUpListSQL1(Optional ByVal sCode As String = "") As String
    Dim SSQL As String

    Select Case Rkey
        Case LC2_Detail, LC2_Panel, LC2_RelTest, LC2_DoctTest
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            If LC2_Panel = Rkey Then
                SSQL = SSQL & " and " & DBW("a.panelfg<>", "G")
            ElseIf LC2_Detail = Rkey Then
                SSQL = SSQL & " and " & DBW("a.detailfg=", "*")
            ElseIf LC2_RelTest = Rkey Then
            End If
            
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sCode)
            SSQL = SSQL & " order by testcd"
        Case LC2_Calculation
            SSQL = " Select a.spccd as testcd, b.field3 as testnm " & _
                   " From  " & T_LAB004 & " a, " & T_LAB032 & " b " & _
                   " Where " & DBW("a.testcd", Trim(UCase(txtFIndex)), 2) & _
                   " and   " & DBW("b.cdindex", LC3_Specimen, 2) & _
                   " and     b.cdval1 = a.spccd "
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("b.cdval1=", sCode)
        Case lc2_warning
            SSQL = " Select cdval2 as testcd, field1 as testnm " & _
                   " From  " & T_LAB031 & _
                   " Where " & DBW("cdindex", lc2_itemresult, 2)
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("cdval1=", sCode)
    End Select
    GetPopUpListSQL1 = SSQL
End Function

'POPUP3 찾기
Private Function GetPopUpListSQL2(Optional ByVal sCode As String = "") As String
    Dim SSQL As String

    Select Case Rkey
        Case LC2_Panel
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            SSQL = SSQL & " and (" & DBW("a.panelfg<>", "G") & " or a.panelfg is null)"
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sCode)
            SSQL = SSQL & " order by testcd"
    
        Case LC2_Detail, LC2_RelTest
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sCode)
        Case LC2_MultiSpc
            SSQL = " select cdval1 as testcd ,field4 as testnm from " & T_LAB032 & _
                   " where " & _
                           DBW("cdindex", LC3_Specimen, 2)
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("cdval1=", sCode)
        Case lc2_warning
            SSQL = " Select cdval2 as testcd, field1 as testnm " & _
                   " From  " & T_LAB031 & _
                   " Where " & DBW("cdindex", lc2_itemresult, 2)
            If sCode <> "" Then SSQL = SSQL & " and " & DBW("cdval1=", sCode)
    End Select
    GetPopUpListSQL2 = SSQL
End Function

Private Sub cmdPopup_Click(Index As Integer)
    Dim tmpSql  As String
    Dim lngTop  As Long
    Dim lngLeft As Long
    
    Set objCodeList = New clsPopUpList
    With objCodeList
            .FormCaption = "코드 리스트"
            .ColumnHeaderText = "코드;코드명"
            .Connection = DBConn
        Select Case Index
            Case 0: tmpSql = GetPopUpListSQL:
                    lngTop = txtFIndex.Top + 2350
                    lngLeft = Me.Left + Frame1.Left + txtFIndex.Left + 50
                    txtFIndex.Text = "": txtVal(0).Text = ""
                    lblCaption(0).Caption = "": lblCaption(1).Caption = ""
                    Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
                    txtFIndex.Text = medGetP(.SelectedString, 1, ";")
                    lblCaption(0).Caption = medGetP(.SelectedString, 2, ";")
                    If LC3_ReportTesctCd = Rkey Then
                        txtVal(1).Text = "": lblCaption(2).Caption = ""
                        txtVal(0).Text = medGetP(.SelectedString, 2, ";")
                    End If
                    Call txtFIndex_LostFocus
            Case 1:
                    If Rkey = lc2_warning Then
                        If Trim(txtFIndex.Text) = "" Then
                            Exit Sub
                        End If
                        tmpSql = GetPopUpListSQL1(Trim(txtFIndex.Text))
                    Else
                        tmpSql = GetPopUpListSQL1
                    End If
                    lngTop = txtSIndex.Top + 2350
                    lngLeft = Me.Left + Frame1.Left + txtSIndex.Left + 50
                    lblCaption(1).Caption = "": txtSIndex.Text = ""
                    Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
                    txtSIndex.Text = medGetP(.SelectedString, 1, ";")
                    lblCaption(1).Caption = medGetP(.SelectedString, 2, ";")
                    
                    If Rkey = LC2_DoctTest Then
                        txtVal(0).Text = medGetP(.SelectedString, 2, ";")
                    End If
                    
                    Call txtSIndex_LostFocus
            Case 2: tmpSql = GetPopUpListSQL2
                    lngTop = txtVal(0).Top + 2350
                    lngLeft = Me.Left + Frame1.Left + txtVal(0).Left + 50
                    lblCaption(2).Caption = "": txtVal(0).Text = ""
                    Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
                    txtVal(0).Text = medGetP(.SelectedString, 1, ";")
                    lblCaption(2).Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
    Set objCodeList = Nothing
End Sub
