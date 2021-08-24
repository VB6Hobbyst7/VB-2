VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm384OtTestcd 
   BackColor       =   &H00DBE6E6&
   Caption         =   "외부의뢰검사 단가설정"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   19
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7425
      Left            =   3600
      TabIndex        =   0
      Top             =   645
      Width           =   7215
      Begin VB.CommandButton cmdPopupList 
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
         Left            =   1515
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frm384OtTestcd.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   585
         Width           =   300
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   270
         TabIndex        =   3
         Top             =   2055
         Width           =   6510
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   285
         TabIndex        =   2
         Top             =   1455
         Width           =   1230
      End
      Begin MedControls1.LisLabel lbltestNm 
         Height          =   345
         Left            =   1830
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   585
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "원"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   1530
         Width           =   1755
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "단가"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   1230
         Width           =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   105
         X2              =   7020
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   105
         X2              =   7005
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Label lblCap 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사항목"
         Height          =   225
         Index           =   0
         Left            =   285
         TabIndex        =   5
         Top             =   300
         Width           =   3765
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "비고"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   4
         Top             =   1845
         Width           =   360
      End
   End
   Begin Crystal.CrystalReport crtReport 
      Left            =   4695
      Top             =   3810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ListView lvwTestcd 
      Height          =   6960
      Left            =   255
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1095
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   12277
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "단가"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "비고"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblSubName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사항목코드"
      Height          =   180
      Left            =   1350
      TabIndex        =   11
      Top             =   795
      Width           =   1080
   End
   Begin VB.Label lblRName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "외부의뢰검사 단가설정"
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
      Left            =   270
      TabIndex        =   10
      Top             =   210
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "☞ 총 "
      Height          =   255
      Left            =   375
      TabIndex        =   9
      Top             =   8385
      Width           =   435
   End
   Begin VB.Label lblSubKeyCnt 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   855
      TabIndex        =   8
      Top             =   8385
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "건이 조회되었습니다."
      Height          =   255
      Left            =   1275
      TabIndex        =   7
      Top             =   8385
      Width           =   1755
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00D191A2&
      BorderWidth     =   3
      FillColor       =   &H00F1F5F4&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   195
      Shape           =   4  '둥근 사각형
      Top             =   105
      Width           =   5115
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00EBEBEB&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   420
      Left            =   255
      Shape           =   4  '둥근 사각형
      Top             =   8265
      Width           =   2835
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   0
      Left            =   255
      Shape           =   4  '둥근 사각형
      Top             =   705
      Width           =   3300
   End
End
Attribute VB_Name = "frm384OtTestcd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private MySqlStmt As New clsLISSqlStatement ' SQL 클래스
Private MyItem As New clsItem               ' 검사항목 클래스

Private Sub cmdClear_Click()
    txtTestCd.Text = ""
    lbltestNm.Caption = ""
    txtVal.Text = ""
    txtRemark.Text = ""
End Sub

Private Sub cmdDelete_Click()
    Dim CDINDEX As String
    
    If txtTestCd.Text = "" Then
        MsgBox "삭제 검사항목을 설정하세요", vbInformation + vbOKOnly, "검사항목설정"
        Exit Sub
    End If
    
    
    On Error GoTo SAVE_ERROR
    
    
    dbconn.BeginTrans
    
    CDINDEX = "C249"
    dbconn.Execute MySqlStmt.TestCharge(CDINDEX, txtTestCd, txtVal.Text, lbltestNm.Caption, txtRemark.Text)
    dbconn.CommitTrans
    MsgBox "삭제되었습니다.", vbInformation + vbOKOnly, "검사단가 삭제"
    Call TestCdDisplay
    Call cmdClear_Click
    Exit Sub
    
SAVE_ERROR:
    dbconn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPopupList_Click()

    Dim tmpSql As String
    Dim lngTop As Long, lngLeft As Long

    Call cmdClear_Click
    Set objCodeList = New clsPopUpList
    With objCodeList
        lngTop = txtTestCd.Top + 2350
        lngLeft = Me.Left + Frame1.Left + txtTestCd.Left + 50
        .Connection = dbconn
        .Tag = "TestCd"
        .FormCaption = "검사항목 리스트"
        .ColumnHeaderText = "검사코드;검사명"
        tmpSql = MySqlStmt.SqlLAB001CodeList
        '.ListPop tmpSql, lngTop, lngLeft
        .LoadPopUp tmpSql ' , lngTop, lngLeft, lstItemList
        txtTestCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
        lbltestNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
        
    End With

End Sub

Private Sub cmdSave_Click()
    Dim CDINDEX As String
    
    On Error GoTo SAVE_ERROR
    
    If txtTestCd.Text = "" Then
        MsgBox "검사항목을 설정하세요", vbInformation + vbOKOnly, "검사항목설정"
        Exit Sub
    End If
    
    dbconn.BeginTrans
    
    CDINDEX = "C249"
    dbconn.Execute MySqlStmt.TestCharge(CDINDEX, txtTestCd, txtVal.Text, lbltestNm.Caption, txtRemark.Text)
    dbconn.Execute MySqlStmt.TestCharge(CDINDEX, txtTestCd, txtVal.Text, lbltestNm.Caption, txtRemark.Text, True)
    dbconn.CommitTrans
    MsgBox "저장되었습니다.", vbInformation + vbOKOnly, "검사단가 저장"
    Call TestCdDisplay
    Call cmdClear_Click
    Exit Sub
    
SAVE_ERROR:
    dbconn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Call MyItem.GetItemList(lstItemList): DoEvents
    Call TestCdDisplay
End Sub

Private Sub TestCdDisplay()
    Dim RS As Recordset
    Dim itmx As ListItem
    
    Set RS = New Recordset
    RS.Open MySqlStmt.TestChangeRs, dbconn
    
    If Not RS.EOF Then
        RS.MoveFirst
        With lvwTestcd
            .ListItems.Clear
            Do Until RS.EOF
                Set itmx = .ListItems.Add()
                itmx.Text = RS.Fields("cdval1").Value & ""
                itmx.SubItems(1) = RS.Fields("text1").Value & ""
                itmx.SubItems(2) = RS.Fields("field1").Value & ""
                itmx.SubItems(3) = RS.Fields("text2").Value & ""
                RS.MoveNext
            Loop
            
        End With
       
    End If
    lblSubKeyCnt.Caption = RS.RecordCount
    Set RS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCodeList = Nothing
    Set MySqlStmt = Nothing
    Set MyItem = Nothing
End Sub

Private Sub lvwTestcd_ItemClick(ByVal Item As MSComctlLib.ListItem)
     txtTestCd.Text = Item.Text
     lbltestNm.Caption = Item.SubItems(1)
     txtVal.Text = Item.SubItems(2)
     txtRemark.Text = Item.SubItems(3)
End Sub

