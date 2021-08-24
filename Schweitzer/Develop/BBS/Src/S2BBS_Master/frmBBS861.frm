VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS861 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "#Caption"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS861.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   1875
      Left            =   1380
      TabIndex        =   13
      Top             =   4260
      Width           =   7935
      Begin VB.CheckBox chkAddInfo2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Yes"
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   1170
         Width           =   2355
      End
      Begin VB.CheckBox chkAddInfo1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Yes"
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   810
         Width           =   2355
      End
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1350
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Top             =   390
         Width           =   6210
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "AAAAAAAAAA"
         Top             =   15
         Width           =   1470
      End
      Begin VB.CheckBox chkExp 
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기여부"
         Height          =   225
         Left            =   4920
         TabIndex        =   6
         Top             =   870
         Width           =   1110
      End
      Begin VB.TextBox txtAddInfo1 
         Height          =   330
         Left            =   1350
         ScrollBars      =   2  '수직
         TabIndex        =   2
         Top             =   810
         Width           =   2430
      End
      Begin VB.TextBox txtAddInfo2 
         Height          =   330
         Left            =   1350
         ScrollBars      =   2  '수직
         TabIndex        =   3
         Top             =   1170
         Width           =   2430
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   195
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   450
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   344
         BackColor       =   14411494
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
         AutoSize        =   -1  'True
         Caption         =   "#명칭"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCode 
         Height          =   315
         Left            =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   15
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   14411494
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
         AutoSize        =   -1  'True
         Caption         =   "#코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExp 
         Height          =   315
         Left            =   4890
         TabIndex        =   16
         Top             =   1170
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         BackColor       =   14411494
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
         Caption         =   "폐기일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpExp 
         Height          =   315
         Left            =   6075
         TabIndex        =   7
         Top             =   1170
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   20840451
         CurrentDate     =   36833
      End
      Begin MedControls1.LisLabel lblAddInfo1 
         Height          =   195
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   870
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   344
         BackColor       =   14411494
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
         AutoSize        =   -1  'True
         Caption         =   "#추가정보1"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAddInfo2 
         Height          =   195
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   344
         BackColor       =   14411494
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
         AutoSize        =   -1  'True
         Caption         =   "#추가정보2"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   4680
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   6765
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   6000
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   6765
      Width           =   1260
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   480
      Left            =   3360
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   6765
      Width           =   1260
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   375
      Left            =   660
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "#Title"
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2835
      Left            =   720
      TabIndex        =   11
      Top             =   1080
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   5001
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "코드"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "명칭"
         Object.Width           =   7057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "추가정보1"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "추가정보2"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "폐기일자"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2535
      Index           =   0
      Left            =   675
      Top             =   3960
      Width           =   9435
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2955
      Index           =   2
      Left            =   675
      Top             =   1020
      Width           =   9435
   End
End
Attribute VB_Name = "frmBBS861"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public CDINDEX As String

Private bkValue As String
Private objcom003 As New clsCom003
Private inputType1 As String
Private inputType2 As String


Private First As Boolean


Private Sub dtpExp_Change()
    If DeleteDate_Handle = False Then
       MsgBox "폐기일자는 현재 이후 일자만 선택 가능합니다!", vbCritical, "입력오류"
       dtpExp.Value = Format(GetSystemDate, "yyyy-mm-dd")
    End If
End Sub

Private Sub dtpExp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    Dim SSQL As String
    Dim Text1 As String
    Dim field1 As String
    Dim RS As Recordset
    Dim objSql As clsBBSMSTStatement
    
    If First = False Then Exit Sub
    
    First = False
    
    chkAddInfo1.Left = txtAddInfo1.Left
    chkAddInfo1.Top = txtAddInfo1.Top
    chkAddInfo2.Left = txtAddInfo2.Left
    chkAddInfo2.Top = txtAddInfo2.Top
        
    Call ClearAll
    Call Query
    
    Set RS = New Recordset
    Set objSql = New clsBBSMSTStatement
    
    Set RS = objSql.Get_FrmBBS861(CDINDEX)
    Set objSql = Nothing
    
    If RS.RecordCount < 1 Then
        MsgBox "마스터 설정이 올바르지 않아서 프로그램을 수행할 수 없습니다.", vbCritical, Me.Caption
        Set RS = Nothing
        Unload Me
        Exit Sub
    End If
    
        
    Text1 = RS.Fields("text1").Value & ""
    field1 = RS.Fields("field1").Value & ""
    
    Set RS = Nothing
    
    Me.Caption = field1 & " 등록"
'    medMain.lblSubMenu.Caption = Me.Caption
    
    lblTitle.Caption = field1
    
    lblCode.Caption = medGetP(Text1, 1, ";")
    lblName.Caption = medGetP(Text1, 2, ";")
    lblAddInfo1.Caption = medGetP(medGetP(Text1, 3, ";"), 1, ","): inputType1 = medGetP(medGetP(Text1, 3, ";"), 2, ",")
    lblAddInfo2.Caption = medGetP(medGetP(Text1, 4, ";"), 1, ","): inputType2 = medGetP(medGetP(Text1, 3, ";"), 2, ",")
    
    
    With lvw
        .ColumnHeaders(1).Text = lblCode.Caption
        .ColumnHeaders(2).Text = lblName.Caption
        .ColumnHeaders(3).Text = lblAddInfo1.Caption
        .ColumnHeaders(4).Text = lblAddInfo2.Caption
    
        If lblAddInfo1.Caption = "" Then
            lblAddInfo1.Visible = False
            chkAddInfo1.Visible = False
            txtAddInfo1.Visible = False
            .ColumnHeaders(2).Width = .ColumnHeaders(2).Width + .ColumnHeaders(3).Width
            .ColumnHeaders(3).Width = 0
        Else
            If inputType1 = "CHECK" Then
                chkAddInfo1.Visible = True
                txtAddInfo1.Visible = False
            Else
                chkAddInfo1.Visible = False
                txtAddInfo1.Visible = True
            End If
        End If
    
        If lblAddInfo2.Caption = "" Then
            lblAddInfo2.Visible = False
            chkAddInfo2.Visible = False
            txtAddInfo2.Visible = False
            .ColumnHeaders(2).Width = .ColumnHeaders(2).Width + .ColumnHeaders(4).Width
            .ColumnHeaders(4).Width = 0
        Else
            If inputType2 = "CHECK" Then
                chkAddInfo2.Visible = True
                txtAddInfo2.Visible = False
            Else
                chkAddInfo2.Visible = False
                txtAddInfo2.Visible = True
            End If
        End If
        
    End With
    
End Sub

Private Sub Form_Load()
    First = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objcom003 = Nothing
End Sub

Private Sub chkExp_Click()
    If chkExp.Value = 0 Then
       lblExp.Visible = False
       dtpExp.Visible = False
    Else
       lblExp.Visible = True
       dtpExp.Visible = True
       
       dtpExp.Value = GetSystemDate
       dtpExp.Enabled = True
    End If
End Sub

Private Sub chkExp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtCode = ""
    Clear
End Sub

Private Sub cmdSave_Click()
    If InputCheck = False Then
       MsgBox "모든 값을 입력해야 합니다!", vbInformation, Me.Caption
    Else
        If Save = True Then
            Call Query(Trim(txtCode))
        Else
        End If
    End If
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static i As Integer
    
    With lvw
         .SortKey = ColumnHeader.Index - 1
         .SortOrder = IIf(i = 0, lvwAscending, lvwDescending)
         .Sorted = True
    End With
    
    i = (i + 1) Mod 2
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
        txtCode = .Text
        txtName = .SubItems(1)
        If inputType1 <> "CHECK" Then
            txtAddInfo1 = .SubItems(2)
        Else
            chkAddInfo1.Value = Val(.SubItems(2))
        End If
        If inputType2 <> "CHECK" Then
            txtAddInfo2 = .SubItems(3)
        Else
            chkAddInfo2.Value = Val(.SubItems(3))
        End If
        
        If Trim(.SubItems(4)) = "" Then
            chkExp.Value = 0
        Else
            chkExp.Value = 1
            dtpExp = .SubItems(4)
        End If
    End With
End Sub

Private Sub txtAddInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCode_GotFocus()
    bkValue = txtCode
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCode_LostFocus()
    Dim itmX As ListItem
    
    If txtCode = "" Then Clear: Exit Sub
    If bkValue = txtCode Then Exit Sub
    
    For Each itmX In lvw.ListItems
        If itmX.Text = Trim(txtCode) Then
            itmX.Selected = True
            itmX.EnsureVisible
            
            Call lvw_ItemClick(itmX)
            Exit Sub
        End If
    Next itmX
    
    Clear
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub









Private Sub ClearAll()
    lvw.ListItems.Clear
    txtCode.Text = ""
    Clear
End Sub

Private Sub Clear()
    txtName.Text = ""
    txtAddInfo1 = ""
    txtAddInfo2 = ""
    
    chkExp.Enabled = True
    chkExp.Value = 0
    
    lblExp.Visible = False
    dtpExp.Visible = False
    
End Sub

Private Function InputCheck() As Boolean
    If Trim(txtCode) = "" Then InputCheck = False: Exit Function
    If Trim(txtName) = "" Then InputCheck = False: Exit Function
'    If lblAddInfo1.Caption <> "" Then
'        If Trim(txtAddInfo1) = "" Then InputCheck = False: Exit Function
'    End If
'    If lblAddInfo2.Caption <> "" Then
'        If Trim(txtAddInfo2) = "" Then InputCheck = False: Exit Function
'    End If
    
    InputCheck = True
End Function

Private Function DeleteDate_Handle() As Boolean
    Dim PDt As String
    Dim SDt As String
    
    PDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    SDt = Format(dtpExp, PRESENTDATE_FORMAT)
    
    If SDt < PDt Then
       DeleteDate_Handle = False
       Exit Function
    End If
    
    DeleteDate_Handle = True
End Function

Private Sub Query(Optional ByVal cdval1 As String = "")
    Dim i As Long
    Dim itmX As ListItem
    Dim DrRS As Recordset
    
    ClearAll
    
    Set DrRS = objcom003.OpenRecordSet(CDINDEX, , 1)
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        For i = 1 To .RecordCount
            Set itmX = lvw.ListItems.Add()
            
            itmX.Text = .Fields("cdval1").Value & ""
            itmX.SubItems(1) = .Fields("field1").Value & ""
            itmX.SubItems(2) = .Fields("field2").Value & ""
            itmX.SubItems(3) = .Fields("field3").Value & ""
            itmX.SubItems(4) = Format(.Fields("field5").Value & "", "####-##-##")
            
            If cdval1 <> "" Then
                itmX.Selected = True
                itmX.EnsureVisible
            End If
            
            .MoveNext
        Next i
    End With
    
    Set DrRS = Nothing
    
    If Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Function Save() As Boolean

On Error GoTo Save_error

    DBConn.BeginTrans
    
    With objcom003
        .CDINDEX = CDINDEX
        .cdval1 = txtCode
        .field1 = txtName
        .field2 = IIf(inputType1 <> "CHECK", txtAddInfo1, chkAddInfo1.Value)
        .Field3 = IIf(inputType2 <> "CHECK", txtAddInfo2, chkAddInfo2.Value)
        .Field5 = IIf(chkExp.Value = 0, "", Format(dtpExp, PRESENTDATE_FORMAT))
        
        If .Save() = False Then GoTo Save_error
    End With

    DBConn.CommitTrans
    Save = True
    Exit Function
    
Save_error:
    DBConn.RollbackTrans
    Save = False
End Function





