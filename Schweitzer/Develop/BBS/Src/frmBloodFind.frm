VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmBloodFind 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "혈액조회"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ListView lvwHosExp 
      Height          =   3135
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "혈액번호"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "component"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ABO/Rh"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "입고일시"
         Object.Width           =   2470
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   3735
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "128"
      Top             =   3300
      Width           =   1320
   End
   Begin VB.CommandButton cmdSel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "선택(&S)"
      Height          =   510
      Left            =   1125
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "124"
      Top             =   3300
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   510
      Left            =   2415
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "15101"
      Top             =   3300
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvwBlood 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "혈액번호"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "환자ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "환자명"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "component"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmBloodFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'INPUT용 변수
Public mode As Long         '0:출고용 혈액 조회
                            '1:반환용 혈액 조회
                            '2:폐기용 혈액 조회
                            '3:회수용 혈액 조회
Public HosExp As Boolean    '혈액자체폐기시 조회


'RETURN용 변수
Public BldSrc As String
Public BldYY As String
Public BldNo As String
Public Compo As String
Public isSelected As Boolean

Private Enum EMode
    modeDELIVERY = 0
    modeRETURN = 1
    modeEXPIRE = 2
    modeBAGRETURN = 3
End Enum
Private First As Boolean




Private Sub cmdExit_Click()
    isSelected = False
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call Query
End Sub

Private Sub cmdSel_Click()
    Dim iTmx As ListItem
    
    
    If HosExp = True Then
        Set iTmx = lvwHosExp.SelectedItem
        If Not (iTmx Is Nothing) Then
            BldSrc = medGetP(iTmx.Text, 1, "-")
            BldYY = medGetP(iTmx.Text, 2, "-")
            BldNo = medGetP(iTmx.Text, 3, "-")
            Compo = iTmx.SubItems(1)
            isSelected = True
            Unload Me
        End If
    Else
        Set iTmx = lvwBlood.SelectedItem
        If Not (iTmx Is Nothing) Then
            BldSrc = medGetP(iTmx.Text, 1, "-")
            BldYY = medGetP(iTmx.Text, 2, "-")
            BldNo = medGetP(iTmx.Text, 3, "-")
                Compo = iTmx.SubItems(3)
            isSelected = True
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Activate()
    If First = False Then Exit Sub
    
    First = False
    Call Query
End Sub

Private Sub Form_Load()
    First = True
End Sub

Private Sub Query()
    Dim i As Long
    Dim iTmx As ListItem
    Dim DrRS As Recordset
    Dim objBldDelivery As clsBldDelivery
    
    Set objBldDelivery = New clsBldDelivery
'    Set objBldDelivery.DrDB = DBConn
    
    lvwBlood.ListItems.Clear
    If mode = 2 And HosExp = True Then
    '혈액자체폐기시 조회
        lvwBlood.Visible = False: lvwHosExp.Visible = True
        Set DrRS = objBldDelivery.GetCmdExpireHospital
        If DrRS Is Nothing Then Exit Sub
        With DrRS
            .MoveFirst
            For i = 1 To .RecordCount
                Set iTmx = lvwHosExp.ListItems.Add()
                iTmx.Text = .Fields("bldsrc").value & "" & "-" & .Fields("bldyy").value & "" & "-" & .Fields("bldno").value & ""
                iTmx.SubItems(1) = .Fields("compocd").value & "" & " " & .Fields("componm").value & ""
                iTmx.SubItems(2) = .Fields("abo").value & "" & .Fields("rh").value & ""
                iTmx.SubItems(3) = Format(.Fields("entdt").value & "", "####-##-##") & " " & _
                                   Format(Mid(.Fields("enttm").value & "", 1, 4), "00:00")
                .MoveNext
            Next i
        End With
    Else
        lvwBlood.Visible = True: lvwHosExp.Visible = False
        Set DrRS = objBldDelivery.GetBloodList(mode)
        If DrRS Is Nothing Then Exit Sub
                
        lvwBlood.SortKey = 2
        lvwBlood.Sorted = True
                
        With DrRS
            For i = 1 To .RecordCount
                Set iTmx = lvwBlood.ListItems.Add()
                iTmx.Text = .Fields("bldsrc").value & "" & "-" & .Fields("bldyy").value & "" & "-" & .Fields("bldno").value & ""
                iTmx.SubItems(1) = .Fields("ptid").value & ""
                iTmx.SubItems(2) = GetPtNm(.Fields("ptid").value & "")
                iTmx.SubItems(3) = .Fields("compocd").value & "" & " " & .Fields("componm").value & ""
                
                .MoveNext
            Next i
        End With
    End If
    
    Set DrRS = Nothing
    Set objBldDelivery = Nothing
    
End Sub

Private Sub Query4Delivery()
    Dim SSQL As String
    Dim DrRS As Recordset
    
End Sub

Private Sub lvwBlood_DblClick()
    cmdSel_Click
End Sub

Private Sub lvwBlood_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSel_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmdExit_Click
    End If
End Sub

Private Sub lvwHosExp_DblClick()
    cmdSel_Click
End Sub

Private Sub lvwHosExp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSel_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmdExit_Click
    End If
End Sub
