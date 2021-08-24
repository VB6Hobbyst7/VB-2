VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS802 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사건물설정"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS802.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00F4F0F2&
      Caption         =   ">>"
      Height          =   1560
      Left            =   3720
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   2460
      Width           =   480
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5790
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "128"
      Top             =   6480
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   4440
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   6480
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvwBuilding 
      Height          =   5055
      Left            =   540
      TabIndex        =   1
      Top             =   900
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "건물"
         Object.Width           =   5292
      EndProperty
   End
   Begin FPSpread.vaSpread tblBuilding 
      Height          =   4950
      Left            =   4320
      TabIndex        =   0
      Top             =   900
      Width           =   6075
      _Version        =   196608
      _ExtentX        =   10716
      _ExtentY        =   8731
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   20
      OperationMode   =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS802.frx":076A
   End
End
Attribute VB_Name = "frmBBS802"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click()
    Dim itmX As ListItem
    
    Set itmX = lvwBuilding.SelectedItem
    
    If itmX Is Nothing Then Exit Sub
    
    With tblBuilding
        .Row = .ActiveRow
        .Col = 2
        .Value = itmX.Text
    End With
    
End Sub

Private Sub cmdSave_Click()
    If Save = True Then Query
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_Load()
    Call ClearAll
    Call Query
End Sub

Private Sub ClearAll()
    lvwBuilding.ListItems.Clear
    tblBuilding.MaxRows = 20
    medClearTable tblBuilding
End Sub

Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    Dim itmX As ListItem
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_CENTER)
    Set objcom003 = Nothing
    
    ClearAll
    
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        For i = 1 To .RecordCount
            Set itmX = lvwBuilding.ListItems.Add()
            itmX.Text = .Fields("cdval1").Value & "" & " " & .Fields("field1").Value & ""
            .MoveNext
        Next i
    
        .MoveFirst
        
        For i = 1 To .RecordCount
            With tblBuilding
                .Row = i
                If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
                
            End With
            
            tblBuilding.Col = 1
            tblBuilding.Value = .Fields("cdval1").Value & "" & " " & .Fields("field1").Value & ""
        
            tblBuilding.Col = 2
            tblBuilding.Value = GetDestBuild(.Fields("cdval1").Value & "")
            
            .MoveNext
        Next i
    End With
    
    Set DrRS = Nothing
End Sub

Private Function GetDestBuild(ByVal cd As String) As String
    Dim RS As Recordset

    Set RS = GetDestBuilding(cd)
    
    GetDestBuild = ""
    
    If RS Is Nothing Then Exit Function
    
    With RS
        If .RecordCount <= 0 Then
            GetDestBuild = ""
        Else
            GetDestBuild = .Fields("gbuilding").Value & "" & " " & GetCenterNm(.Fields("gbuilding").Value & "")
        End If
    End With
    Set RS = Nothing

End Function

Private Function Save() As Boolean
    Dim objBInfo As clsBuildingInfo
    Dim i As Long
    
    DBConn.BeginTrans
    
    Set objBInfo = New clsBuildingInfo
    
    With tblBuilding
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1:  objBInfo.SBuilding = medGetP(.Value, 1, " ")
            .Col = 2:  objBInfo.GBuilding = medGetP(.Value, 1, " ")
            
            If objBInfo.Save = False Then GoTo Save_error
        Next i
    End With
    
    Set objBInfo = Nothing
    
    DBConn.CommitTrans
    Save = True
    Exit Function
    
Save_error:
    DBConn.RollbackTrans
    Save = False
End Function


Private Function GetDestBuilding(ByVal code As String) As Recordset
    Dim SSQL As String
    
    If code = "" Then Set GetDestBuilding = Nothing: Exit Function
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS004 & " " & _
           "WHERE " & DBW("sbuilding=", code)
           
    Set GetDestBuilding = New Recordset
    Call GetDestBuilding.Open(SSQL, DBConn)
    
'    If GetDestBuilding.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetDestBuilding = Nothing
'    End If
End Function

