VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQryGeom 
   Caption         =   "검체코드 조회"
   ClientHeight    =   6750
   ClientLeft      =   3525
   ClientTop       =   1365
   ClientWidth     =   5550
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   5550
   Begin Threed.SSCommand cmdQuit 
      Height          =   435
      Left            =   3900
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "종료"
      BevelWidth      =   1
      Font3D          =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdQrySample 
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "조회[일반]"
      BevelWidth      =   1
      Font3D          =   1
      Outline         =   0   'False
   End
   Begin MSComctlLib.TreeView tvSample 
      Height          =   5895
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   10398
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin Threed.SSCommand cmdQrySample 
      Height          =   435
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "조회[미생물]"
      BevelWidth      =   1
      Font3D          =   1
      Outline         =   0   'False
   End
End
Attribute VB_Name = "frmQryGeom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdQrySample_Click(Index As Integer)
    Dim sText       As String
    Dim sRowID      As String
    Dim NodeX       As Node
    
    
    DoEvents
    GoSub TreeView_Select
    
    cmdQrySample(0).ForeColor = RGB(0, 0, 0)
    cmdQrySample(1).ForeColor = RGB(0, 0, 0)
    If Index = 0 Then
        cmdQrySample(0).ForeColor = RGB(255, 0, 0)
    Else
        cmdQrySample(1).ForeColor = RGB(255, 0, 0)
    End If
    
    Exit Sub
    
    
TreeView_Select:
    tvSample.Nodes.Clear
    If Index = 0 Then
        Set NodeX = tvSample.Nodes.Add(, , "A0", "일반검체 분류")
    Else
        Set NodeX = tvSample.Nodes.Add(, , "A0", "미생물검체 분류")
    End If
    
    StrSql = ""
    StrSql = StrSql & " SELECT Class2, Max(RowID) RWID"
    StrSql = StrSql & " FROM   TWEXAM_Sample"
    
    If Index = 0 Then
        StrSql = StrSql & " WHERE  CLass1 = 'a'"
    Else
        StrSql = StrSql & " WHERE  CLass1 = 'm'"
    End If
    
    StrSql = StrSql & " GROUP  BY Class2"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowID = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("Class2").Value & "")
        Set NodeX = tvSample.Nodes.Add("A0", tvwChild, "A1" & sRowID, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    tvSample.Nodes("A0").Expanded = True
    Return
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    StrSql = ""
    StrSql = StrSql & " SELECT Code, class1, Codenm, anatomy, RowID"
    StrSql = StrSql & " FROM   TWEXAM_Sample"
    StrSql = StrSql & " WHERE  Class2 = '" & tvSample.Nodes("A1" & sRowID).Text & "'"
    If Index = 0 Then
        StrSql = StrSql & " AND    CLass1 = 'a'"
    Else
        StrSql = StrSql & " AND    CLass1 = 'm'"
    End If
    StrSql = StrSql & " ORDER BY Code"
    
    If False = adoSetOpen(StrSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        
        sSubText2 = Trim(adoSubCode1.Fields("Code").Value & ".") & _
                   StrConv(Trim(adoSubCode1.Fields("Codenm").Value & ""), vbProperCase)
        If Trim(adoSubCode1.Fields("Anatomy").Value & "") <> "" Then
            sSubText2 = sSubText2 & "(" & adoSubCode1.Fields("anatomy").Value & ")"
        End If
        
        Set NodeX = tvSample.Nodes.Add("A1" & sRowID, tvwChild, sSubText1, sSubText2)
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    Return

End Sub



Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call cmdQrySample_Click(0)
    
End Sub


Private Sub tvSample_DblClick()
    Dim sCode       As String * 8
    
    If Left(tvSample.SelectedItem.Key, 2) = "B2" Then
        sCode = Left(tvSample.SelectedItem.Text, 8)
        Call SetWindowText(hWndReturn, sCode)
        Unload Me
    End If
    
End Sub
