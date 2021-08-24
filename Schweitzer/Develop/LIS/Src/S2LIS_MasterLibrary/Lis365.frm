VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm365SpcGroup 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11070
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11070
   Begin VB.OptionButton optOrderBy 
      BackColor       =   &H00DBE6E6&
      Caption         =   "접수시간순"
      Height          =   225
      Index           =   1
      Left            =   9510
      TabIndex        =   35
      Top             =   1410
      Width           =   1215
   End
   Begin VB.OptionButton optOrderBy 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검체순"
      Height          =   225
      Index           =   0
      Left            =   8550
      TabIndex        =   34
      Top             =   1410
      Value           =   -1  'True
      Width           =   930
   End
   Begin VB.CheckBox chkWcGs 
      BackColor       =   &H00DBE6E6&
      Caption         =   "여성클리닉GS"
      Height          =   240
      Left            =   8610
      TabIndex        =   33
      Top             =   705
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   14
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00F4F0F2&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "25612"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.ListBox lstGroup 
      Height          =   2040
      Left            =   3450
      TabIndex        =   8
      Top             =   8460
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.ListBox lstWA 
      BackColor       =   &H00F7FFFF&
      Height          =   2040
      Left            =   2130
      TabIndex        =   3
      Top             =   8280
      Visible         =   0   'False
      Width           =   3315
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5385
      Left            =   2400
      TabIndex        =   26
      Top             =   2670
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9499
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   14411494
      TabCaption(0)   =   "검체"
      TabPicture(0)   =   "Lis365.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "spdSpccd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "배지"
      TabPicture(1)   =   "Lis365.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spdMedia"
      Tab(1).ControlCount=   1
      Begin FPSpread.vaSpread spdMedia 
         Height          =   4815
         Left            =   -74910
         TabIndex        =   29
         Top             =   90
         Width           =   8235
         _Version        =   196608
         _ExtentX        =   14526
         _ExtentY        =   8493
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   20
         ScrollBars      =   1
         SpreadDesigner  =   "Lis365.frx":0038
      End
      Begin FPSpread.vaSpread spdSpccd 
         Height          =   4815
         Left            =   90
         TabIndex        =   28
         Top             =   90
         Width           =   8235
         _Version        =   196608
         _ExtentX        =   14526
         _ExtentY        =   8493
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   20
         ScrollBars      =   1
         SpreadDesigner  =   "Lis365.frx":04BB
      End
   End
   Begin VB.CommandButton cmdWAhelp 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   270
      Left            =   3570
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   1020
      Width           =   345
   End
   Begin VB.CommandButton cmdGroupHelp 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   270
      Left            =   3570
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   1335
      Width           =   345
   End
   Begin VB.TextBox txtSpcGroupNm 
      BackColor       =   &H00F1F5F4&
      Height          =   270
      Left            =   6630
      TabIndex        =   1
      Top             =   690
      Width           =   1545
   End
   Begin VB.TextBox txtWACd 
      BackColor       =   &H00F1F5F4&
      Height          =   270
      Left            =   3900
      TabIndex        =   4
      Top             =   1012
      Width           =   1275
   End
   Begin VB.TextBox txtStart 
      BackColor       =   &H00F1F5F4&
      Height          =   270
      Left            =   6630
      TabIndex        =   5
      Top             =   1012
      Width           =   1275
   End
   Begin VB.TextBox txtEnd 
      BackColor       =   &H00F1F5F4&
      Height          =   300
      Left            =   8895
      TabIndex        =   6
      Top             =   990
      Width           =   1275
   End
   Begin VB.TextBox txtGroupCd 
      BackColor       =   &H00F1F5F4&
      Height          =   270
      Left            =   3900
      TabIndex        =   9
      Top             =   1335
      Width           =   1275
   End
   Begin VB.TextBox txtReportSeq 
      BackColor       =   &H00F1F5F4&
      Height          =   270
      Left            =   6630
      TabIndex        =   10
      Top             =   1335
      Width           =   1275
   End
   Begin VB.TextBox txtSpcGroupCd 
      BackColor       =   &H00F1F5F4&
      Height          =   270
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   0
      Top             =   690
      Width           =   1305
   End
   Begin VB.ListBox lstSpcGroup 
      BackColor       =   &H00F5FFF4&
      Height          =   7440
      Left            =   90
      TabIndex        =   18
      Top             =   780
      Width           =   2235
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   690
      Left            =   30
      TabIndex        =   15
      Top             =   -60
      Width           =   10800
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "미생물 검체군 등록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   420
         TabIndex        =   17
         Top             =   240
         Width           =   3150
      End
      Begin VB.Label Label6 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "( 여기서 등록된 검체군 정보를 근거로 WorkSheet을 작성하게 되므로 배지코드와 WorkArea코드는 미리 등록되어 있어야 합니다. )"
         Height          =   375
         Left            =   4620
         TabIndex        =   16
         Top             =   210
         Width           =   5625
      End
   End
   Begin VB.Label lblMedia 
      BackColor       =   &H00D1D8D3&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   405
      Left            =   2880
      TabIndex        =   32
      Top             =   2220
      Width           =   7935
   End
   Begin VB.Label lblSpccd 
      BackColor       =   &H00D1D8D3&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   435
      Left            =   2880
      TabIndex        =   31
      Top             =   1710
      Width           =   7935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "배지"
      Height          =   225
      Left            =   2400
      TabIndex        =   30
      Top             =   2220
      Width           =   435
   End
   Begin VB.Label Label10 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검체"
      Height          =   225
      Left            =   2400
      TabIndex        =   27
      Top             =   1710
      Width           =   435
   End
   Begin VB.Label Label9 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출력순서"
      Height          =   225
      Left            =   5580
      TabIndex        =   25
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label Label8 
      BackColor       =   &H00DBE6E6&
      Caption         =   "끝"
      Height          =   225
      Left            =   8565
      TabIndex        =   24
      Top             =   1050
      Width           =   345
   End
   Begin VB.Label Label7 
      BackColor       =   &H00DBE6E6&
      Caption         =   "WorkArea"
      Height          =   225
      Left            =   2400
      TabIndex        =   23
      Top             =   1050
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00DBE6E6&
      Caption         =   "시작"
      Height          =   225
      Left            =   5580
      TabIndex        =   22
      Top             =   1050
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Group설정"
      Height          =   225
      Left            =   2400
      TabIndex        =   21
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검체군명"
      Height          =   225
      Left            =   5580
      TabIndex        =   20
      Top             =   735
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검체군코드"
      Height          =   225
      Left            =   2400
      TabIndex        =   19
      Top             =   735
      Width           =   1155
   End
End
Attribute VB_Name = "frm365SpcGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private flgSpdClick  As Boolean
Private objSql As New clsLISSqlStatement

Private Sub cmdClear_Click()
    Call clearAllInfo
    Call clearlstSpcGroup
    txtSpcGroupCd.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer
    
    If Trim(txtSpcGroupCd.Text) = "" Then Exit Sub

    sMsg = txtSpcGroupNm.Text & " 에 관한 정보를 모두 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteSpcGroup
        Call ReLoadForm
'        medMain.stsBar.Panels(2).Text = "정상적으로 삭제 처리 되었습니다. 다음 작업을 처리하세요"
    Else
        Exit Sub
    End If
    
End Sub
Private Sub DeleteSpcGroup()
    
    Dim sSqlDelC217 As String
    Dim sSqlDelC106 As String
    Dim sSqlDelC215 As String
       
    sSqlDelC217 = objSql.SqlDeleteLAB032(LC3_SGroup, Trim(txtSpcGroupCd.Text))
    sSqlDelC106 = objSql.SqlDeleteLAB032(LC2_SpcMedia, Trim(txtSpcGroupCd.Text))
    sSqlDelC215 = objSql.SqlDeleteLAB032(LC3_Specimen, Trim(txtSpcGroupCd.Text))
                  
On Error GoTo DBExecError

    dbconn.BeginTrans
    
    dbconn.Execute (sSqlDelC217) ' 검체군정보삭제
    dbconn.Execute (sSqlDelC106) ' 배지삭제
    dbconn.Execute (sSqlDelC215) ' 검체 - 검체군삭제
    
    dbconn.CommitTrans
    Exit Sub

DBExecError:
    dbconn.RollbackTrans

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGroupHelp_Click()
    lstGroup.Visible = True
    Call locateLst(lstGroup, cmdGroupHelp)
    Call LoadLstGroup
End Sub

Private Sub cmdSave_Click()
    Dim sSqlDelC217 As String
    Dim sSqlDelC106 As String
    Dim sSqlDelC215 As String
    Dim sWcGs As String
    
    Dim sSqlInC217 As String    ' 검체군코드, 검체군명, Workarea, 시작, 끝,
                                ' 그룹설정, 출력순서
    Dim sField2Data  As String
    Dim i%
    
    
    If txtSpcGroupCd.Text = "" Then Exit Sub
    
    sSqlDelC217 = objSql.SqlDeleteLAB032(LC3_SGroup, Trim(txtSpcGroupCd.Text))
    sSqlDelC106 = objSql.SqlDeleteLAB031(LC2_SpcMedia, Trim(txtSpcGroupCd.Text))
                  
                  
    sWcGs = Choose(chkWcGs.Value + 1, "", "W")
    
    sField2Data = txtWACd.Text & ";" & txtStart.Text & ";" & txtEnd.Text
    
'2001'04/02 수정
    sSqlInC217 = objSql.SqlSaveLAB032(LC3_SGroup, Trim(txtSpcGroupCd.Text), Trim(txtSpcGroupNm.Text), _
                                    sField2Data, Trim(txtReportSeq.Text), Trim(txtGroupCd.Text), _
                                    IIf(optOrderBy(0).Value, "0", "1"), "", "", 1)
                            

On Error GoTo DBExecError

    dbconn.BeginTrans
    
    dbconn.Execute (sSqlDelC217)
    dbconn.Execute (sSqlInC217)

    dbconn.Execute (sSqlDelC106) ' 배지삭제
    Call insertMedia             ' 배지삽입
    Call UpdateSpccd             ' 검체삽입
    
    dbconn.CommitTrans
    Call ReLoadForm
    
    Exit Sub

DBExecError:
    dbconn.RollbackTrans
     
End Sub

Private Sub ReLoadForm()
    Call clearAllInfo
    lstSpcGroup.Clear
    Call LoadSpcGroup
    Call LoadTotalspdSpccd
    Call LoadTotalspdMedia
    lblSpcCd.Caption = ""
    lblMedia.Caption = ""
    Call clearlstSpcGroup

End Sub

Private Sub insertMedia()
    Dim i%, sMediaCd$, Oldcol%, iCol%, j%
    Dim sSqlInC106 As String    ' 검체군-배지 정보
    
    With spdMedia
        For i = 2 To .MaxCols Step 2
            .Col = i
            For j = 1 To 20
                .Row = j
                If (.Value = True) And (.TypeCheckText <> "") Then
                    .Col = .Col - 1  ' 배지코드
                    sMediaCd = .TypeCheckText
                    sSqlInC106 = objSql.SqlSaveLAB031(LC2_SpcMedia, Trim(txtSpcGroupCd.Text), _
                                                      sMediaCd, "", "", "", "", "", "", "", 1)
                    dbconn.Execute (sSqlInC106)
                    .Col = i
                End If
            Next j
        Next i
    End With

End Sub

Private Sub UpdateSpccd()
    Dim i%, sSpccdCd$, j%
    Dim sSqlUpdateC215 As String    ' 검체-검체군 정보
        
    
    With spdSpccd
        For i = 2 To .MaxCols Step 2
            .Col = i
            For j = 1 To 20
                .Row = j
                
                If (.Value = True) And (.TypeCheckText <> "") Then
                    .Col = .Col - 1
                    sSpccdCd = .TypeCheckText
                    sSqlUpdateC215 = objSql.SqlSaveSpcGrp(sSpccdCd, Trim(txtSpcGroupCd.Text))
                    dbconn.Execute (sSqlUpdateC215)
                    .Col = i
                ElseIf (.Value = False) And (.TypeCheckText <> "") Then
                    .Col = .Col - 1
                    sSpccdCd = .TypeCheckText
                    If chkExistSpcGroup(sSpccdCd) = False Then    ' Case Not exist 검체군코드
                        sSqlUpdateC215 = objSql.SqlSaveSpcGrp(sSpccdCd, "Null")
                        dbconn.Execute (sSqlUpdateC215)
                    End If
                    .Col = i
                End If
            Next j
        Next i
    End With

End Sub

Private Function chkExistSpcGroup(sSpccdCd As String) As Boolean
    Dim sSqlGetSpcGroupCd As String
    Dim rsGetSpcGroupCd As Recordset
    
    sSqlGetSpcGroupCd = objSql.SqlLAB032CodeList(LC3_Specimen, "field2", sSpccdCd)
                        
    Set rsGetSpcGroupCd = New Recordset
    rsGetSpcGroupCd.Open sSqlGetSpcGroupCd, dbconn
    
'이전의 검체군코드값이 Null, "" , 또는 현재 자신의 검체군 코드이면
'Update가 가능하다.
    If "" & rsGetSpcGroupCd.Fields("field2").Value = Null Or _
       "" & rsGetSpcGroupCd.Fields("field2").Value = "" Or _
       "" & rsGetSpcGroupCd.Fields("field2").Value = Trim(txtSpcGroupCd.Text) Then
        
        chkExistSpcGroup = False
        
'이전의 검체군코드값이 Null, ""가 아니며 현재 자신의 검체군 코드도
'아니면 다른 검체군의 검체로 지정된것이므로 Update 불가
    Else
        chkExistSpcGroup = True
    End If
    Set rsGetSpcGroupCd = Nothing
    
End Function

Private Sub cmdWAhelp_Click()
    lstWA.Visible = True
    Call locateLst(lstWA, cmdWAhelp)
    Call LoadLstWA
    
End Sub

Private Sub locateLst(lstbox As ListBox, BaseCtrl As Control)
    lstbox.Top = BaseCtrl.Top + BaseCtrl.Height
    lstbox.Left = BaseCtrl.Left
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        lstWA.Visible = False
        lstGroup.Visible = False
    End If
End Sub

Private Sub Form_Load()
'    SetPosition 2, Me
    Call LoadSpcGroup
    Call LoadTotalspdSpccd
    Call LoadTotalspdMedia
End Sub

Private Sub LoadTotalspdMedia()
    Dim sSqlGetTotalMedia As String
    Dim rsTotalMedia As Recordset
    Dim i%, mok%
    
    sSqlGetTotalMedia = objSql.SqlLAB032CodeList(LC3_Media, "cdval1 as MediaCd, text1 as MediaNm")
                        
    Set rsTotalMedia = New Recordset
    rsTotalMedia.Open sSqlGetTotalMedia, dbconn
    
    With spdMedia

        If rsTotalMedia.RecordCount > 100 Then
            .MaxCols = (rsTotalMedia.RecordCount \ 20) + 2
            .ColWidth(.MaxCols - 1) = 7
            .Col = .MaxCols - 1: .ColHidden = True
            .ColWidth(.MaxCols) = 14
        End If
        
        
        .Row = 0
        
        For i = 1 To rsTotalMedia.RecordCount
            mok = (i \ 20)          ' 몫
           
            .Col = mok + (mok + 2) - 1  '코드값 col
            If (i Mod 20) = 0 Then
                .Col = .Col - 2
                .Row = 20
            Else
                .Row = i Mod 20     ' 나머지
            End If

            .TypeCheckText = "" & rsTotalMedia.Fields("MediaCd").Value
            .Col = mok + (mok + 2)
            If (i Mod 20) = 0 Then
                .Col = .Col - 2
                .Row = 20
            Else
                .Row = i Mod 20
            End If
            .TypeCheckText = "" & rsTotalMedia.Fields("MediaNm").Value
            
            rsTotalMedia.MoveNext
        Next i
    End With
    Set rsTotalMedia = Nothing
                        
End Sub

Private Sub LoadTotalspdSpccd()
    Dim sSqlTotalSpccd As String
    Dim rsTotalSpccd As Recordset
    Dim i%, mok%
    
    sSqlTotalSpccd = objSql.SqlLAB032CodeList(LC3_Specimen, "cdval1 as spccdCd, field4 as spccdNm")
    
    Set rsTotalSpccd = New Recordset
    rsTotalSpccd.Open sSqlTotalSpccd, dbconn
    
    
    With spdSpccd
        
        If rsTotalSpccd.RecordCount > 100 Then
            .MaxCols = (rsTotalSpccd.RecordCount \ 20 + 1) * 2
            '.ColWidth(.MaxCols - 1) = 7
            '.Col = .MaxCols - 1: .ColHidden = True      ' Code col
            '.ColWidth(.MaxCols) = 14                    ' CodeName Col
        End If
        
        .Row = 0
        
        For i = 1 To rsTotalSpccd.RecordCount
            mok = (i \ 20)          ' 몫
            
            .Col = mok + (mok + 2) - 1      ' 코드값 Col
            If (i Mod 20) = 0 Then
                .Col = .Col - 2
                .Row = 20
            Else
                .Row = i Mod 20     ' 나머지
            End If
            
            .TypeCheckText = "" & rsTotalSpccd.Fields("spccdCd").Value
            .ColHidden = True
            
            .Col = mok + (mok + 2)          ' 코드이름 Col
            If (i Mod 20) = 0 Then
                .Col = .Col - 2
                .Row = 20
            Else
                .Row = i Mod 20     ' 나머지
            End If
            
            .TypeCheckText = "" & rsTotalSpccd.Fields("spccdNm").Value
            .ColHidden = False
            .ColWidth(.Col) = 14
            
            rsTotalSpccd.MoveNext
        Next i
    End With
            
    Set rsTotalSpccd = Nothing
            
            
End Sub

Private Sub LoadLstGroup()
    Dim sSqlGetGroup As String
    Dim rsGetGroup As Recordset
    Dim i%
    
    sSqlGetGroup = objSql.SqlLAB032CodeList(LC3_MWSKinds, "cdval1 as GroupCd, field1 as GroupNm ")
                    
    Set rsGetGroup = New Recordset
    rsGetGroup.Open sSqlGetGroup, dbconn
    
    If rsGetGroup.EOF = True Then Exit Sub
    
    lstGroup.Clear
    For i = 1 To rsGetGroup.RecordCount
        lstGroup.AddItem "" & rsGetGroup.Fields("GroupCd").Value & vbTab & _
                         "" & rsGetGroup.Fields("GroupNm").Value
        rsGetGroup.MoveNext
    Next i
    
    Set rsGetGroup = Nothing
End Sub

Private Sub LoadSpcGroup()

    Dim sSqlGetSpcGroup As String
    Dim rsGetSpcGroup As Recordset
    Dim i%
    
    sSqlGetSpcGroup = objSql.SqlLAB032CodeList(LC3_SGroup, "cdval1 as spcGroupCd, field1 as spcGroupNm")
                      
    Set rsGetSpcGroup = New Recordset
    rsGetSpcGroup.Open sSqlGetSpcGroup, dbconn
    
    For i = 1 To rsGetSpcGroup.RecordCount
    
        lstSpcGroup.AddItem "" & rsGetSpcGroup.Fields("spcGroupCd").Value & vbTab & _
                            "" & rsGetSpcGroup.Fields("spcGroupNm").Value
        rsGetSpcGroup.MoveNext
    Next i
    
    Set rsGetSpcGroup = Nothing
    
End Sub

Private Sub LoadLstWA()
    Dim sSqlGetWorkarea As String
    Dim rsGetWorkarea As Recordset
    Dim i%
    
    sSqlGetWorkarea = objSql.SqlLAB032CodeList(LC3_WorkArea, "cdval1 as WACd, field1 as WANm")
    
    Set rsGetWorkarea = New Recordset
    rsGetWorkarea.Open sSqlGetWorkarea, dbconn
    
    If rsGetWorkarea.EOF = True Then Exit Sub
    
    lstWA.Clear
    For i = 1 To rsGetWorkarea.RecordCount
        lstWA.AddItem "" & rsGetWorkarea.Fields("WACd").Value & vbTab & _
                      "" & rsGetWorkarea.Fields("WANm").Value
        rsGetWorkarea.MoveNext
    Next i
    
    Set rsGetWorkarea = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub

Private Sub lstGroup_Click()
    txtGroupCd.Text = Trim(Mid(lstGroup.Text, 1, _
                 InStr(1, lstGroup.Text, vbTab) - 1))
    lstGroup.Visible = False
End Sub

Private Sub lstSpcGroup_Click()
    Dim sSpcGroupCd As String
    
    If lstSpcGroup.Text = "" Then Exit Sub
    
    sSpcGroupCd = Trim(Mid(lstSpcGroup.Text, 1, _
                 InStr(1, lstSpcGroup.Text, vbTab) - 1))
    
    
    Call clearAllInfo
    Call DspSpcGroupInfo(sSpcGroupCd)
    
End Sub

Private Sub clearAllInfo()
    txtSpcGroupCd.Text = ""
    txtSpcGroupNm.Text = ""
    txtWACd.Text = ""
    txtStart.Text = ""
    txtEnd.Text = ""
    txtGroupCd.Text = ""
    txtReportSeq.Text = ""
    lblSpcCd.Caption = ""
    lblMedia.Caption = ""
    chkWcGs.Value = 0
    Call clearspdSpccd
    Call clearspdMedia
End Sub
Private Sub clearlstSpcGroup()
    Dim i%
    
    For i = 0 To lstSpcGroup.ListCount - 1
        lstSpcGroup.Selected(i) = False
    Next i
End Sub
Private Sub clearspdSpccd()
    Dim i%, j%, iCol%
    
    With spdSpccd
        .Row = 0
        iCol = 2
        .Col = iCol
        Do
            For j = 1 To .MaxRows
                .Row = j
                .Value = False
            Next j
            iCol = iCol + 2
            .Col = iCol
            If iCol > .MaxCols Then Exit Do
        Loop
                
    End With
           
End Sub
Private Sub clearspdMedia()
    Dim i%, j%, iCol%
    
    With spdMedia
        .Row = 0
        iCol = 2
        .Col = iCol
        Do
            For j = 1 To .MaxRows
                .Row = j
                .Value = False
            Next j
            iCol = iCol + 2
            .Col = iCol
            If iCol > .MaxCols Then Exit Do
        Loop

    End With
           
End Sub

Private Sub DspSpcGroupInfo(sSpcGroupCd As String)
    Dim sSqlGetC217 As String
    Dim rsGetC217 As Recordset
    
    sSqlGetC217 = objSql.SqlLAB032CodeList(LC3_SGroup, "field1 as SpcGroupNm, field2 as WAStEn , " & _
                                            "field3 as ReportSeq , field4 as GroupCd, field5 as RptSeq ", _
                                            sSpcGroupCd)
                  
    Set rsGetC217 = New Recordset
    rsGetC217.Open sSqlGetC217, dbconn
    
    If rsGetC217.EOF = True Then Exit Sub
    
    txtSpcGroupCd.Text = sSpcGroupCd
    txtSpcGroupNm.Text = "" & rsGetC217.Fields("SpcGroupNm").Value
    
    
    txtWACd.Text = medGetP("" & rsGetC217.Fields("WAStEn").Value, 1, ";")
    txtStart.Text = medGetP("" & rsGetC217.Fields("WAStEn").Value, 2, ";")
    txtEnd.Text = medGetP("" & rsGetC217.Fields("WAStEn").Value, 3, ";")
    
    txtGroupCd.Text = "" & rsGetC217.Fields("GroupCd").Value
    txtReportSeq.Text = "" & rsGetC217.Fields("ReportSeq").Value
    
    optOrderBy(Val("" & rsGetC217.Fields("RptSeq").Value)).Value = True
    
'    If "" & rsGetC217.Fields("WcGs").Value = "W" Then
'        chkWcGs.Value = 1
'    Else
'        chkWcGs.Value = 0
'    End If
    
    Call DspCheckedSpccd(sSpcGroupCd)
    Call DspCheckedMedia(sSpcGroupCd)
    
    Set rsGetC217 = Nothing
End Sub

Private Sub DspCheckedSpccd(sSpcGroupCd As String)
    Dim sSqlGetCheckedSpccd As String
    Dim rsCheckedSpccd As Recordset
    Dim i%
    
    sSqlGetCheckedSpccd = objSql.SqlLAB032CodeList(LC3_Specimen, "cdval1 as SpccdCd,  field4 as spccdNm", _
                                                    , , " and " & DBW("field2=", Trim(sSpcGroupCd)))
    Set rsCheckedSpccd = New Recordset
    rsCheckedSpccd.Open sSqlGetCheckedSpccd, dbconn
    
    If rsCheckedSpccd.EOF = True Then
        Exit Sub
    End If
        
    For i = 1 To rsCheckedSpccd.RecordCount
        Call FindAndCheckSpccd("" & rsCheckedSpccd.Fields("SpccdCd").Value)
        Call DspLblSpccd("" & rsCheckedSpccd.Fields("SpccdCd").Value)
        rsCheckedSpccd.MoveNext
    Next i
            
    Set rsCheckedSpccd = Nothing
End Sub

Private Sub DspLblSpccd(sCheckedSpccdCd As String)
    Dim iStartPos%, iCommaPos%
    Dim sSpccdCdOfLable As String
    Dim cComma As String
    Dim tmpLbl As String
    Dim foreStr As String
    Dim backStr As String
    
    cComma = ","
    If Len(lblSpcCd.Caption) = 0 Then       ' 처음입력되는 검체이름일경우
        lblSpcCd.Caption = sCheckedSpccdCd
        Exit Sub
    Else                                    ' 두번째 이상의 입력일 경우
        iStartPos = 1
        Do
            iCommaPos = InStr(iStartPos, lblSpcCd.Caption, cComma, vbTextCompare)
            If iCommaPos <> 0 And iStartPos = 1 Then     ' 첫번째 검체에 대해
                sSpccdCdOfLable = Mid(lblSpcCd.Caption, iStartPos, iCommaPos - iStartPos)
                If Trim(sSpccdCdOfLable) = Trim(sCheckedSpccdCd) Then
                    tmpLbl = lblSpcCd.Caption
                    lblSpcCd.Caption = Mid(tmpLbl, iCommaPos + 2, Len(lblSpcCd.Caption))
                    Exit Sub
                End If
            ElseIf iCommaPos = 0 And iStartPos = 1 Then  ' 두번째 입력일경우
                sSpccdCdOfLable = Mid(lblSpcCd.Caption, iStartPos, Len(lblSpcCd.Caption))
                If Trim(sSpccdCdOfLable) = Trim(sCheckedSpccdCd) Then
                    tmpLbl = lblSpcCd.Caption
                    lblSpcCd.Caption = ""
                    Exit Sub
                Else
                    lblSpcCd.Caption = lblSpcCd.Caption & ", " & sCheckedSpccdCd
                    Exit Sub
                End If
            ElseIf iCommaPos = 0 And iStartPos <> 1 Then ' Lable의 마지막 검체일 경우
                sSpccdCdOfLable = Mid(lblSpcCd.Caption, iStartPos, Len(lblSpcCd.Caption))
                If Trim(sSpccdCdOfLable) = (sCheckedSpccdCd) Then
                    tmpLbl = lblSpcCd.Caption
                    lblSpcCd.Caption = Mid(tmpLbl, 1, iStartPos - 3)
                    Exit Sub
                Else
                    lblSpcCd.Caption = lblSpcCd.Caption & ", " & sCheckedSpccdCd
                    Exit Sub
                End If
            Else                                        'lable의 중간검체일 경우
                sSpccdCdOfLable = Mid(lblSpcCd.Caption, iStartPos, iCommaPos - iStartPos)
                If Trim(sSpccdCdOfLable) = Trim(sCheckedSpccdCd) Then
                    tmpLbl = lblSpcCd.Caption
                    foreStr = Mid(tmpLbl, 1, iStartPos - 3)
                    backStr = Mid(tmpLbl, iCommaPos, Len(lblSpcCd.Caption))
                    lblSpcCd.Caption = foreStr & backStr
                    Exit Sub
                End If
            End If
            
            iStartPos = iCommaPos + 2
        Loop
        lblSpcCd.Caption = lblSpcCd.Caption & ", " & sCheckedSpccdCd
        
    End If
End Sub

Private Sub DspCheckedMedia(sSpcGroupCd As String)
    Dim sSqlGetCheckedMedia As String
    Dim rsCheckedMedia As Recordset
    Dim i%
    Dim objWsSql As New clsLISSqlMasters
    
    Set rsCheckedMedia = New Recordset
    rsCheckedMedia.Open objWsSql.SqlGetCheckedMedia(sSpcGroupCd), dbconn
    
    If rsCheckedMedia.RecordCount < 1 Then
        Exit Sub
    End If

    For i = 1 To rsCheckedMedia.RecordCount
        Call FindAndCheckMedia("" & rsCheckedMedia.Fields("MediaCd").Value)
        Call DspLblMedia("" & rsCheckedMedia.Fields("MediaCd").Value)
        rsCheckedMedia.MoveNext
    Next i
    
    Set rsCheckedMedia = Nothing
    Set objWsSql = Nothing
End Sub

Private Sub DspLblMedia(sCheckedMediaCd As String)
    
    
    Dim iStartPos%, iCommaPos%
    Dim sMediaCdOfLable As String
    Dim cComma As String
    Dim tmpLbl As String
    Dim foreStr As String
    Dim backStr As String
    
    cComma = ","
    If Len(lblMedia.Caption) = 0 Then       ' 처음입력되는 검체이름일경우
        lblMedia.Caption = sCheckedMediaCd
        Exit Sub
    Else                                    ' 두번째 이상의 입력일 경우
        iStartPos = 1
        Do
            iCommaPos = InStr(iStartPos, lblMedia.Caption, cComma, vbTextCompare)
            If iCommaPos <> 0 And iStartPos = 1 Then     '첫번째 검체의 경우
                sMediaCdOfLable = Mid(lblMedia.Caption, iStartPos, iCommaPos - iStartPos)
                If Trim(sMediaCdOfLable) = Trim(sCheckedMediaCd) Then
                    tmpLbl = lblMedia.Caption
                    lblMedia.Caption = Mid(tmpLbl, iCommaPos + 2, Len(lblMedia.Caption))
                    Exit Sub
                End If
            
            ElseIf iCommaPos = 0 And iStartPos = 1 Then  ' 두번째 입력일경우
                sMediaCdOfLable = Mid(lblMedia.Caption, iStartPos, Len(lblMedia.Caption))
                If Trim(sMediaCdOfLable) = Trim(sCheckedMediaCd) Then
                    tmpLbl = lblMedia.Caption
                    lblMedia.Caption = ""
                    Exit Sub
                Else
                    lblMedia.Caption = lblMedia.Caption & ", " & sCheckedMediaCd
                    Exit Sub
                End If
            ElseIf iCommaPos = 0 And iStartPos <> 1 Then ' Lable의 마지막 검체일 경우
                sMediaCdOfLable = Mid(lblMedia.Caption, iStartPos, Len(lblMedia.Caption))
                If Trim(sMediaCdOfLable) = (sCheckedMediaCd) Then
                    tmpLbl = lblMedia.Caption
                    lblMedia.Caption = Mid(tmpLbl, 1, iStartPos - 3)
                    Exit Sub
                Else
                    lblMedia.Caption = lblMedia.Caption & ", " & sCheckedMediaCd
                    Exit Sub
                End If
            Else                                        'lable의 중간검체일 경우
                sMediaCdOfLable = Mid(lblMedia.Caption, iStartPos, iCommaPos - iStartPos)
                If Trim(sMediaCdOfLable) = Trim(sCheckedMediaCd) Then
                    tmpLbl = lblMedia.Caption
                    foreStr = Mid(tmpLbl, 1, iStartPos - 3)
                    backStr = Mid(tmpLbl, iCommaPos, Len(lblMedia.Caption))
                    lblMedia.Caption = foreStr & backStr
                    Exit Sub
                End If
            End If
            
            iStartPos = iCommaPos + 2
        Loop
        lblMedia.Caption = lblMedia.Caption & ", " & sCheckedMediaCd
        
    End If
End Sub
Private Sub FindAndCheckMedia(sMediaCd As String)
    Dim i%, j%
    With spdMedia
        For i = 1 To .MaxCols Step 2
            .Col = i
            
            For j = 1 To .MaxRows
                .Row = j
                If .TypeCheckText = sMediaCd Then
                    .Col = i + 1
                    .Value = True
                End If
            Next j
        Next i
    End With

End Sub
Private Sub FindAndCheckSpccd(sSpccdCd As String)
    Dim i%, j%
    With spdSpccd
        For i = 1 To .MaxCols Step 2
            .Col = i
            
            For j = 1 To .MaxRows
                .Row = j
                If .TypeCheckText = sSpccdCd Then
                    .Col = i + 1
                    .Value = True
                End If
            Next j
        Next i
    End With
End Sub
Private Sub lstWA_Click()
    txtWACd.Text = Trim(Mid(lstWA.Text, 1, _
                 InStr(1, lstWA.Text, vbTab) - 1))
    lstWA.Visible = False
End Sub

Private Sub spdMedia_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If flgSpdClick = True Then
        With spdMedia
            .Col = Col - 1
            .Row = Row
            Call DspLblMedia(.TypeCheckText)
        End With
    End If
End Sub

Private Sub spdMedia_Click(ByVal Col As Long, ByVal Row As Long)
    flgSpdClick = True
End Sub

Private Sub spdMedia_LostFocus()
    flgSpdClick = False
End Sub

Private Sub spdSpccd_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If flgSpdClick = True Then
        With spdSpccd
            .Col = Col - 1
            .Row = Row
            Call DspLblSpccd(.TypeCheckText)
        End With
    End If
End Sub

Private Sub spdSpccd_Click(ByVal Col As Long, ByVal Row As Long)
    flgSpdClick = True
End Sub

Private Sub spdSpccd_LostFocus()
    flgSpdClick = False
End Sub
