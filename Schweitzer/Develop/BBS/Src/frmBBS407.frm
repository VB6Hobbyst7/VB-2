VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBBS407 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'Å©±â °íÁ¤ ´ëÈ­ »óÀÚ
   Caption         =   "DM ¹ß¼Û"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmBBS407.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin TabDlg.SSTab SSTab1 
      Height          =   1320
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2328
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   14411494
      TabCaption(0)   =   "ÇåÇ÷ÀÏÀÚº°"
      TabPicture(0)   =   "frmBBS407.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÇåÇ÷Áõ¼­º°"
      TabPicture(1)   =   "frmBBS407.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   975
         Left            =   -75000
         TabIndex        =   8
         Top             =   320
         Width           =   5235
         Begin VB.TextBox Text2 
            Appearance      =   0  'Æò¸é
            Height          =   285
            Left            =   3300
            TabIndex        =   10
            Top             =   420
            Width           =   1290
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Æò¸é
            Height          =   285
            Left            =   1620
            TabIndex        =   9
            Top             =   420
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "~"
            Height          =   180
            Left            =   3060
            TabIndex        =   12
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Áõ¼­ ¹øÈ£ :"
            Height          =   180
            Left            =   600
            TabIndex        =   11
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   975
         Left            =   0
         TabIndex        =   3
         Top             =   320
         Width           =   5235
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3300
            TabIndex        =   4
            Top             =   420
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   67043331
            CurrentDate     =   36803
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Top             =   420
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   67043331
            CurrentDate     =   36803
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "~"
            Height          =   180
            Left            =   3060
            TabIndex        =   7
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Çå  Ç÷  ÀÏ :"
            Height          =   180
            Left            =   600
            TabIndex        =   6
            Top             =   480
            Width           =   900
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Á¾·á(&X)"
      Height          =   510
      Left            =   4095
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   1
      Tag             =   "128"
      Top             =   1740
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Ãâ·Â(&P)"
      Height          =   510
      Left            =   2775
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   1740
      Width           =   1320
   End
End
Attribute VB_Name = "frmBBS407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub


Private Sub DM_Print()

End Sub

'Private Sub cmdQuery_Click()
'    Dim objGetSql As New clsGetSqlStatement
'    Dim objProBar As clsprogress
'    Dim Rs        As New RECORDSET
'    Dim strTmp    As String
'    Dim FrDt      As String
'    Dim ToDt      As String
'    Dim ii        As Integer
'    Dim strBldno  As String
'    Dim ptid      As String
'
''    objGetSql.setDbConn DBConn
'
'    FrDt = Format(dtpFrDt.value, PRESENTDATE_FORMAT)
'    ToDt = Format(dtpToDt.value, PRESENTDATE_FORMAT)
'    Set Rs = objGetSql.Get_DonorQuery(FrDt, ToDt)
'
'    If Rs.RecordCount > 0 Then
'        Set objProBar = New clsprogress
'        Set objProBar.StatusBar = medMain.stsBar
'        objProBar.Max = Rs.RecordCount
'        With tblList
'            ii = 1
'            .MaxRows = 0
'            .MaxRows = Rs.RecordCount
'            .ReDraw = False
'            Rs.MoveFirst
'            Do Until Rs.EOF
'                If chkAll.value = 0 Then
'                    If Not optDonorCd(Val("" & Rs.Fields("donorcd").value)).value Then
'                        .MaxRows = .MaxRows - 1
'                        GoTo Skip
'                    End If
'                End If
'                .Row = ii
'                If strTmp <> Rs.Fields("donorid").value Then
'                    .Col = TblColumn.TcName: .value = Rs.Fields("donornm").value
'                    .Col = TblColumn.tcDOB: .value = Format(Rs.Fields("dob").value, "####-##-##")
'                    .Col = TblColumn.TcSEXAGE: .value = Rs.Fields("sex").value & "/"
'                                                If Trim(Rs.Fields("dob").value & "") <> "" Then
'                                                    .value = .value & medFindAge(Rs.Fields("dob").value, "Y")
'                                                End If
'                    .Col = TblColumn.TcABO: .value = Rs.Fields("abo").value & Rs.Fields("rh").value
'                End If
'                .Col = TblColumn.tcACCDT: .value = Format(Rs.Fields("donoraccdt").value, "####/##/##")
'                .Col = TblColumn.TcTMPID: .value = Rs.Fields("tmpid").value
'                .Col = TblColumn.TcDONORTYPE
'                    Select Case Rs.Fields("donorcd").value
'                        Case "0": .value = "ÀÓÀÇÇåÇ÷"
'                        Case "1": .value = "ÁöÁ¤ÇåÇ÷"
'                        Case "2": .value = "Autologos"
'                        Case "3": .value = "Pheresis"
'                        Case "4": .value = "Phlebotomy"
'                    End Select
'                .Col = TblColumn.TcSELPTID
'                If Rs.Fields("reservedid").value <> "" And Rs.Fields("reservedid").value <> "0" Then
'                    ptid = getptnm(Rs.Fields("reservedid").value)
'                    If ptid <> "" Then
'                        .value = ptid & "(" & Rs.Fields("reservedid").value & ")"
'                    End If
'                Else
'                    .value = ""
'                End If
'                strBldno = Rs.Fields("bldsrc").value & "-" & Rs.Fields("bldyy").value & "-" & Format(Rs.Fields("bldno").value, "000000")
'                If strBldno = "--000000" Then strBldno = ""
'                .Col = TblColumn.TcBLOODNO: .value = strBldno
'                .Col = TblColumn.TcVOLUMN: .value = Rs.Fields("volumn").value
'                .Col = TblColumn.TcACCVAL: .value = IIf(Rs.Fields("okdiv1").value = "1", "Ok", IIf(Rs.Fields("okdiv1").value = "0", "Not", "")): .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
'                .Col = TblColumn.TcRMKVAL: .value = IIf(Rs.Fields("okdiv2").value = "1", "Ok", IIf(Rs.Fields("okdiv1").value = "0", "Not", "")):  .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
'                .Col = TblColumn.TcTESTVAL: .value = IIf(Rs.Fields("okdiv3").value = "1", "Ok", IIf(Rs.Fields("okdiv1").value = "0", "Not", "")): .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
'                .Col = TblColumn.TcCANCEL: .value = IIf(Rs.Fields("cancelfg").value = "1", "Y", "")
'                .Col = TblColumn.tcDONORID: .value = Rs.Fields("donorid").value
'                strTmp = .value
'                objProBar.value = ii
'                ii = ii + 1
'Skip:
'                Rs.MoveNext
'            Loop
'            .ReDraw = True
'        End With
'
'    End If
'
'    Set Rs = Nothing
'    Set objGetSql = Nothing
'End Sub
'

