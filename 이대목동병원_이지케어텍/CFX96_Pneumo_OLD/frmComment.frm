VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmComment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "코멘트설정"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11775
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11775
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9090
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10260
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6105
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   11385
      _Version        =   393216
      _ExtentX        =   20082
      _ExtentY        =   10769
      _StockProps     =   64
      ColsFrozen      =   2
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   4
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmComment.frx":014A
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdConfirm_Click()
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
    
'        Call WritePrivateProfileString("HOSP", "APIURL", txtAPIURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("HOSP", "STDURL", txtSTDURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("HOSP", "EDUURL", txtEDUURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("HOSP", "OPRURL", txtOPRURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
'
'        Call WritePrivateProfileString("COMM", "ORDPATH", txtOrderPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMM", "RSTPATH", txtResultPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\INI\" & gMACH & ".ini")
                
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("경로가 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    With spdResult
        If gCmnt.BPCmnt <> "" Then
            .Row = 1
            .Col = 1:  .Text = "BP+"
            .Col = 2:  .Text = gCmnt.BPCmnt
            
            .Row = 2
            .Col = 1:  .Text = "CP+"
            .Col = 2:  .Text = gCmnt.CPCmnt
            
            .Row = 3
            .Col = 1:  .Text = "LP+"
            .Col = 2:  .Text = gCmnt.LPCmnt
            
            .Row = 4
            .Col = 1:  .Text = "MP+"
            .Col = 2:  .Text = gCmnt.MPCmnt
        End If
    End With
    
End Sub

Private Sub optAPIURL_Click(Index As Integer)
    
'    Select Case Index
'        Case 0:     txtAPIURL.Text = txtSTDURL.Text
'        Case 1:     txtAPIURL.Text = txtEDUURL.Text
'        Case 2:     txtAPIURL.Text = txtOPRURL.Text
'    End Select
    
End Sub

