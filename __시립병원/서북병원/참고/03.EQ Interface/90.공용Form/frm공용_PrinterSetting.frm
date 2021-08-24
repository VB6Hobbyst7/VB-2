VERSION 5.00
Begin VB.Form frm공용_PrinterSetting 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Printer Setting"
   ClientHeight    =   2295
   ClientLeft      =   10335
   ClientTop       =   4755
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm공용_PrinterSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6135
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4980
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3780
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "▶ 프린터 목록"
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   1260
   End
End
Attribute VB_Name = "frm공용_PrinterSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HWND_BROADCAST    As Long = &HFFFF&
Private Const WM_WININICHANGE   As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Screen.MousePointer = 11
    With List1
        Call WriteProfileString("Windows", "Device", List1.List(List1.ListIndex)) ''** INI 또는 레지스트리에 기본프린터 정보를 등록...
        Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "Windows") ''** 모든 응용프로그램이 INI 파일 또는 레지스트리 정보를 다시 읽도록 한다...
    End With
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lngi As Long
    Dim lngLength As Long
    Dim X As Printer
    Dim str기본프린터 As String
    
    Me.Height = 2805
    Me.Width = 6255
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Screen.MousePointer = 11
    With List1
        .Clear
    
        For Each X In Printers
            .AddItem X.DeviceName & "," & X.DriverName & "," & X.Port
        Next X
    
        str기본프린터 = Space(8192) ''** 1024 * 8 = 8192
        Call GetProfileString("Windows", "Device", "", str기본프린터, Len(str기본프린터)) ''** 기본프린터 정보를 알아낸다...
        
        For lngi = 0 To .ListCount - 1 ''** 기본프린터가 선택된 상태로 설정...
            lngLength = Len(.List(lngi))
            If .List(lngi) = Left(str기본프린터, lngLength) Then
                .ListIndex = lngi
                Exit For
            End If
        Next lngi
    End With
    Screen.MousePointer = 0
End Sub

Private Sub List1_DblClick()
    Screen.MousePointer = 11
    With List1
        Call WriteProfileString("Windows", "Device", .Text) ''** INI 또는 레지스트리에 기본프린터 정보를 등록...
        Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "Windows") ''** 모든 응용프로그램이 INI 파일 또는 레지스트리 정보를 다시 읽도록 한다...
    End With
    Screen.MousePointer = 0
    Unload Me
End Sub
