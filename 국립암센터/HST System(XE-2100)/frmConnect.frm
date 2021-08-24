VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmConnect 
   BorderStyle     =   1  '단일 고정
   Caption         =   "연결선택"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   7590
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtTemp 
      Height          =   270
      Left            =   960
      TabIndex        =   2
      Top             =   1710
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "확 인"
      Height          =   525
      Left            =   3105
      TabIndex        =   1
      Top             =   1410
      Width           =   1305
   End
   Begin FPSpread.vaSpread vasComList 
      Height          =   1125
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   7215
      _Version        =   196613
      _ExtentX        =   12726
      _ExtentY        =   1984
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   3
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "frmConnect.frx":0442
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirm_Click()
    Dim i, j As Integer
    
    j = -1
    For i = 1 To vasComList.DataRowCnt
        vasComList.Row = i
        vasComList.Col = 1
        If vasComList.Value = 1 Then
            j = 1
            
            If Trim(GetText(vasComList, i, 2)) = "IPU1" Then
                IPU1.ComPort = Trim(GetText(vasComList, i, 3))
                IPU1.Speed = Trim(GetText(vasComList, i, 4))
                IPU1.Parity = Trim(GetText(vasComList, i, 5))
                IPU1.DataBit = Trim(GetText(vasComList, i, 6))
                IPU1.StopBit = Trim(GetText(vasComList, i, 7))
                IPU1.RTSEnable = Trim(GetText(vasComList, i, 8))
                IPU1.DTREnable = Trim(GetText(vasComList, i, 9))
                IPU1.ConnectFlag = True
            ElseIf Trim(GetText(vasComList, i, 2)) = "IPU2" Then
                IPU2.ComPort = Trim(GetText(vasComList, i, 3))
                IPU2.Speed = Trim(GetText(vasComList, i, 4))
                IPU2.Parity = Trim(GetText(vasComList, i, 5))
                IPU2.DataBit = Trim(GetText(vasComList, i, 6))
                IPU2.StopBit = Trim(GetText(vasComList, i, 7))
                IPU2.RTSEnable = Trim(GetText(vasComList, i, 8))
                IPU2.DTREnable = Trim(GetText(vasComList, i, 9))
                IPU2.ConnectFlag = True
            End If
        Else
            If Trim(GetText(vasComList, i, 2)) = "IPU1" Then
                IPU1.ConnectFlag = False
            ElseIf Trim(GetText(vasComList, i, 2)) = "IPU2" Then
                IPU2.ConnectFlag = False
            End If
        End If
    Next i
    
    If j <> 1 Then
        If MsgBox("선택한 연결이 없습니다. 다시 선택하시겠습니까? " & vbCrLf & vbCrLf & "[No]를 선택하면 프로그램 종료합니다", vbInformation + vbYesNo, "알림") = vbNo Then
            'End
        Else
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
       
    lRow = 0
    For i = 1 To 4
        db_tmp = ""
        Call GetPrivateProfileString("COM " & CStr(i), "Use", "", db_tmp, 20, App.Path & "\Interface.ini")
        txtTemp = Trim(db_tmp)
        If Trim(txtTemp) <> "" Then
'            lRow = lRow + 1
'
'            vasComList.Row = lRow
'            vasComList.Col = 1
'            If Trim(txtTemp) = "1" Then
'                vasComList.Value = 1
'            Else
'                vasComList.Value = 0
'            End If
            
            db_tmp = ""
            Call GetPrivateProfileString("COM " & CStr(i), "Gubun", "", db_tmp, 20, App.Path & "\Interface.ini")
            txtTemp = Trim(db_tmp)
            
            If Left(Trim(txtTemp), 3) = "IPU" Then
                lRow = lRow + 1
                
                SetText vasComList, Trim(txtTemp), lRow, 2
                
                vasComList.Row = lRow
                vasComList.Col = 1
                If Trim(txtTemp) = "IPU1" Then
                    If frmInterface.MSComm1.PortOpen = True Then
                        vasComList.Value = 1
                    Else
                        vasComList.Value = 0
                    End If
                ElseIf Trim(txtTemp) = "IPU2" Then
                    If frmInterface.MSComm2.PortOpen = True Then
                        vasComList.Value = 1
                    Else
                        vasComList.Value = 0
                    End If
                End If
                
                SetText vasComList, CStr(i), lRow, 3
                
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 4
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 5
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 6
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 7
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 8
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                SetText vasComList, Trim(txtTemp), lRow, 9
            End If
        End If
    Next i

End Sub
