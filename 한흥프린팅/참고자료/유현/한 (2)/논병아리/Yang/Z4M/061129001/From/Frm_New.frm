VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Frm_New 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15090
   Icon            =   "Frm_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   15090
   Begin VB.Frame Fam_C 
      Caption         =   "바코드 - 레코드 처리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   14655
      Begin VB.CommandButton Cmd_Printer 
         Caption         =   "발행"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   10830
         TabIndex        =   5
         Top             =   6630
         Width           =   3630
      End
      Begin FPSpread.vaSpread Spr_C 
         Height          =   6375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   14415
         _Version        =   393216
         _ExtentX        =   25426
         _ExtentY        =   11245
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   30
         ShadowColor     =   16761024
         SpreadDesigner  =   "Frm_New.frx":628A
      End
   End
   Begin VB.Frame Fam_B 
      Caption         =   "바코드 - 포맷 생성"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   240
      TabIndex        =   1
      Top             =   915
      Width           =   14400
      Begin VB.ComboBox Cbo_HeadDarkness 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":7430
         Left            =   11895
         List            =   "Frm_New.frx":748E
         TabIndex        =   18
         Text            =   "15"
         Top             =   2940
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_PrinterSpeed 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":750B
         Left            =   11895
         List            =   "Frm_New.frx":7521
         TabIndex        =   16
         Text            =   " 3"
         Top             =   2415
         Width           =   1710
      End
      Begin VB.ComboBox cbo_Port 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":753D
         Left            =   11895
         List            =   "Frm_New.frx":755C
         TabIndex        =   10
         Text            =   "COM1"
         Top             =   525
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_Baud 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":7596
         Left            =   11895
         List            =   "Frm_New.frx":75A9
         TabIndex        =   9
         Text            =   "9600"
         Top             =   1245
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_Dpi 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11895
         TabIndex        =   8
         Text            =   "200 dpi"
         Top             =   1920
         Width           =   1710
      End
      Begin VB.TextBox Txt_CenterX 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11895
         TabIndex        =   7
         Text            =   "0"
         Top             =   3570
         Width           =   810
      End
      Begin VB.TextBox Txt_CenterY 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11895
         TabIndex        =   6
         Text            =   "0"
         Top             =   3915
         Width           =   810
      End
      Begin FPSpread.vaSpread Spr_B 
         Height          =   5895
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9495
         _Version        =   393216
         _ExtentX        =   16748
         _ExtentY        =   10398
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   11
         Protect         =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         ShadowColor     =   16761024
         SpreadDesigner  =   "Frm_New.frx":75D0
      End
      Begin VB.Label Label5 
         Caption         =   "해드 온도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10800
         TabIndex        =   19
         Top             =   2970
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "프린터 스피드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10380
         TabIndex        =   17
         Top             =   2445
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "통신 포트 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10260
         TabIndex        =   15
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "데이터 전송 속도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10080
         TabIndex        =   14
         Top             =   1260
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "장비 DPI 값"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10575
         TabIndex        =   13
         Top             =   1980
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "원점 - X(mm)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10395
         TabIndex        =   12
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "원점 - Y"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10305
         TabIndex        =   11
         Top             =   3990
         Width           =   990
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   7890
      Left            =   45
      TabIndex        =   0
      Top             =   375
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   13917
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "바코드 - 포맷 생성"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "바코드 - 레코드 처리"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'***  Description   : Printer 발행
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************
Private Sub Cmd_Printer_Click()
'call Label_Print

 Dim Li_LocCount As Integer
 Dim Ls_Cbo_Baud As String
 Dim Ls_SprB_Data(10) As String
 Dim Ls_SprBTempData(15) As String
 Dim Ll_SprB_Count As Integer
 Dim Li_Count As Integer
 Dim Ll_SprB_MaxCount As Integer
 Dim Ls_TempDataSetting As String
 Dim Li_SprCLocCount As Integer
 Dim Li_SprCount As Integer
 Dim Ls_SprCTempData(30) As String
 Dim Ls_SprC_Data(12) As String
 Dim Ll_SprC_MaxCount As Integer
 Dim Li_SprC_MaxTemp As Integer
 Dim Li_SprC_MaxTemp1 As Integer
 Dim Li_SprC_MaxCountTemp As Integer
 Dim Ls_SprOrnData(10) As String
 Dim Count As Integer
 
 Dim Ls_SprDataA(10) As String
 Dim Li_SprC_MaxTempA As Integer
 Dim Li_CountA As Integer
 
 Dim Ls_SprDataB(10) As String
 Dim Li_SprC_MaxTempB As Integer
 Dim Li_CountB As Integer
 
 Dim Ls_SprDataC(10) As String
 Dim Li_SprC_MaxTempC As Integer
 Dim Li_CountC As Integer
 
 Dim Ls_SprDataD(10) As String
 Dim Li_SprC_MaxTempD As Integer
 Dim Li_CountD As Integer
 
 Dim Ls_SprDataE(10) As String
 Dim Li_SprC_MaxTempE As Integer
 Dim Li_CountE As Integer
 
 Dim Ls_SprDataF(10) As String
 Dim Li_SprC_MaxTempF As Integer
 Dim Li_CountF As Integer
 Dim Ls_TrueBarcodeDataNo As String

 Dim Ls_SprDataG(10) As String
 Dim Li_SprC_MaxTempG As Integer
 Dim Li_CountG As Integer
 
 Dim Ls_SprDataH(10) As String
 Dim Li_SprC_MaxTempH As Integer
 Dim Li_CountH As Integer
 
 Dim Ls_SprDataI(10) As String
 Dim Li_SprC_MaxTempI As Integer
 Dim Li_CountI As Integer
                      
 Dim Ls_SprDataJ(10) As String
 Dim Li_SprC_MaxTempJ As Integer
 Dim Li_CountJ As Integer

 Dim Ls_SprDataK(10) As String
 Dim Li_SprC_MaxTempK As Integer
 Dim Li_CountK As Integer

 Dim Ls_SprDataL(10) As String
 Dim Li_SprC_MaxTempL As Integer
 Dim Li_CountL As Integer
 
 Dim Ls_SprDataZ(10) As String
 Dim Li_SprC_MaxTempZ As Integer
 Dim Li_CountZ As Integer
 
 Dim Li_CountBarcode As Integer
 Dim Li_SprC_MaxBarcode As Integer
 Dim Ls_BarcodeData(2) As String
 Dim Ls_MaxBarcodeData As String
 Dim Li_CountBarcodeA As Integer
 Dim Li_SprC_MaxBarcodeA As Integer
 Dim Ls_ColorSave As String
 Dim Ls_TrueBarcodeData As String
 Dim Ls_TrueBarcodeDataTemp As String * 13
 Dim Ls_Temp As String
 Dim Ls_SprC_Datatmp As String
 
 Cmd_Printer.Enabled = False
 MousePointer = 11
 Li_LocCount = 0
 
 Frm_Main.Mcom.CommPort = Right(Me.cbo_Port, 1)
 Frm_Main.Mcom.PortOpen = True
 Frm_Main.Mcom.RThreshold = 1
 Ls_Cbo_Baud = Trim(Me.Cbo_Baud)
 
 Frm_Main.Mcom.Settings = Ls_Cbo_Baud & ",n,8,1"

 Li_Count = 0
 Ll_SprB_Count = 0
 Ll_SprB_MaxCount = 0
                  
                
 With Spr_B
      
      Li_LocCount = .MaxRows
        
      For Li_LocCount = 1 To .MaxRows
        
          Li_Count = Li_Count + 1
          Ll_SprB_Count = Ll_SprB_Count + 1
            
         .Row = Ll_SprB_Count
         .Col = 1
          Ls_SprB_Data(1) = .Text
            
         .Row = Ll_SprB_Count
         .Col = 2
          Ls_SprB_Data(2) = .Text
            
         .Row = Ll_SprB_Count
         .Col = 3
          Ls_SprB_Data(3) = .Text
            
         .Row = Ll_SprB_Count
         .Col = 4
          Ls_SprB_Data(4) = .Text
            
         .Row = Ll_SprB_Count
         .Col = 5
          Ls_SprB_Data(5) = .Text
           
         .Row = Ll_SprB_Count
         .Col = 6
          Ls_SprB_Data(6) = .Text
           
          Ls_SprBTempData(Li_Count) = Ls_SprB_Data(1) & "," & Ls_SprB_Data(2) & "," & _
                                      Ls_SprB_Data(3) & "," & Ls_SprB_Data(4) & "," & _
                                      Ls_SprB_Data(5) & "," & Ls_SprB_Data(6) & ","
          
          Ll_SprB_MaxCount = Ll_SprB_MaxCount + 1
        
      Next Li_LocCount
                 
 End With

 
 With Spr_C
        
      Li_SprCLocCount = 0
      Li_SprCount = 0
      Ll_SprC_MaxCount = 0

      Li_SprCLocCount = .MaxRows
       
      For Li_SprCLocCount = 1 To .MaxRows
        
          Li_SprCount = Li_SprCount + 1
          Ll_SprC_Count = Ll_SprC_Count + 1
            
         .Row = Ll_SprC_Count
         .Col = 1
          Ls_SprC_Data(1) = .Text
            
         .Row = Ll_SprC_Count
         .Col = 2
          Ls_SprC_Data(2) = .Text
            
         .Row = Ll_SprC_Count
         .Col = 3
          Ls_SprC_Data(3) = .Text
            
         .Row = Ll_SprC_Count
         .Col = 4
          Ls_SprC_Data(4) = .Text
            
         .Row = Ll_SprC_Count
         .Col = 5
          Ls_SprC_Data(5) = .Text
            
         .Row = Ll_SprC_Count
         .Col = 6
          Ls_SprC_Data(6) = .Text
         '  Ls_SprC_Datatmp = Ls_SprC_Data(6)
         .Row = Ll_SprC_Count
         .Col = 7
          Ls_SprC_Data(7) = .Text

         .Row = Ll_SprC_Count
         .Col = 8
          Ls_SprC_Data(8) = .Text
          
         .Row = Ll_SprC_Count
         .Col = 9
          Ls_SprC_Data(9) = .Text
          
          .Row = Ll_SprC_Count
         .Col = 10
          Ls_SprC_Data(10) = .Text
          
          .Row = Ll_SprC_Count
         .Col = 11
          Ls_SprC_Data(11) = .Text
          
          If Ls_SprC_Data(1) = "1" Then
                
                GS_Path = App.Path & "\Setting\Setting.ini"

                Open GS_Path For Input As #2
                  
                     While Not EOF(2)
                           Line Input #2, Ls_TempDataSetting
                     Wend
                
                Close #2
            
                Li_Count = 0
             
                LS_StrarryS = Split(Ls_TempDataSetting, ",")
             
                cnt = UBound(LS_StrarryS)
            
                If cnt > 0 Then
                      
                      Do
                            If Ls_SprC_Data(4) = LS_StrarryS(Li_Count) Then
                                  Ls_SprC_DataTemp = Ls_SprC_Data(4)
                                  Ls_SprC_Data(4) = LS_StrarryS(Li_Count + 1)
                
                                  Exit Do
                
                            Else
                            
                            End If
                
                            Li_Count = Li_Count + 1
                   
                      Loop Until Li_Count = cnt
                Else
                
                End If
            
                             
                          
                Ls_SprCTempData(Li_SprCount) = Ls_SprC_Data(1) & "," & Ls_SprC_Data(2) & "," & _
                                               Ls_SprC_Data(3) & "," & Ls_SprC_Data(4) & "," & _
                                               Ls_SprC_Data(5) & "," & Ls_SprC_Data(6) & "," & _
                                               Ls_SprC_Data(7) & "," & Ls_SprC_Data(8) & "," & _
                                               Ls_SprC_Data(9) & "," & Ls_SprC_Data(10) & "," & Ls_SprC_Data(11) & ","
                
                Ls_TempData = Ls_TempData & Li_SprCount & ","
                Ll_SprC_MaxCount = Ll_SprC_MaxCount + 1
          Else
               
               Ls_SprCTempData(Li_SprCount) = Ls_SprC_Data(1) & "," & Ls_SprC_Data(2) & "," & _
                                              Ls_SprC_Data(3) & "," & Ls_SprC_Data(4) & "," & _
                                              Ls_SprC_Data(5) & "," & Ls_SprC_Data(6) & "," & _
                                              Ls_SprC_Data(7) & "," & Ls_SprC_Data(8) & "," & _
                                              Ls_SprC_Data(9) & "," & Ls_SprC_Data(10) & "," & Ls_SprC_Data(11) & ","
          End If
      
      Next Li_SprCLocCount
                 
 End With
 
 
 On Error GoTo Errorhandler
 
 Li_SprC_MaxTemp = 0
 Li_SprC_MaxTemp1 = 0


 LS_Strarry = Split(Ls_TempData, ",")
             
 cnt = UBound(LS_Strarry)

 If cnt > 0 Then
     
       Do
        
            LS_Strarry1 = Split(Ls_SprCTempData(LS_Strarry(Li_SprC_MaxTemp)), ",")
            Li_SprC_MaxTemp1 = 0
            
            Count = UBound(LS_Strarry1)
            
            If Count > 0 Then
                 Do
                    Ls_SprOrnData(Li_SprC_MaxTemp1) = LS_Strarry1(Li_SprC_MaxTemp1)
                    Li_SprC_MaxTemp1 = Li_SprC_MaxTemp1 + 1
                 Loop Until Li_SprC_MaxTemp1 = Count
            End If
            

            '----------------------------------------------------------------------------------------------------------------------------
            ' Starting Point
            '----------------------------------------------------------------------------------------------------------------------------
            Frm_Main.Mcom.Output = "^XA"
            Frm_Main.Mcom.Output = "^PR" & Cbo_PrinterSpeed.Text & "^FS"
            Frm_Main.Mcom.Output = "^LH" & Txt_CenterX.Text & "," & Txt_CenterY.Text & "^FS"
            
            
            'LOT-NO
            ' Frm_Main.Mcom.Output = "^FO60,20" & "^A0,24,24" & "^FD" & "LOT-NO:" & Ls_SprOrnData(1) & Ls_SprOrnData(2) & "^FS"
            
            '----------------------------------------------------------------------------------------------------------------------------
            ' 1 A 바코드
            '----------------------------------------------------------------------------------------------------------------------------
            LS_StrarryA = Split(Ls_SprBTempData(1), ",")  '바코드 상
            Li_CountA = UBound(LS_StrarryA)
            Li_SprC_MaxTempA = 0
            
            If Li_CountA > 0 Then
                Do
                    Ls_SprDataA(Li_SprC_MaxTempA) = LS_StrarryA(Li_SprC_MaxTempA)
                    Li_SprC_MaxTempA = Li_SprC_MaxTempA + 1
                Loop Until Li_SprC_MaxTempA = Li_CountA
            End If
            
            If Ls_SprDataA(0) = "1" Then
            
                
                    Frm_Main.Mcom.Output = "^FO" & Ls_SprDataA(1) & "," & Ls_SprDataA(2) & "^BY" & Ls_SprDataA(3) & "^BCN," & Ls_SprDataA(4) & ",N,N,N^FD" & Ls_SprOrnData(1) & Ls_SprOrnData(2) & Ls_SprC_DataTemp & Ls_SprOrnData(4) & "^FS"
                    'Frm_Main.Mcom.Output = "^FO" & Ls_SprDataA(1) & "," & Ls_SprDataA(2) & "^A0," & Ls_SprDataA(3) & "," & Ls_SprDataA(4) & "^FD" & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(3) & "^FS"
                     Debug.Print "^FO" & Ls_SprDataA(1) & "," & Ls_SprDataA(2) & "^BY" & Ls_SprDataA(3) & "^BCN," & Ls_SprDataA(4) & ",N,N,N^FD" & Ls_SprOrnData(1) & Ls_SprOrnData(2) & Ls_SprC_DataTemp & Ls_SprOrnData(4) & "^FS"; " 1번째 : 바코드상"
               
            End If


            '----------------------------------------------------------------------------------------------------------------------------
            ' 2 B C/NO 출력
            '----------------------------------------------------------------------------------------------------------------------------
            Li_SprC_MaxTempB = 0
            LS_StrarryB = Split(Ls_SprBTempData(2), ",")   'C/NO
            Li_CountB = UBound(LS_StrarryB)

            If Li_CountB > 0 Then
                  Do
                     Ls_SprDataB(Li_SprC_MaxTempB) = LS_StrarryB(Li_SprC_MaxTempB)
                     Li_SprC_MaxTempB = Li_SprC_MaxTempB + 1

                     Loop Until Li_SprC_MaxTempB = Li_CountB
            End If

            If Ls_SprDataB(0) = "1" Then
                  If Ls_SprOrnData(5) = "" Then
                     Frm_Main.Mcom.Output = "^FO" & Ls_SprDataB(1) & "," & Ls_SprDataB(2) & "^A0," & Ls_SprDataB(3) & "," & Ls_SprDataB(4) & "^FD" & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(4) & "^FS"
                     Debug.Print "^FO" & Ls_SprDataB(1) & "," & Ls_SprDataB(2) & "^A0," & Ls_SprDataB(3) & "," & Ls_SprDataB(4) & "^FD" & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(4) & "^FS"; "   2번째 : C/NO"
                  Else
                     Frm_Main.Mcom.Output = "^FO" & Ls_SprDataB(1) & "," & Ls_SprDataB(2) & "^A0," & Ls_SprDataB(3) & "," & Ls_SprDataB(4) & "^FD" & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(5) & "^FS"
                   '  Debug.Print "^FO" & Ls_SprDataB(1) & "," & Ls_SprDataB(2) & "^A0," & Ls_SprDataB(3) & "," & Ls_SprDataB(4) & "^FD" & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(4) & "^FS"; "   2번째 : C/NO"
                  End If
                  
            End If





            '----------------------------------------------------------------------------------------------------------------------------
            ' 3 C 색상 출력
            '----------------------------------------------------------------------------------------------------------------------------
            Li_SprC_MaxTempC = 0
            LS_StrarryC = Split(Ls_SprBTempData(3), ",")   '색상
            Li_CountC = UBound(LS_StrarryC)
            Li_SprC_MaxTempC = 0

            If Li_CountC > 0 Then
                  Do
                     Ls_SprDataC(Li_SprC_MaxTempC) = LS_StrarryC(Li_SprC_MaxTempC)
                     Li_SprC_MaxTempC = Li_SprC_MaxTempC + 1

                  Loop Until Li_SprC_MaxTempC = Li_CountC
            End If

            If Ls_SprDataC(0) = "1" Then
                  Call Label_Print(Ls_SprOrnData(3), Val(Ls_SprDataC(1)), Val(Ls_SprDataC(2)), Val(Ls_SprDataC(3)))

            End If

            '----------------------------------------------------------------------------------------------------------------------------
            ' 4 D 품명 출력 (LOT-NO)
            '----------------------------------------------------------------------------------------------------------------------------
            LS_StrarryD = Split(Ls_SprBTempData(4), ",")
            Li_CountD = UBound(LS_StrarryD)
            Li_SprC_MaxTempD = 0

            If Li_CountD > 0 Then
                  Do
                     Ls_SprDataD(Li_SprC_MaxTempD) = LS_StrarryD(Li_SprC_MaxTempD)
                     Li_SprC_MaxTempD = Li_SprC_MaxTempD + 1

                  Loop Until Li_SprC_MaxTempD = Li_CountD

            End If

            If Ls_SprDataD(0) = "1" Then
                Call Label_Print(Ls_SprOrnData(6), Val(Ls_SprDataD(1)), Val(Ls_SprDataD(2)), Val(Ls_SprDataD(3)))
                'Frm_Main.Mcom.Output = "^FO" & Ls_SprDataE(1) & "," & Ls_SprDataE(2) & "^A0," & Ls_SprDataE(3) & "," & Ls_SprDataE(4) & "^FD" & Ls_SprOrnData(6) & "^FS"
                'Debug.Print "^FO" & Ls_SprDataE(1) & "," & Ls_SprDataE(2) & "^BY" & Ls_SprDataE(3) & "^BAN," & Ls_SprDataE(4) & ",N,N,N^FD" & Ls_SprOrnData(6) & "^FS"; "   5번째 : 품명"
            End If


            '----------------------------------------------------------------------------------------------------------------------------
            ' 5 E LOT-NO 출력 (판매가격)
            '----------------------------------------------------------------------------------------------------------------------------
            Li_SprC_MaxTempE = 0
            LS_StrarryE = Split(Ls_SprBTempData(5), ",")
            Li_CountE = UBound(LS_StrarryE)

            If Li_CountE > 0 Then

                  Do

                     Ls_SprDataE(Li_SprC_MaxTempE) = LS_StrarryE(Li_SprC_MaxTempE)
                     Li_SprC_MaxTempE = Li_SprC_MaxTempE + 1

                     Loop Until Li_SprC_MaxTempE = Li_CountE

            End If

            If Ls_SprDataE(0) = "1" Then
                '  If Ls_SprOrnData(2) = "" Then
                     Frm_Main.Mcom.Output = "^FO" & Ls_SprDataE(1) & "," & Ls_SprDataE(2) & "^A0," & Ls_SprDataE(3) & "," & Ls_SprDataE(4) & "^FD" & "LOT-NO:" & Ls_SprOrnData(7) & Ls_SprOrnData(1) & Ls_SprOrnData(2) & "^FS"
                         Debug.Print "^FO" & Ls_SprDataF(1) & "," & Ls_SprDataE(2) & "^A0," & Ls_SprDataE(3) & "," & Ls_SprDataE(4) & "^FD" & "LOT-NO:  " & Ls_SprOrnData(1) & Ls_SprOrnData(2) & "^FS"; "   6번째 : LOT-NO"
            End If


            '----------------------------------------------------------------------------------------------------------------------------
            ' 6 G 가격 출력
            '----------------------------------------------------------------------------------------------------------------------------
            LS_StrarryG = Split(Ls_SprBTempData(6), ",")
            Li_CountG = UBound(LS_StrarryG)
            Li_SprC_MaxTempG = 0

            If Li_CountG > 0 Then

                  Do
                     Ls_SprDataG(Li_SprC_MaxTempG) = LS_StrarryG(Li_SprC_MaxTempG)
                     Li_SprC_MaxTempG = Li_SprC_MaxTempG + 1

                  Loop Until Li_SprC_MaxTempG = Li_CountG

            End If

            If Ls_SprDataG(0) = "1" Then
            Ls_Temp = Ls_SprOrnData(8)
                    If InStr(Ls_SprOrnData(8), ",") = 0 And Len(Ls_SprOrnData(8)) > 0 Then
                      If Len(Ls_SprOrnData(8)) <= 3 Then

                          Ls_SprOrnData(8) = Ls_SprOrnData(8) & ",000"
                      Else
                          Ls_SprOrnData(8) = Mid(Ls_SprOrnData(8), 1, (Len(Ls_SprOrnData(8)) - 3)) & "," & Right(Ls_SprOrnData(8), 3)
                      End If
                 End If

                Frm_Main.Mcom.Output = "^FO" & Ls_SprDataG(1) & "," & Ls_SprDataG(2) & "^A0," & Ls_SprDataG(3) & "," & Ls_SprDataG(4) & "^FD" & Ls_SprOrnData(8) & "^FS"
                 Debug.Print "^FO" & Ls_SprDataG(1) & "," & Ls_SprDataG(2) & "^A0," & Ls_SprDataG(3) & "," & Ls_SprDataG(4) & "^FD" & Ls_SprOrnData(8) & "^FS"; "   7번째 : 가격"

            End If


            '----------------------------------------------------------------------------------------------------------------------------
            ' Ending Point
            '----------------------------------------------------------------------------------------------------------------------------
          Frm_Main.Mcom.Output = "^PQ" & Ls_SprOrnData(9) & ",0,1,Y^XZ"
          Frm_Main.Mcom.Output = "~HS"
          Frm_Main.Mcom.Output = "^IDR:*.*"

          Li_SprC_MaxTemp = Li_SprC_MaxTemp + 1
          Ls_MaxBarcodeData = ""
 
       Loop Until Li_SprC_MaxTemp = cnt
        
  End If
  
  Li_SprC_MaxCountTemp = Li_SprC_MaxCountTemp + 1
          
  Frm_Main.Mcom.PortOpen = False
  MousePointer = 0
  Cmd_Printer.Enabled = True
  
Errorhandler:
    If Err.Number <> 0 Then
        MsgBox ("바코드 발행오류 : " & Err.Description)
    End If
    
End Sub



'***********************************************************************************
'***  Description   :  폼 Activate 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Activate()
 
 Frm_Main.Mun_Save.Enabled = True
 Frm_Main.Mun_Close.Enabled = True
 Frm_Main.Mun_AllClose.Enabled = True
 Frm_Main.Mun_Setting.Enabled = True
 Frm_Main.Mun_View.Enabled = True
 Frm_Main.Mun_Windows.Enabled = True
 Frm_Main.tlbMain.Buttons(4).Enabled = True

End Sub

'***********************************************************************************
'***  Description   :  폼 로드 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Load()
 
 Dim LS_Filename As String
 

' Frm_Main.Mcom.Output = "~JA"

 
 GS_FromCount = GS_FromCount + 1
 Me.Tag = Str(GS_FromCount)
 
 If CurrentFilename <> "" Then
       
       Me.Caption = CurrentFilename
       CurrentFilename = ""
 Else

       Me.Caption = "새로운 파일"
       
 End If

End Sub

'***********************************************************************************
'***  Description   :  폼 언로드 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Unload(Cancel As Integer)
 
 GS_FromCount = GS_FromCount - 1

 If GS_FromCount = 0 Then
       
       Frm_Main.Mun_Save.Enabled = False
       Frm_Main.Mun_Close.Enabled = False
       Frm_Main.Mun_AllClose.Enabled = False
       Frm_Main.Mun_Setting.Enabled = False
       Frm_Main.Mun_View.Enabled = False
       Frm_Main.Mun_Windows.Enabled = False
       Frm_Main.tlbMain.Buttons(4).Enabled = False
       CurrentFilename = ""
 
 End If

End Sub

'***********************************************************************************
'***  Description   : TabStrip 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub TabStrip_Click()
  
 If TabStrip.SelectedItem.Index = 1 Then
      
      Fam_B.Visible = True
      Fam_C.Visible = False
 
 ElseIf TabStrip.SelectedItem.Index = 2 Then
      
      Fam_B.Visible = False
      Fam_C.Visible = True
 
 End If

End Sub

