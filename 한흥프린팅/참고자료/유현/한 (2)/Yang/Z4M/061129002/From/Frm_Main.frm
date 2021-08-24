VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm Frm_Main 
   BackColor       =   &H8000000C&
   Caption         =   "한흥프린팅 #2 바코드 Ver 1.0"
   ClientHeight    =   9750
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14115
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '화면 가운데
   WindowState     =   2  '최대화
   Begin MSCommLib.MSComm Mcom 
      Left            =   6030
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CDlog 
      Left            =   6045
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.dc)|*.dc"
      Flags           =   4
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '위 맞춤
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "iglToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "새 폴리곤 문서을 만듭니다."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "폴리곤 문서를 엽니다."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "폴리곤 문서를 저장합니다."
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iglToolBar 
      Left            =   6570
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":639C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":64AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":65C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":66D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":67E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":68F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":6A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":6B1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mun_Files 
      Caption         =   "파일(&F)"
      WindowList      =   -1  'True
      Begin VB.Menu Mun_New 
         Caption         =   "새로 만들기"
      End
      Begin VB.Menu Mun_Open 
         Caption         =   "열기.."
      End
      Begin VB.Menu Mun_Save 
         Caption         =   "저장(&S)"
      End
      Begin VB.Menu Spr 
         Caption         =   "-"
      End
      Begin VB.Menu Mun_Close 
         Caption         =   "닫기"
      End
      Begin VB.Menu Mun_AllClose 
         Caption         =   "모두닫기"
      End
      Begin VB.Menu Mun_spr 
         Caption         =   "-"
      End
      Begin VB.Menu Mun_Exit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu Mun_View 
      Caption         =   "보기(&V)"
      Begin VB.Menu Mun_Tool 
         Caption         =   "도구 모음(&T)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Mun_Windows 
      Caption         =   "창(&W)"
      Begin VB.Menu Mun_GeDan 
         Caption         =   "계단식 배열(&C)"
      End
      Begin VB.Menu Mun_BaD 
         Caption         =   "바둑판식 배열(&T)"
      End
      Begin VB.Menu Mun_Icon 
         Caption         =   "아이콘 정렬(&A)"
      End
   End
   Begin VB.Menu Mun_Setting 
      Caption         =   "환경설정(&W)"
      Begin VB.Menu Mun_SettingColor 
         Caption         =   "색상설정"
      End
      Begin VB.Menu Mun_SettingSize 
         Caption         =   "사이즈설정"
      End
   End
   Begin VB.Menu Mun_Help 
      Caption         =   "도움말(&H)"
      Begin VB.Menu Mun_Helps 
         Caption         =   "도울말(&H)"
      End
      Begin VB.Menu Mun_InFor 
         Caption         =   "ID 정보(&A)"
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'***********************************************************************************
'***  Description   :  MDI 폼 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub MDIForm_Load()
 
 Mun_Save.Enabled = False
 Mun_Close.Enabled = False
 Mun_AllClose.Enabled = False
 Mun_Setting.Enabled = False
 Mun_View.Enabled = False
 Mun_Windows.Enabled = False
 Frm_Main.tlbMain.Buttons(4).Enabled = False
 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

'***********************************************************************************
'***  Description   :  자식 폼 모두 닫기 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_AllClose_Click()
 
Dim Li_FromCount As Integer

 For Li_FromCount = 1 To GS_FromCount
     
     Unload Me.ActiveForm
     
 Next Li_FromCount

End Sub


'***********************************************************************************
'***  Description   :  폼 바둑판식 배열
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_BaD_Click()

 Frm_Main.Arrange vbTileHorizontal
 
End Sub

'***********************************************************************************
'***  Description   :  폼(자식폼) 닫기 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_Close_Click()

 Unload Me.ActiveForm
 
End Sub

'***********************************************************************************
'***  Description   :  프로그램 종료 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_Exit_Click()

 Unload Me

End Sub

'***********************************************************************************
'***  Description   :  폼 계단식 배열
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_GeDan_Click()
 
 Frm_Main.Arrange vbCascade
 
End Sub

'***********************************************************************************
'***  Description   :  도움말  정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_Helps_Click()

 Frm_Information.Show 0
 
End Sub

'***********************************************************************************
'***  Description   :  폼 아이콘 배열
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_Icon_Click()

  Frm_Main.Arrange vbTileHorizontal
  
End Sub

'***********************************************************************************
'***  Description   :  URL 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_InFor_Click()

 Call S_HomePage("http://www.idif.co.kr")
 
End Sub

'***********************************************************************************
'***  Description   :  폼(자식폼) 생성 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_New_Click()

Dim Frm_New As New Frm_New

 Frm_New.Show

 Frm_New.Width = 14000
 Frm_New.Height = 9000
 
 Mun_Save.Enabled = True
 Mun_Close.Enabled = True
 Mun_AllClose.Enabled = True
 Mun_Setting.Enabled = True
 Mun_View.Enabled = True
 Mun_Windows.Enabled = True
 Frm_Main.tlbMain.Buttons(4).Enabled = True
 
End Sub

'***********************************************************************************
'***  Description   : TXT파일 OPEN 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_Open_Click()

Dim Li_FileNumber As Integer
Dim Li_FrmCount As Integer
Dim LS_Filename As String
Dim Ls_TempData As String
Dim Frm_New As New Frm_New
Dim spacePos As Integer
Dim Li_Count As Integer
Dim i, j As Long
Dim Ls_Count As Integer

On Error Resume Next

With CDlog
         
         .CancelError = True
         .FileName = Getcursor
         .InitDir = App.Path
         .Filter = "파일(*.Han)|*.Han"
         .DefaultExt = "*.Han"
         .FilterIndex = 2
         .ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
          
    On Error GoTo 0
          LS_Filename = .FileName
End With

Li_FrmCount = 0


If Me.CDlog.FileTitle <> "" Then

      CurrentFilename = Me.CDlog.FileTitle
      Li_FileNumber = FreeFile
      
      Open LS_Filename For Input As #2 ' 파일을 입력모드로 오픈한다.
       
        While Not EOF(2)
        
           Line Input #2, Ls_TempData
           
        Wend
      
      Close #2
      
      Li_Count = 0

      LS_Strarry = Split(Ls_TempData, ",")

      cnt = UBound(LS_Strarry)

      If cnt > 0 Then
      
         Do
            Debug.Print LS_Strarry(Li_Count)
            Li_Count = Li_Count + 1
               
         Loop Until Li_Count = cnt
      
      End If
      
      Frm_New.Caption = CDlog.FileTitle
      
      Me.ActiveForm.cbo_Port.Text = LS_Strarry(0)
      Me.ActiveForm.Cbo_Baud.Text = LS_Strarry(1)
      Me.ActiveForm.Cbo_Dpi.Text = LS_Strarry(2)
      Me.ActiveForm.Txt_CenterX.Text = LS_Strarry(3)
      Me.ActiveForm.Txt_CenterY.Text = LS_Strarry(4)
      Me.ActiveForm.Cbo_PrinterSpeed.Text = LS_Strarry(5)
      Me.ActiveForm.Cbo_HeadDarkness.Text = LS_Strarry(6)
 
      Ls_Count = 7
 
      With Me.ActiveForm.Spr_B
      
           j = 0
      
           For i = 1 To .MaxRows
                  j = j + 1
                  .Row = j
                  .Col = 1
                  .Text = LS_Strarry(Ls_Count)
                 
                  .Row = j
                  .Col = 2
                  .Text = LS_Strarry(Ls_Count + 1)
                  
                  .Row = j
                  .Col = 3
                  .Text = LS_Strarry(Ls_Count + 2)
                  
                  .Row = j
                  .Col = 4
                  .Text = LS_Strarry(Ls_Count + 3)
                  
                  .Row = j
                  .Col = 5
                  .Text = LS_Strarry(Ls_Count + 4)
                  
                  .Row = j
                  .Col = 6
                  .Text = LS_Strarry(Ls_Count + 5)
                  'Debug.Print Ls_Count
                   Ls_Count = Ls_Count + 5
                 
                   Ls_Count = Ls_Count + 1
                 
           Next i
 
      End With
      
      Ls_Count = 73
      j = 0
      
      With Me.ActiveForm.Spr_C

           For i = 1 To .MaxRows
               j = j + 1
               .Row = j
               .Col = 1
               .Text = LS_Strarry(Ls_Count)
               .Row = j
               .Col = 2
               .Text = LS_Strarry(Ls_Count + 1)
            
               .Row = j
               .Col = 3
               .Text = LS_Strarry(Ls_Count + 2)
             
               .Row = j
               .Col = 4
               .Text = LS_Strarry(Ls_Count + 3)
            
               .Row = j
               .Col = 5
               .Text = LS_Strarry(Ls_Count + 4)
            
               .Row = j
               .Col = 6
               .Text = LS_Strarry(Ls_Count + 5)
               .Row = j
                
               .Row = j
               .Col = 7
               .Text = LS_Strarry(Ls_Count + 6)
    
               .Row = j
               .Col = 8
               .Text = LS_Strarry(Ls_Count + 7)
               .Row = j

               .Row = j
               .Col = 9
               .Text = LS_Strarry(Ls_Count + 8)
               .Row = j

               .Row = j
               .Col = 10
               .Text = LS_Strarry(Ls_Count + 9)
               .Row = j
'
               .Row = j
               .Col = 11
               .Text = LS_Strarry(Ls_Count + 10)
               .Row = j
'
              '  Ls_Count = Ls_Count + 11
                Ls_Count = Ls_Count + 10
                Ls_Count = Ls_Count + 1
           Next i
      End With
   
End If

End Sub

'***********************************************************************************
'***  Description   :  폼 Activate 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************
Private Sub Mun_Save_Click()

Dim Ll_SprB_Count As Long
Dim Ll_SprBi_Count As Long
Dim Ll_SprC_Count As Long
Dim Ll_SprCi_Count As Long
Dim Ls_data As String
Dim LsData() As String
Dim Ls_MainData As String
Dim Ls_SprB_Data(6) As String       'Spr_B 셀 계산
Dim Ls_SprB_DataSave(16) As String  'Spr_B 셀 계산 모음
Dim Ls_SprC_Data(12) As String      'Spr_C 셀 계산
Dim Ls_SprC_DataSave(101) As String 'Spr_C 셀 계산 모음
Dim Ls_DataMain(2) As String        '셀 계산 통합(SPR_B,SPR_C)
Dim fileNumber As Integer
Dim strsize  As String

 Ll_SprB_Count = 0

 On Error GoTo ErrHandler

 CDlog.FileName = Me.ActiveForm.Caption
 CDlog.Filter = "파일(*.Han)|*.Han"
 CDlog.ShowSave

 With Me.ActiveForm.Spr_B

      For Ll_SprBi_Count = 1 To .MaxRows
        
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
            
          Ls_SprB_DataSave(Ll_SprB_Count) = Ls_SprB_Data(1) & "," & _
                                            Ls_SprB_Data(2) & "," & _
                                            Ls_SprB_Data(3) & "," & _
                                            Ls_SprB_Data(4) & "," & _
                                            Ls_SprB_Data(5) & "," & _
                                            Ls_SprB_Data(6) & ","
      Next Ll_SprBi_Count
        
 End With
   
 Ls_DataMain(1) = Ls_SprB_DataSave(1) & _
                  Ls_SprB_DataSave(2) & _
                  Ls_SprB_DataSave(3) & _
                  Ls_SprB_DataSave(4) & _
                  Ls_SprB_DataSave(5) & _
                  Ls_SprB_DataSave(6) & _
                  Ls_SprB_DataSave(7) & _
                  Ls_SprB_DataSave(8) & _
                  Ls_SprB_DataSave(9) & _
                  Ls_SprB_DataSave(10) & _
                  Ls_SprB_DataSave(11) & _
                  Ls_SprB_DataSave(12)

 With Me.ActiveForm.Spr_C

      For Ll_SprCi_Count = 1 To .MaxRows
       
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
          
         
          Ls_SprC_DataSave(Ll_SprC_Count) = Ls_SprC_Data(1) & "," & _
                                            Ls_SprC_Data(2) & "," & _
                                            Ls_SprC_Data(3) & "," & _
                                            Ls_SprC_Data(4) & "," & _
                                            Ls_SprC_Data(5) & "," & _
                                            Ls_SprC_Data(6) & "," & _
                                            Ls_SprC_Data(7) & "," & _
                                            Ls_SprC_Data(8) & "," & _
                                            Ls_SprC_Data(9) & "," & _
                                            Ls_SprC_Data(10) & "," & _
                                            Ls_SprC_Data(11) & ","
                                               
                  '   Debug.Print Ls_SprC_DataSave(Ll_SprC_Count)
                     
          Ls_DataMain(2) = Ls_DataMain(2) + Ls_SprC_DataSave(Ll_SprC_Count)
         
      Next Ll_SprCi_Count
        
 End With


 fileNumber = FreeFile

 Open CDlog.FileName For Output As #fileNumber
   
      Ls_MainData = Me.ActiveForm.cbo_Port.Text & "," & Me.ActiveForm.Cbo_Baud.Text & "," & _
                    Me.ActiveForm.Cbo_Dpi.Text & "," & Me.ActiveForm.Txt_CenterX.Text & "," & _
                    Me.ActiveForm.Txt_CenterY.Text & "," & Me.ActiveForm.Cbo_PrinterSpeed.Text & "," & _
                    Me.ActiveForm.Cbo_HeadDarkness.Text & "," & Ls_DataMain(1) & Ls_DataMain(2)
    Debug.Print Ls_MainData
 Print #fileNumber, Ls_MainData
 Close #fileNumber


 Me.ActiveForm.Caption = CDlog.FileTitle

ErrHandler:
 
End Sub

'***********************************************************************************
'***  Description   :  Color폼 로드 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_SettingColor_Click()
 
 Frm_Setting.Show 0

End Sub

Private Sub Mun_SettingSize_Click()
Frm_Size.Show 0
End Sub

'***********************************************************************************
'***  Description   : Tool Bar Visible 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Mun_Tool_Click()

 If Mun_Tool.Checked = True Then
       
       tlbMain.Visible = False
       Mun_Tool.Checked = False
 
 Else
       
       tlbMain.Visible = True
       Mun_Tool.Checked = True
 
 End If

End Sub

'***********************************************************************************
'***  Description   :  tlbMain Click 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

 Select Case Button.Index
        
        Case 2
               Mun_New_Click
        Case 3
               Mun_Open_Click
        Case 4
               Mun_Save_Click
 End Select

End Sub
