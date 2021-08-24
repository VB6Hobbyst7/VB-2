VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHelpSpec 
   Caption         =   "Specode List"
   ClientHeight    =   2760
   ClientLeft      =   3570
   ClientTop       =   2880
   ClientWidth     =   3165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3165
   Begin VB.ListBox lstSpecode 
      Height          =   1860
      ItemData        =   "frmHelpSpec.frx":0000
      Left            =   60
      List            =   "frmHelpSpec.frx":000D
      TabIndex        =   0
      Top             =   120
      Width           =   2955
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   495
      Left            =   1620
      TabIndex        =   2
      Top             =   2100
      Width           =   1395
      Caption         =   "Exit"
      PicturePosition =   327683
      Size            =   "2461;873"
      Picture         =   "frmHelpSpec.frx":0045
      FontName        =   "±º∏≤"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSelect 
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   2100
      Width           =   1455
      BackColor       =   12632256
      Caption         =   "º±≈√"
      PicturePosition =   327683
      Size            =   "2566;873"
      Picture         =   "frmHelpSpec.frx":091F
      FontName        =   "±º∏≤"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmHelpSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Public Sub cmdSelect_Click()
    
    If lstSpecode.ListIndex = -1 Then
        lstSpecode.ListIndex = 0
    End If
    
    Call ClearForm(frmSpec12)
    
    frmSpec12.txtCodegu.Text = Left(lstSpecode.Text, 2)
    frmSpec12.txtCodeName.Text = Mid(lstSpecode.Text, 5, Len(lstSpecode.Text) - 4)
    
    Unload Me
    
End Sub

Private Sub lstSpecode_DblClick()

    Call cmdSelect_Click
    
End Sub

