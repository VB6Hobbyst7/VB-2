VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   645
      Left            =   660
      TabIndex        =   1
      Top             =   1410
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Private Sub Command1_Click()
    Call RotateControl(Label1, 90)
End Sub
Private Sub RotateControl(ctl As Control, intAngle As Integer)
    Dim lnghNewFont As Long
    Dim lnghOriginalFonrt As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    
    With Me
        .ScaleMode = vbPixels
        .AutoRedraw = True
        lngHeight = .TextHeight(ctl)
        lngWidth = 0
        
        With .Font
            lnghNewFont = CreateFont(lngHeight, lngWidth, intAngle * 10, intAngle * 10, .Weight, .Italic, .Underline, .Strikethrough, .Charset, 0, 0, 0, 0, .Name)
        End With
        lnghOriginalFonrt = SelectObject(.hdc, lnghNewFont)
        .CurrentX = ctl.Left
        .CurrentY = ctl.Top
        Me.Print ctl
    
        lnghNewFont = SelectObject(.hdc, lnghOriginalFonrt)
        .AutoRedraw = False
    End With
    DeleteObject lnghNewFont
    ctl.Visible = False
End Sub
