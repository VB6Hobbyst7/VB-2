VERSION 4.00
Begin VB.Form FrmLOGO 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4500
   ClientLeft      =   2430
   ClientTop       =   1935
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   4905
   Left            =   2370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Top             =   1590
   Width           =   6105
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   -30
      Picture         =   "IIMAGE16.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   -30
      Width           =   6060
   End
   Begin VB.PictureBox Picture2 
      Height          =   4545
      Left            =   -30
      ScaleHeight     =   299
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   1
      Top             =   -30
      Width           =   6015
   End
End
Attribute VB_Name = "FrmLOGO"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit









Private Sub Form_Load()
    Dim nRanDom As Single
    Dim X As Integer, k%
            
            
    Me.Show
    
    Picture1.Visible = False
            
    Randomize
    nRanDom = Int((5 * Rnd) + 1)
    
    Picture2.Cls
    
    Select Case nRanDom
    
        Case 1: Call BlindVert2Pass(4, 10000, FrmLOGO.Picture1, FrmLOGO.Picture2)

        Case 2: Call BlocksRandom(10, 400, FrmLOGO.Picture1, FrmLOGO.Picture2)
        
        Case 3

            FrmLOGO.Picture2.Width = FrmLOGO.Picture1.Width
            FrmLOGO.Picture2.Height = FrmLOGO.Picture1.Height
            
            X = Picture2.ScaleHeight / 2
            
            Call DropPaint(X, 10, FrmLOGO.Picture1, FrmLOGO.Picture2)
            
        Case 4: Call Diamond(4, 4, 8000, FrmLOGO.Picture1, FrmLOGO.Picture2)
        
        Case 5
            Call VenetianBlindVert(12, 10000, FrmLOGO.Picture1, FrmLOGO.Picture2)
            Call delay(70000)
                
    End Select
    
    
   'Unload Me
    
End Sub

