VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PageSetup 
   Caption         =   "Setup"
   ClientHeight    =   3780
   ClientLeft      =   4275
   ClientTop       =   2235
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   4455
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      Top             =   3285
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   3285
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Page Margins (inch)"
      Height          =   975
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   4035
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   2
         Left            =   2700
         TabIndex        =   13
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   14
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         Caption         =   "Bottom:"
         Height          =   255
         Index           =   3
         Left            =   1980
         TabIndex        =   7
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         Caption         =   "Top:"
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   6
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         Caption         =   "Right:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         Caption         =   "Left:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview Zoom"
      Height          =   855
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   2265
      Visible         =   0   'False
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1080
         Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
         TabIndex        =   8
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         Caption         =   "Zoom:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page Orientation"
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      Begin VB.OptionButton porientation 
         Caption         =   "Landscape"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   16
         Top             =   480
         Width           =   1260
      End
      Begin VB.OptionButton porientation 
         Caption         =   "Portrait"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   1
         Left            =   2100
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   480
         Top             =   360
         Width           =   405
      End
   End
End
Attribute VB_Name = "PageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click(Index As Integer)
    'OK button
    If Index = 0 Then
        'GetZoom Combo1.ListIndex
        
        '** Ãâ·ÂÆû ¼±ÅÃ(Æû ÀÌ¸§)==============================================
        Select Case Gbl_FormName
            Case "frmConsum000r"
                With frmConsum000r
                    'Update margins
                    .spConsum000.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum000.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum000.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum000.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum000.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum000.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum001r"
                With frmConsum001r
                    'Update margins
                    .spConsum001.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum001.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum001.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum001.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum001.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum001.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum002r"
                With frmConsum002r
                    'Update margins
                    .spConsum002.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum002.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum002.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum002.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum002.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum002.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum003r"
                With frmConsum003r
                    'Update margins
                    .spConsum003.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum003.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum003.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum003.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum003.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum003.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum004r"
                With frmConsum004r
                    'Update margins
                    .spConsum004.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum004.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum004.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum004.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum004.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum004.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum005r"
                With frmConsum005r
                    'Update margins
                    .spConsum005.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum005.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum005.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum005.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum005.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum005.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum006r"
                With frmConsum006r
                    'Update margins
                    .spConsum006.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum006.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum006.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum006.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum006.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum006.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmConsum007r"
                With frmConsum007r
                    'Update margins
                    .spConsum007.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spConsum007.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spConsum007.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spConsum007.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spConsum007.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spConsum007.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmHistory000r"
                With frmHistory000r
                    'Update margins
                    .spHistory000.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spHistory000.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spHistory000.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spHistory000.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spHistory000.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spHistory000.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmHistory001r"
                With frmHistory001r
                    'Update margins
                    .spHistory001.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spHistory001.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spHistory001.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spHistory001.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spHistory001.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spHistory001.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmHistory002r"
                With frmHistory002r
                    'Update margins
                    .spHistory002.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spHistory002.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spHistory002.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spHistory002.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spHistory002.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spHistory002.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
            Case "frmHistory003r"
                With frmHistory003r
                    'Update margins
                    .spHistory003.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
                    .spHistory003.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
                    .spHistory003.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
                    .spHistory003.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
                    
                    'Change the page orientation
                    'Portrait
                    If porientation(0).Value = True Then
                        .spHistory003.PrintOrientation = PrintOrientationPortrait
                    'Landscape
                    Else
                        .spHistory003.PrintOrientation = PrintOrientationLandscape
                    End If
                End With
                
        End Select
        '=====================================================================
        
        'set zoom attributes
        zoomindex = Combo1.ListIndex
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    '** Ãâ·ÂÆû ¼±ÅÃ(Æû ÀÌ¸§)==============================================
    Select Case Gbl_FormName
        Case "frmStatistics"
            With frmStatistics
                'Get page margins (convert to inches) and format
                pagemargin(0).Text = Format(.spdResult1.PrintMarginLeft / 1440, "0.00")
                pagemargin(1).Text = Format(.spdResult1.PrintMarginRight / 1440, "0.00")
                pagemargin(2).Text = Format(.spdResult1.PrintMarginTop / 1440, "0.00")
                pagemargin(3).Text = Format(.spdResult1.PrintMarginBottom / 1440, "0.00")
                
                'Get page orientation
                If .spdResult1.PrintOrientation = PrintOrientationLandscape Then
                    porientation(1) = True
                Else
                    porientation(0) = True
                End If
            End With
            
        Case Else
            
    End Select
    '=====================================================================
    
    'Populate Zooming combobox
    Combo1.AddItem "200%"
    Combo1.AddItem "150%"
    Combo1.AddItem "100%"
    Combo1.AddItem "75%"
    Combo1.AddItem "50%"
    Combo1.AddItem "25%"
    Combo1.AddItem "10%"
    Combo1.AddItem "Page Width"
    Combo1.AddItem "Page Height"
    Combo1.AddItem "Whole Page"
    Combo1.AddItem "Two Pages"
    Combo1.AddItem "Three Pages"
    Combo1.AddItem "Four Pages"
    Combo1.AddItem "Six Pages"
    
    'Get the zoom display
    'Combo1.ListIndex = zoomindex
    
End Sub

