VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form FormMain 
   Caption         =   "ActiveResize Demo"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6690
      Top             =   3780
   End
   Begin VB.CommandButton cmdNotes 
      Caption         =   "Important Notes..."
      Height          =   405
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3180
      Width           =   1570
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Demo..."
      Height          =   405
      Left            =   60
      TabIndex        =   10
      Top             =   3180
      Width           =   1570
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   6030
      TabIndex        =   9
      Top             =   3180
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected Demo Description"
      Height          =   3045
      Left            =   3330
      TabIndex        =   7
      Top             =   60
      Width           =   4305
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   3720
         Top             =   2490
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   18
         ScreenHeight    =   1024
         ScreenWidth     =   1280
         ScreenHeightDT  =   1024
         ScreenWidthDT   =   1280
         AutoCenterForm  =   -1  'True
         FormHeightDT    =   4020
         FormWidthDT     =   7800
         FormScaleHeightDT=   3615
         FormScaleWidthDT=   7680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000080&
         Height          =   2625
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   4005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Demo"
      Height          =   3045
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3195
      Begin VB.OptionButton Option1 
         Caption         =   "Form Background Picture Demo"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Top             =   2700
         Width           =   2865
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Run-Time Controls Demo"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Top             =   2310
         Width           =   2865
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Data Grid Resize Demo"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   1920
         Width           =   2865
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sample Chart Resize Demo"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   1515
         Width           =   2865
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sample Calculator Resize Demo"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   2865
      End
      Begin VB.OptionButton Option1 
         Caption         =   "VB Controls Resize Demo - B"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   2865
      End
      Begin VB.OptionButton Option1 
         Caption         =   "VB Controls Resize Demo - A"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3330
      Picture         =   "FormMain.frx":08CA
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   2625
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'-------------------------------------------------------------------------
'                   Not even a single line of code!!!
'-------------------------------------------------------------------------
'
'   The code in this form is to run the demo application only.
'   You do not need to write any code in your forms.
    

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNotes_Click()
    Timer1.Enabled = False
    cmdNotes.BackColor = &H8000000F
    FormNotes.Show
End Sub

'The follwoing code shows the various demo forms
Private Sub cmdRun_Click()
    
    If Option1(0).Value Then
        VBControls.Show
    ElseIf Option1(1).Value Then
        VBControls2.Show
    ElseIf Option1(2).Value Then
        Calculator.Show
    ElseIf Option1(3).Value Then
        Chart.Show
    ElseIf Option1(4).Value Then
        DataGrid.Show
    ElseIf Option1(5).Value Then
        RunTimeControls.Show
    ElseIf Option1(6).Value Then
        BackgroundPicture.Show
    End If

End Sub

Private Sub Form_Load()
    Option1_Click (0)
    Timer1.Enabled = True
    
    'Uncomment the following line to enable restoring the form size/position
    'from a persistently saved state
    'ActiveResize1.RestoreForm Me, True, App.ProductName

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Uncomment the following line to enable saving the form size/position
    'persistently
    'ActiveResize1.SaveForm Me, True, App.ProductName

End Sub

'This is only to explain what each demo does...
Private Sub Option1_Click(Index As Integer)
    
    Select Case Index
        Case Is = 0
            Label1 = "This demo shows how ActiveResize can resize and reposition " _
            & "a form containing various control types including Frames, PictureBox " _
            & "controls, Buttons, Text boxes, List boxes, Check boxes and Labels. " _
            & " It also shows how to save the form state (its size / position and " _
            & "size / position of all form's controls) and then restore the form " _
            & "at any later time to this exact state. " _
            & " In addition, it shows how ActiveResize auto-resizes the form on load " _
            & "and centers it on the screen. ActiveResize also resizes the captions " _
            & "and text of all controls. The controls are set hidden when resizing the form " _
            & "(by setting the HideControlsOnResize to True) to boost the resizing process speed."
        Case Is = 1
            Label1 = "This demo shows how ActiveResize handles the complex SSTab control " _
            & "and all the controls contained within it, including Drive, Dir and " _
            & "File list boxes, Combo boxes and List boxes, Picture boxes and Frames, " _
            & "all within the tabs of the SSTab control. It also shows how ActiveResize " _
            & "can prevent resizing the form to less than 2/3 of its original size."
        Case Is = 2
            Label1 = "This demo shows a 'dummy' calculator. The form can not be maximized " _
            & "and can not be resized more than 50% of its original size."
        Case Is = 3
            Label1 = "This demo shows how ActiveResize handles Chart controls."
        Case Is = 4
            Label1 = "This demo shows how ActiveResize can efficiently resize a Data Grid " _
            & "including all of its columns heights and widths as well as its fonts. " _
            & "You can tell ActiveResize not to resize the grid columns (and fonts) when the " _
            & "grid is resized, so that more columns are displayed when the grid becomes wider. " _
            & "Just tick the check box in the next form before you resize the form."
        Case Is = 5
            Label1 = "This demo shows how ActiveResize can handle new control array " _
            & "elements that are created at run-time, with only a single line of code!... " _
            & "just a call to the method 'Reset'..."
        Case Is = 6
            Label1 = "This demo shows how ActiveResize can resize the background picture " _
            & "of the form when it loads, maintaining high picture quality when resized. ActiveResize " _
            & "automatically resizes the background picture of the form so that it fills the entire form " _
            & "even if the original picture is smaller or larger than the form size! " _
            & "Notice how ActiveResize resizes the background picture with 100% filcker-free processing when the form is resized!"
    
    End Select
        
End Sub

Private Sub Timer1_Timer()
    
    With cmdNotes
        If .BackColor = &H8000000F Then
            .BackColor = &H80FF&
        Else
            .BackColor = &H8000000F
        End If
    End With
End Sub
