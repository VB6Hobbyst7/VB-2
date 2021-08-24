VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form FormNotes 
   Caption         =   "Important Notes..."
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   5190
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   18
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   6075
      FormWidthDT     =   8175
      FormScaleHeightDT=   5670
      FormScaleWidthDT=   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6900
      TabIndex        =   1
      Top             =   5310
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Important Notes On Using ActiveResize Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5205
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7935
      Begin VB.TextBox Text1 
         BackColor       =   &H00E1FFFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4905
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   7815
      End
   End
End
Attribute VB_Name = "FormNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    txt = "1. Never change the screen resolution when your project is open in the VB IDE. If you do so, this will likely cause VB to report incorrect screen height and / or width to ActiveResize, causing it to store these incorrect values, resulting in improper functionality. If you want to test your application, compile it to EXE, close your project and the VB IDE, and run the EXE in the resolution your like to test your application in." & vbCrLf & vbCrLf
    txt = txt & "2.  Make sure that the ScaleMode property of your forms and any controls that support this property is set to '1-Twip' (VB default)." & vbCrLf & vbCrLf
    txt = txt & "3.  When designing your forms in the VB IDE (VB design environment), keep in mind that your forms will be shown at run-time as they appear at design-time. That is, at design-time, your form has to show in its normal height and width, and no controls should be located beyond the form edges (hidden controls). This is because ActiveResize uses your form's design-time dimensions to calculate its new dimensions at run-time when it detects a different screen resolution or when the user resizes your form." & vbCrLf & vbCrLf
    txt = txt & "4.  It is recommended not to change the form size in the Form_Load() event. If this is necessary, then use the FormX.Height=X twice and then FormX.Width=Y once (or vice-versa), or use the FormX.Move method twice. This is necessary because the first code line that tries to resize the form (e.g Formx.Height=x) is ignored by ActiveResize when it is made in the Form_Load() event." & vbCrLf & vbCrLf
    txt = txt & "5.  Always use a True-Type font for your controls captions, text, etc., such as Tahoma, Verdana, Times New Roman, etc. The MS Sans Serif which is the default form font is not a true-type font and will not show nice when the form is resized." & vbCrLf & vbCrLf
    txt = txt & "6.  If you use Sheridan (Infragistics) grids on your form, it is recommended that you set the HideControlsOnResize property of ActiveResize to True. This will greatly enhance the resizing process speed."
    Text1 = txt

End Sub
