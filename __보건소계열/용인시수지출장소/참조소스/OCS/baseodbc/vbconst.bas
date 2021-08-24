Attribute VB_Name = "VbConstant"
''''''''''''''''''''''''''''
' Visual Basic public constant file. This file can be loaded
' into a code module.
'
' Some constants are commented out because they have
' duplicates (e.g., NONE appears several places).
'
' If you are updating a Visual Basic application written with
' an older version, you should replace your public constants
' with the constants in this file.
'
''''''''''''''''''''''''''''
'Public Declare Function GetActiveWindow Lib "user32" () As Long

'아래 함수 최상위 윈도우를 만들기 위해 사용한다.
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal insaft As Long, _
                 ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal Flags As Long) As Long
'아래 레지스트리 설정할수 있다.1
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
   (ByVal hKey As Long, _
    ByVal lpszValueName As String, _
    ByVal dwReserved As Long, _
    ByVal fdwType As Long, _
    lpbData As Any, _
    ByVal cbData As Long) _
    As Long
'아래 레지스트리 설정할수 있다.2
Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" _
   (ByVal hKey As Long, _
    ByVal lpszSubKey As String, _
    phkResult As Long) _
    As Long
'아래 레지스트리 설정할수 있다.3
Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) _
    As Long
    
                 
'Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, _
'                 ByVal fsModifiers As Long, ByVal vk As Long) As Long

' Clipboard formats
Public Const CF_LINK = &HBF00
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILE = 3
Public Const CF_DIB = 8
Public Const CF_PALETTE = 9

' DragOver
Public Const ENTER = 0
Public Const LEAVE = 1
Public Const OVER = 2

' Drag (controls)
Public Const Cancel = 0
Public Const BEGIN_DRAG = 1
Public Const END_DRAG = 2

' Show parameters
Public Const MODAL = 1
Public Const MODELESS = 0

' Arrange Method
' for MDI Forms
Public Const CASCADE = 0
Public Const TILE_HORIZONTAL = 1
Public Const TILE_VERTICAL = 2
Public Const ARRANGE_ICONS = 3

'ZOrder Method
Public Const BRINGTOFRONT = 0
Public Const SENDTOBACK = 1

' Key Codes
Public Const KEY_LBUTTON = &H1
Public Const KEY_RBUTTON = &H2
Public Const KEY_CANCEL = &H3
Public Const KEY_MBUTTON = &H4    ' NOT contiguous with L & RBUTTON
Public Const KEY_BACK = &H8
Public Const KEY_TAB = &H9
Public Const KEY_CLEAR = &HC
Public Const KEY_RETURN = &HD
Public Const KEY_SHIFT = &H10
Public Const KEY_CONTROL = &H11
Public Const KEY_MENU = &H12
Public Const KEY_PAUSE = &H13
Public Const KEY_CAPITAL = &H14
Public Const KEY_ESCAPE = &H1B
Public Const KEY_SPACE = &H20
Public Const KEY_PRIOR = &H21
Public Const KEY_NEXT = &H22
Public Const KEY_END = &H23
Public Const KEY_HOME = &H24
Public Const KEY_LEFT = &H25
Public Const KEY_UP = &H26
Public Const KEY_RIGHT = &H27
Public Const KEY_DOWN = &H28
Public Const KEY_SELECT = &H29
Public Const KEY_PRINT = &H2A
Public Const KEY_EXECUTE = &H2B
Public Const KEY_SNAPSHOT = &H2C
Public Const KEY_INSERT = &H2D
Public Const KEY_DELETE = &H2E
Public Const KEY_HELP = &H2F

' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'

Public Const KEY_NUMPAD0 = &H60
Public Const KEY_NUMPAD1 = &H61
Public Const KEY_NUMPAD2 = &H62
Public Const KEY_NUMPAD3 = &H63
Public Const KEY_NUMPAD4 = &H64
Public Const KEY_NUMPAD5 = &H65
Public Const KEY_NUMPAD6 = &H66
Public Const KEY_NUMPAD7 = &H67
Public Const KEY_NUMPAD8 = &H68
Public Const KEY_NUMPAD9 = &H69
Public Const KEY_MULTIPLY = &H6A
Public Const KEY_ADD = &H6B
Public Const KEY_SEPARATOR = &H6C
Public Const KEY_SUBTRACT = &H6D
Public Const KEY_DECIMAL = &H6E
Public Const KEY_DIVIDE = &H6F
Public Const KEY_F1 = &H70
Public Const KEY_F2 = &H71
Public Const KEY_F3 = &H72
Public Const KEY_F4 = &H73
Public Const KEY_F5 = &H74
Public Const KEY_F6 = &H75
Public Const KEY_F7 = &H76
Public Const KEY_F8 = &H77
Public Const KEY_F9 = &H78
Public Const KEY_F10 = &H79
Public Const KEY_F11 = &H7A
Public Const KEY_F12 = &H7B
Public Const KEY_F13 = &H7C
Public Const KEY_F14 = &H7D
Public Const KEY_F15 = &H7E
Public Const KEY_F16 = &H7F

Public Const KEY_NUMLOCK = &H90

' Event Parameters

' ErrNum (LinkError)
Public Const WRONG_FORMAT = 1
Public Const DDE_SOURCE_CLOSED = 6
Public Const TOO_MANY_LINKS = 7
Public Const DATA_TRANSFER_FAILED = 8

' QueryUnload
Public Const FORM_CONTROLMENU = 0
Public Const FORM_CODE = 1
Public Const APP_WINDOWS = 2
Public Const APP_TASKMANAGER = 3
Public Const FORM_MDIFORM = 4

' Properties

' Colors
Public Const BLACK = &H0&
Public Const RED = &HFF&
Public Const GREEN = &HFF00&
Public Const YELLOW = &HFFFF&
Public Const BLUE = &HFF0000
Public Const MAGENTA = &HFF00FF
Public Const CYAN = &HFFFF00
Public Const WHITE = &HFFFFFF

' System Colors
Public Const SCROLL_BARS = &H80000000           ' Scroll-bars gray area.
Public Const DESKTOP = &H80000001               ' Desktop.
Public Const ACTIVE_TITLE_BAR = &H80000002      ' Active window caption.
Public Const INACTIVE_TITLE_BAR = &H80000003    ' Inactive window caption.
Public Const MENU_BAR = &H80000004              ' Menu background.
Public Const WINDOW_BACKGROUND = &H80000005     ' Window background.
Public Const WINDOW_FRAME = &H80000006          ' Window frame.
Public Const MENU_TEXT = &H80000007             ' Text in menus.
Public Const WINDOW_TEXT = &H80000008           ' Text in windows.
Public Const TITLE_BAR_TEXT = &H80000009        ' Text in caption, size box, scroll-bar arrow box..
Public Const ACTIVE_BORDER = &H8000000A         ' Active window border.
Public Const INACTIVE_BORDER = &H8000000B       ' Inactive window border.
Public Const APPLICATION_WORKSPACE = &H8000000C ' Background color of multiple document interface (MDI) applications.
Public Const HIGHLIGHT = &H8000000D             ' Items selected item in a control.
Public Const HIGHLIGHT_TEXT = &H8000000E        ' Text of item selected in a control.
Public Const BUTTON_FACE = &H8000000F           ' Face shading on command buttons.
Public Const BUTTON_SHADOW = &H80000010         ' Edge shading on command buttons.
Public Const GRAY_TEXT = &H80000011             ' Grayed (disabled) text.  This color is set to 0 if the current display driver does not support a solid gray color.
Public Const BUTTON_TEXT = &H80000012           ' Text on push buttons.

' Enumerated Types

' Align (picture box)
Public Const NONE = 0
Public Const ALIGN_TOP = 1
Public Const ALIGN_BOTTOM = 2

' Alignment
Public Const LEFT_JUSTIFY = 0  ' 0 - Left Justify
Public Const RIGHT_JUSTIFY = 1 ' 1 - Right Justify
Public Const CENTER = 2        ' 2 - Center

' BorderStyle (form)
'public Const NONE = 0          ' 0 - None
Public Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Public Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Public Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

' MousePointer
Public Const DEFAULT = 0        ' 0 - Default
Public Const ARROW = 1          ' 1 - Arrow
Public Const CROSSHAIR = 2      ' 2 - Cross
Public Const IBEAM = 3          ' 3 - I-Beam
Public Const ICON_POINTER = 4   ' 4 - Icon
Public Const SIZE_POINTER = 5   ' 5 - Size
Public Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Public Const SIZE_N_S = 7       ' 7 - Size N S
Public Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Public Const SIZE_W_E = 9       ' 9 - Size W E
Public Const UP_ARROW = 10      ' 10 - Up Arrow
Public Const HOURGLASS = 11     ' 11 - Hourglass
Public Const NO_DROP = 12       ' 12 - No drop

' DragMode
Public Const MANUAL = 0    ' 0 - Manual
Public Const AUTOMATIC = 1 ' 1 - Automatic

' DrawMode
Public Const BLACKNESS = 1      ' 1 - Blackness
Public Const NOT_MERGE_PEN = 2  ' 2 - Not Merge Pen
Public Const MASK_NOT_PEN = 3   ' 3 - Mask Not Pen
Public Const NOT_COPY_PEN = 4   ' 4 - Not Copy Pen
Public Const MASK_PEN_NOT = 5   ' 5 - Mask Pen Not
Public Const INVERT = 6         ' 6 - Invert
Public Const XOR_PEN = 7        ' 7 - Xor Pen
Public Const NOT_MASK_PEN = 8   ' 8 - Not Mask Pen
Public Const MASK_PEN = 9       ' 9 - Mask Pen
Public Const NOT_XOR_PEN = 10   ' 10 - Not Xor Pen
Public Const NOP = 11           ' 11 - Nop
Public Const MERGE_NOT_PEN = 12 ' 12 - Merge Not Pen
Public Const COPY_PEN = 13      ' 13 - Copy Pen
Public Const MERGE_PEN_NOT = 14 ' 14 - Merge Pen Not
Public Const MERGE_PEN = 15     ' 15 - Merge Pen
Public Const WHITENESS = 16     ' 16 - Whiteness

' DrawStyle
Public Const SOLID = 0        ' 0 - Solid
Public Const DASH = 1         ' 1 - Dash
Public Const DOT = 2          ' 2 - Dot
Public Const DASH_DOT = 3     ' 3 - Dash-Dot
Public Const DASH_DOT_DOT = 4 ' 4 - Dash-Dot-Dot
Public Const INVISIBLE = 5    ' 5 - Invisible
Public Const INSIDE_SOLID = 6 ' 6 - Inside Solid

' FillStyle
' public Const SOLID = 0           ' 0 - Solid
Public Const TRANSPARENT = 1       ' 1 - Transparent
Public Const HORIZONTAL_LINE = 2   ' 2 - Horizontal Line
Public Const VERTICAL_LINE = 3     ' 3 - Vertical Line
Public Const UPWARD_DIAGONAL = 4   ' 4 - Upward Diagonal
Public Const DOWNWARD_DIAGONAL = 5 ' 5 - Downward Diagonal
Public Const CROSS = 6             ' 6 - Cross
Public Const DIAGONAL_CROSS = 7    ' 7 - Diagonal Cross

' ScaleMode
Public Const USER = 0        ' 0 - User
Public Const Twips = 1       ' 1 - Twip
Public Const POINTS = 2      ' 2 - Point
Public Const PIXELS = 3      ' 3 - Pixel
Public Const CHARACTERS = 4  ' 4 - Character
Public Const INCHES = 5      ' 5 - Inch
Public Const MILLIMETERS = 6 ' 6 - Millimeter
Public Const CENTIMETERS = 7 ' 7 - Centimeter

' ScrollBar
' public Const NONE     = 0 ' 0 - None
Public Const HORIZONTAL = 1 ' 1 - Horizontal
Public Const VERTICAL = 2   ' 2 - Vertical
Public Const BOTH = 3       ' 3 - Both

' WindowState
Public Const NORMAL = 0    ' 0 - Normal
Public Const MINIMIZED = 1 ' 1 - Minimized
Public Const MAXIMIZED = 2 ' 2 - Maximized

' Check Value
Public Const UNCHECKED = 0 ' 0 - Unchecked
Public Const CHECKED = 1   ' 1 - Checked
Public Const GRAYED = 2    ' 2 - Grayed

' Shift parameter masks
Public Const SHIFT_MASK = 1
Public Const CTRL_MASK = 2
Public Const ALT_MASK = 4

' Button parameter masks
Public Const LEFT_BUTTON = 1
Public Const RIGHT_BUTTON = 2
Public Const MIDDLE_BUTTON = 4

' Function Parameters
' MsgBox parameters
Public Const MB_OK = 0                 ' OK button only
Public Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Public Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Public Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Public Const MB_YESNO = 4              ' Yes and No buttons
Public Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Public Const MB_ICONSTOP = 16          ' Critical message
Public Const MB_ICONQUESTION = 32      ' Warning query
Public Const MB_ICONEXCLAMATION = 48   ' Warning message
Public Const MB_ICONINFORMATION = 64   ' Information message

Public Const MB_APPLMODAL = 0          ' Application Modal Message Box
Public Const MB_DEFBUTTON1 = 0         ' First button is default
Public Const MB_DEFBUTTON2 = 256       ' Second button is default
Public Const MB_DEFBUTTON3 = 512       ' Third button is default
Public Const MB_SYSTEMMODAL = 4096      'System Modal

' MsgBox return values
Public Const IDOK = 1                  ' OK button pressed
Public Const IDCANCEL = 2              ' Cancel button pressed
Public Const IDABORT = 3               ' Abort button pressed
Public Const IDRETRY = 4               ' Retry button pressed
Public Const IDIGNORE = 5              ' Ignore button pressed
Public Const IDYES = 6                 ' Yes button pressed
Public Const IDNO = 7                  ' No button pressed

' SetAttr, Dir, GetAttr functions
Public Const ATTR_NORMAL = 0
Public Const ATTR_READONLY = 1
Public Const ATTR_HIDDEN = 2
Public Const ATTR_SYSTEM = 4
Public Const ATTR_VOLUME = 8
Public Const ATTR_DIRECTORY = 16
Public Const ATTR_ARCHIVE = 32


'Special Verb Values
Public Const VERB_PRIMARY = 0
Public Const VERB_SHOW = -1
Public Const VERB_OPEN = -2
Public Const VERB_HIDE = -3
Public Const VERB_INPLACEUIACTIVATE = -4
Public Const VERB_INPLACEACTIVATE = -5
'The last two verbs are for objects that support "inside-out" activation,
'meaning they can be edited in-place, and that they support being left
'in-place-active even when the input focus moves to another control or form.
'These objects actually have 2 levels of being active.  "InPlace Active"
'means that the object is ready for the user to click inside it and start
'working with it.  "In-Place UI-Active" means that, in addition, if the object
'has any other UI associated with it, such as floating palette windows,
'that those windows are visible and ready for use.  Any number of objects
'can be "In-Place Active" at a time, although only one can be
'"InPlace UI-Active".

'You can cause an object to move to either one of states programmatically by
'setting the Verb property to the appropriate verb and setting
'Action=OLE_ACTIVATE.

'Also, if you set AutoActivate = OLE_ACTIVATE_GETFOCUS, the server will
'automatically be put into "InPlace UI-Active" state when the user clicks
'on or tabs into the control.

'VerbFlag Bit Masks
Public Const VERBFLAG_GRAYED = &H1
Public Const VERBFLAG_DISABLED = &H2
Public Const VERBFLAG_CHECKED = &H8
Public Const VERBFLAG_SEPARATOR = &H800

'MiscFlag Bits - Or these together as desired for special behaviors

'MEMSTORAGE causes the control to use memory to store the object while
'           it is loaded.  This is faster than the default (disk-tempfile),
'           but can consume a lot of memory for objects whose data takes
'           up a lot of space, such as the bitmap for a paint program.
Public Const OLE_MISCFLAG_MEMSTORAGE = &H1

'DISABLEINPLACE overrides the control's default behavior of allowing
'           in-place activation for objects that support it.  If you
'           are having problems activating an object inplace, you can
'           force it to always activate in a separate window by setting this
'           bit
Public Const OLE_MISCFLAG_DISABLEINPLACE = &H2

'Common Dialog Control
'Action Property
Public Const DLG_FILE_OPEN = 1
Public Const DLG_FILE_SAVE = 2
Public Const DLG_COLOR = 3
Public Const DLG_FONT = 4
Public Const DLG_PRINT = 5
Public Const DLG_HELP = 6

'File Open/Save Dialog Flags
Public Const OFN_READONLY = &H1&
Public Const OFN_OVERWRITEPROMPT = &H2&
Public Const OFN_HIDEREADONLY = &H4&
Public Const OFN_NOCHANGEDIR = &H8&
Public Const OFN_SHOWHELP = &H10&
Public Const OFN_NOVALIDATE = &H100&
Public Const OFN_ALLOWMULTISELECT = &H200&
Public Const OFN_EXTENSIONDIFFERENT = &H400&
Public Const OFN_PATHMUSTEXIST = &H800&
Public Const OFN_FILEMUSTEXIST = &H1000&
Public Const OFN_CREATEPROMPT = &H2000&
Public Const OFN_SHAREAWARE = &H4000&
Public Const OFN_NOREADONLYRETURN = &H8000&

'Color Dialog Flags
Public Const CC_RGBINIT = &H1&
Public Const CC_FULLOPEN = &H2&
Public Const CC_PREVENTFULLOPEN = &H4&
Public Const CC_SHOWHELP = &H8&

'Fonts Dialog Flags
Public Const CF_SCREENFONTS = &H1&
Public Const CF_PRINTERFONTS = &H2&
Public Const CF_BOTH = &H3&
Public Const CF_SHOWHELP = &H4&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000&         'must also have CF_SCREENFONTS & CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000

'Printer Dialog Flags
Public Const PD_ALLPAGES = &H0&
Public Const PD_SELECTION = &H1&
Public Const PD_PAGENUMS = &H2&
Public Const PD_NOSELECTION = &H4&
Public Const PD_NOPAGENUMS = &H8&
Public Const PD_COLLATE = &H10&
Public Const PD_PRINTTOFILE = &H20&
Public Const PD_PRINTSETUP = &H40&
Public Const PD_NOWARNING = &H80&
Public Const PD_RETURNDC = &H100&
Public Const PD_RETURNIC = &H200&
Public Const PD_RETURNDEFAULT = &H400&
Public Const PD_SHOWHELP = &H800&
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000

'Help Constants
Public Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Public Const HELP_QUIT = &H2              'Terminate help
Public Const HELP_INDEX = &H3             'Display index
Public Const HELP_CONTENTS = &H3
Public Const HELP_HELPONHELP = &H4        'Display help on using help
Public Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Public Const HELP_SETCONTENTS = &H5
Public Const HELP_CONTEXTPOPUP = &H8
Public Const HELP_FORCEFILE = &H9
Public Const HELP_KEY = &H101             'Display topic for keyword in offabData
Public Const HELP_COMMAND = &H102
Public Const HELP_PARTIALKEY = &H105      'call the search engine in winhelp

'Error Constants
Public Const CDERR_DIALOGFAILURE = -32768

Public Const CDERR_GENERALCODES = &H7FFF
Public Const CDERR_STRUCTSIZE = &H7FFE
Public Const CDERR_INITIALIZATION = &H7FFD
Public Const CDERR_NOTEMPLATE = &H7FFC
Public Const CDERR_NOHINSTANCE = &H7FFB
Public Const CDERR_LOADSTRFAILURE = &H7FFA
Public Const CDERR_FINDRESFAILURE = &H7FF9
Public Const CDERR_LOADRESFAILURE = &H7FF8
Public Const CDERR_LOCKRESFAILURE = &H7FF7
Public Const CDERR_MEMALLOCFAILURE = &H7FF6
Public Const CDERR_MEMLOCKFAILURE = &H7FF5
Public Const CDERR_NOHOOK = &H7FF4

'Added for CMDIALOG.VBX
Public Const CDERR_CANCEL = &H7FF3
Public Const CDERR_NODLL = &H7FF2
Public Const CDERR_ERRPROC = &H7FF1
Public Const CDERR_ALLOC = &H7FF0
Public Const CDERR_HELP = &H7FEF

Public Const PDERR_PRINTERCODES = &H6FFF
Public Const PDERR_SETUPFAILURE = &H6FFE
Public Const PDERR_PARSEFAILURE = &H6FFD
Public Const PDERR_RETDEFFAILURE = &H6FFC
Public Const PDERR_LOADDRVFAILURE = &H6FFB
Public Const PDERR_GETDEVMODEFAIL = &H6FFA
Public Const PDERR_INITFAILURE = &H6FF9
Public Const PDERR_NODEVICES = &H6FF8
Public Const PDERR_NODEFAULTPRN = &H6FF7
Public Const PDERR_DNDMMISMATCH = &H6FF6
Public Const PDERR_CREATEICFAILURE = &H6FF5
Public Const PDERR_PRINTERNOTFOUND = &H6FF4

Public Const CFERR_CHOOSEFONTCODES = &H5FFF
Public Const CFERR_NOFONTS = &H5FFE

Public Const FNERR_FILENAMECODES = &H4FFF
Public Const FNERR_SUBCLASSFAILURE = &H4FFE
Public Const FNERR_INVALIDFILENAME = &H4FFD
Public Const FNERR_BUFFERTOOSMALL = &H4FFC

Public Const FRERR_FINDREPLACECODES = &H3FFF
Public Const CCERR_CHOOSECOLORCODES = &H2FFF


'-------------------------------------------------------------------
'3D Controls
'-------------------------------------------------------------------
'Alignment (Check Box)
Public Const SSCB_TEXT_RIGHT = 0         '0 - Text to the right
Public Const SSCB_TEXT_LEFT = 1          '1 - Text to the left

'Alignment (Option Button)
Public Const SSOB_TEXT_RIGHT = 0         '0 - Text to the right
Public Const SSOB_TEXT_LEFT = 1          '1 - Text to the left

'Alignment (Frame)
Public Const SSFR_LEFT_JUSTIFY = 0       '0 - Left justify text
Public Const SSFR_RIGHT_JUSTIFY = 1      '1 - Right justify text
Public Const SSFR_CENTER = 2             '2 - Center text

'Alignment (Panel)
Public Const SSPN_LEFT_TOP = 0           '0 - Text to left and top
Public Const SSPN_LEFT_MIDDLE = 1        '1 - Text to left and middle
Public Const SSPN_LEFT_BOTTOM = 2        '2 - Text to left and bottom
Public Const SSPN_RIGHT_TOP = 3          '3 - Text to right and top
Public Const SSPN_RIGHT_MIDDLE = 4       '4 - Text to right and middle
Public Const SSPN_RIGHT_BOTTOM = 5       '5 - Text to right and bottom
Public Const SSPN_CENTER_TOP = 6         '6 - Text to center and top
Public Const SSPN_CENTER_MIDDLE = 7      '7 - Text to center and middle
Public Const SSPN_CENTER_BOTTOM = 8      '8 - Text to center and bottom

'Autosize (Command Button)
'public Const SS_AUTOSIZE_NONE = 0        '0 - No Autosizing
Public Const SSPB_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Public Const SSPB_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Ribbon Button)
'public Const SS_AUTOSIZE_NONE      = 0  '0 - No Autosizing
Public Const SSRI_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Public Const SSRI_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Panel)
'public Const SS_AUTOSIZE_NONE    = 0    '0 - No Autosizing
Public Const SSPN_AUTOSIZE_WIDTH = 1     '1 - Autosize Panel width to Caption
Public Const SSPN_AUTOSIZE_HEIGHT = 2    '2 - Autosize Panel height to Caption
Public Const SSPN_AUTOSIZE_CHILD = 3     '3 - Autosize Child to Panel

'BevelInner (Panel)
Public Const SS_BEVELINNER_NONE = 0      '0 - No Inner Bevel
Public Const SS_BEVELINNER_INSET = 1     '1 - Inset Inner Bevel
Public Const SS_BEVELINNER_RAISED = 2    '2 - Raised Inner Bevel

'BevelOuter (Panel)
Public Const SS_BEVELOUTER_NONE = 0      '0 - No Outer Bevel
Public Const SS_BEVELOUTER_INSET = 1     '1 - Inset Outer Bevel
Public Const SS_BEVELOUTER_RAISED = 2    '2 - Raised Outer Bevel

'FloodType (Panel)
Public Const SS_FLOODTYPE_NONE = 0       '0 - No flood
Public Const SS_FLOODTYPE_L_TO_R = 1     '1 - Left to light
Public Const SS_FLOODTYPE_R_TO_L = 2     '2 - Right to left
Public Const SS_FLOODTYPE_T_TO_B = 3     '3 - Top to bottom
Public Const SS_FLOODTYPE_B_TO_T = 4     '4 - Bottom to top
Public Const SS_FLOODTYPE_CIRCLE = 5     '5 - Widening circle

'Font3D (Panel, Command Button, Option Button, Check Box, Frame)
Public Const SS_FONT3D_NONE = 0          '0 - No 3-D text
Public Const SS_FONT3D_RAISED_LIGHT = 1  '1 - Raised with light shading
Public Const SS_FONT3D_RAISED_HEAVY = 2  '2 - Raised with heavy shading
Public Const SS_FONT3D_INSET_LIGHT = 3   '3 - Inset with light shading
Public Const SS_FONT3D_INSET_HEAVY = 4   '4 - Inset with heavy shading

'PictureDnChange (Ribbon Button)
Public Const SS_PICDN_NOCHANGE = 0       '0 - Use 'Up'bitmap with no change
Public Const SS_PICDN_DITHER = 1         '1 - Dither 'Up'bitmap
Public Const SS_PICDN_INVERT = 2         '2 - Invert 'Up'bitmap

'ShadowColor (Panel, Frame)
Public Const SS_SHADOW_DARKGREY = 0      '0 - Dark grey shadow
Public Const SS_SHADOW_BLACK = 1         '1 - Black shadow

'ShadowStyle (Frame)
Public Const SS_SHADOW_INSET = 0         '0 - Shadow inset
Public Const SS_SHADOW_RAISED = 1        '1 - Shadow raised


'---------------------------------------
'GAUGE
'---------------------------------------
'Style Property
Public Const GAUGE_HORIZ = 0
Public Const GAUGE_VERT = 1
Public Const GAUGE_SEMI = 2
Public Const GAUGE_FULL = 3


'----------------------------------------
'Graph Control
'----------------------------------------
'General
Public Const G_NONE = 0
Public Const G_DEFAULT = 0

Public Const G_OFF = 0
Public Const G_ON = 1

Public Const G_MONO = 0
Public Const G_COLOR = 1

'Graph Types
Public Const G_PIE2D = 1
Public Const G_PIE3D = 2
Public Const G_BAR2D = 3
Public Const G_BAR3D = 4
Public Const G_GANTT = 5
Public Const G_LINE = 6
Public Const G_LOGLIN = 7
Public Const G_AREA = 8
Public Const G_SCATTER = 9
Public Const G_POLAR = 10
Public Const G_HLC = 11

'Colors
Public Const G_BLACK = 0
Public Const G_BLUE = 1
Public Const G_GREEN = 2
Public Const G_CYAN = 3
Public Const G_RED = 4
Public Const G_MAGENTA = 5
Public Const G_BROWN = 6
Public Const G_LIGHT_GRAY = 7
Public Const G_DARK_GRAY = 8
Public Const G_LIGHT_BLUE = 9
Public Const G_LIGHT_GREEN = 10
Public Const G_LIGHT_CYAN = 11
Public Const G_LIGHT_RED = 12
Public Const G_LIGHT_MAGENTA = 13
Public Const G_YELLOW = 14
Public Const G_WHITE = 15
Public Const G_AUTOBW = 16

'Patterns
Public Const G_SOLID = 0
Public Const G_HOLLOW = 1
Public Const G_HATCH1 = 2
Public Const G_HATCH2 = 3
Public Const G_HATCKF3_ = 4
Public Const G_HATCH4 = 5
Public Const G_HATCH5 = 6
Public Const G_HATCH6 = 7
Public Const G_BITMAP1 = 16
Public Const G_BITMAP2 = 17
Public Const G_BITMAP3 = 18
Public Const G_BITMAP4 = 19
Public Const G_BITMAP5 = 20
Public Const G_BITMAP6 = 21
Public Const G_BITMAP7 = 22
Public Const G_BITMAP8 = 23
Public Const G_BITMAP9 = 24
Public Const G_BITMAP10 = 25
Public Const G_BITMAP11 = 26
Public Const G_BITMAP12 = 27
Public Const G_BITMAP13 = 28
Public Const G_BITMAP14 = 29
Public Const G_BITMAP15 = 30
Public Const G_BITMAP16 = 31

'Symbols
Public Const G_CROSS_PLUS = 0
Public Const G_CROSS_TIMES = 1
Public Const G_TRIANGLE_UP = 2
Public Const G_SOLID_TRIANGLE_UP = 3
Public Const G_TRIANGLE_DOWN = 4
Public Const G_SOLID_TRIANGLE_DOWN = 5
Public Const G_SQUARE = 6
Public Const G_SOLID_SQUARE = 7
Public Const G_DIAMOND = 8
Public Const G_SOLID_DIAMOND = 9

'Line Styles
'public Const G_SOLID = 0
Public Const G_DASH = 1
Public Const G_DOT = 2
Public Const G_DASHDOT = 3
Public Const G_DASHDOTDOT = 4

'Grids
Public Const G_HORIZONTAL = 1
Public Const G_VERTICAL = 2

'Statistics
Public Const G_MEAN = 1
Public Const G_MIN_MAX = 2
Public Const G_STD_DEV = 4
Public Const G_BEST_FIT = 8

'Data Arrays
Public Const G_GRAPH_DATA = 1
Public Const G_COLOR_DATA = 2
Public Const G_EXTRA_DATA = 3
Public Const G_Panel_TEXT = 4
Public Const G_LEGEND_TEXT = 5
Public Const G_PATTERN_DATA = 6
Public Const G_SYMBOL_DATA = 7
Public Const G_XPOS_DATA = 8
Public Const G_ALL_DATA = 9

'Draw Mode
Public Const G_NO_ACTION = 0
Public Const G_CLEAR = 1
Public Const G_DRAW = 2
Public Const G_BLIT = 3
Public Const G_COPY = 4
Public Const G_PRINT = 5
Public Const G_WRITE = 6

'Print Options
Public Const G_BORDER = 2

'Pie   Options             '
Public Const G_NO_LINES = 1
Public Const G_COLORED = 2
Public Const G_PERCENTS = 4

'Bar Chart Options             '
'public Const G_HORIZONTAL = 1
Public Const G_STACKED = 2
Public Const G_PERCENTAGE = 4
Public Const G_Z_CLUSTERED = 6

'Gantt Chart Options           '
Public Const G_SPACED_BARS = 1

'Line/Polar Chart Options      '
Public Const G_SYMBOLS = 1
Public Const G_STICKS = 2
Public Const G_LINES = 4

'Area Chart Options            '
Public Const G_ABSOLUTE = 1
Public Const G_PERCENT = 2

'HLC Chart Options             '
Public Const G_NO_CLOSE = 1
Public Const G_NO_HIGH_LOW = 2


'---------------------------------------
'Key Status Control
'---------------------------------------
'Style
Public Const KEYSTAT_CAPSLOCK = 0
Public Const KEYSTAT_NUMLOCK = 1
Public Const KEYSTAT_INSERT = 2
Public Const KEYSTAT_SCROLLLOCK = 3



'---------------------------------------
'Spin Button
'---------------------------------------
'SpinOrientation
Public Const SPIN_VERTICAL = 0
Public Const SPIN_HORIZONTAL = 1


'---------------------------------------
'Masked Edit Control
'---------------------------------------
'ClipMode
Public Const ME_INCLIT = 0
Public Const ME_EXCLIT = 1


'---------------------------------------
'Comm Control
'---------------------------------------
'Handshaking
Public Const MSCOMM_HANDSHAKE_NONE = 0
Public Const MSCOMM_HANDSHAKE_XONXOFF = 1
Public Const MSCOMM_HANDSHAKE_RTS = 2
Public Const MSCOMM_HANDSHAKE_RTSXONXOFF = 3

'Event constants
Public Const MSCOMM_EV_SEND = 1
Public Const MSCOMM_EV_RECEIVE = 2
Public Const MSCOMM_EV_CTS = 3
Public Const MSCOMM_EV_DSR = 4
Public Const MSCOMM_EV_CD = 5
Public Const MSCOMM_EV_RING = 6
Public Const MSCOMM_EV_EOF = 7

'Error code constants
Public Const MSCOMM_ER_BREAK = 1001
Public Const MSCOMM_ER_CTSTO = 1002
Public Const MSCOMM_ER_DSRTO = 1003
Public Const MSCOMM_ER_FRAME = 1004
Public Const MSCOMM_ER_OVERRUN = 1006
Public Const MSCOMM_ER_CDTO = 1007
Public Const MSCOMM_ER_RXOVER = 1008
Public Const MSCOMM_ER_RXPARITY = 1009
Public Const MSCOMM_ER_TXFULL = 1010


'---------------------------------------
' MAPI SESSION CONTROL CONSTANTS
'---------------------------------------
'Action
Public Const SESSION_SIGNON = 1
Public Const SESSION_SIGNOFF = 2


'---------------------------------------
' MAPI MESSAGE CONTROL CONSTANTS
'---------------------------------------
'Action
Public Const MESSAGE_FETCH = 1             ' Load all messages from message store
Public Const MESSAGE_SENDDLG = 2           ' Send mail bring up default mapi dialog
Public Const MESSAGE_SEND = 3              ' Send mail without default mapi dialog
Public Const MESSAGE_SAVEMSG = 4           ' Save message in the compose buffer
Public Const MESSAGE_COPY = 5              ' Copy current message to compose buffer
Public Const MESSAGE_COMPOSE = 6           ' Initialize compose buffer (previous
                                           ' data is lost
Public Const MESSAGE_REPLY = 7             ' Fill Compose buffer as REPLY
Public Const MESSAGE_REPLYALL = 8          ' Fill Compose buffer as REPLY ALL
Public Const MESSAGE_FORWARD = 9           ' Fill Compose buffer as FORWARD
Public Const MESSAGE_DELETE = 10           ' Delete current message
Public Const MESSAGE_SHOWADBOOK = 11       ' Show Address book
Public Const MESSAGE_SHOWDETAILS = 12      ' Show details of the current recipient
Public Const MESSAGE_RESOLVENAME = 13      ' Resolve the display name of the recipient
Public Const RECIPIENT_DELETE = 14            ' Fill Compose buffer as FORWARD
Public Const ATTACHMENT_DELETE = 15          ' Delete current message


'---------------------------------------
'  ERROR CONSTANT DECLARATIONS (MAPI CONTROLS)
'---------------------------------------
Public Const SUCCESS_SUCCESS = 32000
Public Const MAPI_USER_ABORT = 32001
Public Const MAPI_E_FAILURE = 32002
Public Const MAPI_E_LOGIN_FAILURE = 32003
Public Const MAPI_E_DISK_FULL = 32004
Public Const MAPI_E_INSUFFICIENT_MEMORY = 32005
Public Const MAPI_E_ACCESS_DENIED = 32006
Public Const MAPI_E_TOO_MANY_SESSIONS = 32008
Public Const MAPI_E_TOO_MANY_FILES = 32009
Public Const MAPI_E_TOO_MANY_RECIPIENTS = 32010
Public Const MAPI_E_ATTACHMENT_NOT_FOUND = 32011
Public Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 32012
Public Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 32013
Public Const MAPI_E_UNKNOWN_RECIPIENT = 32014
Public Const MAPI_E_BAD_RECIPTYPE = 32015
Public Const MAPI_E_NO_MESSAGES = 32016
Public Const MAPI_E_INVALID_MESSAGE = 32017
Public Const MAPI_E_TEXT_TOO_LARGE = 32018
Public Const MAPI_E_INVALID_SESSION = 32019
Public Const MAPI_E_TYPE_NOT_SUPPORTED = 32020
Public Const MAPI_E_AMBIGUOUS_RECIPIENT = 32021
Public Const MAPI_E_MESSAGE_IN_USE = 32022
Public Const MAPI_E_NETWORK_FAILURE = 32023
Public Const MAPI_E_INVALID_EDITFIELDS = 32024
Public Const MAPI_E_INVALID_RECIPS = 32025
Public Const MAPI_E_NOT_SUPPORTED = 32026

Public Const CONTROL_E_SESSION_EXISTS = 32050
Public Const CONTROL_E_INVALID_BUFFER = 32051
Public Const CONTROL_E_INVALID_READ_BUFFER_ACTION = 32052
Public Const CONTROL_E_NO_SESSION = 32053
Public Const CONTROL_E_INVALID_RECIPIENT = 32054
Public Const CONTROL_E_INVALID_COMPOSE_BUFFER_ACTION = 32055
Public Const CONTROL_E_FAILURE = 32056
Public Const CONTROL_E_NO_RECIPIENTS = 32057
Public Const CONTROL_E_NO_ATTACHMENTS = 32058


'---------------------------------------
'  MISCELLANEOUS public CONSTANT DECLARATIONS (MAPI CONTROLS)
'---------------------------------------
Public Const RECIPTYPE_ORIG = 0
Public Const RECIPTYPE_TO = 1
Public Const RECIPTYPE_CC = 2
Public Const RECIPTYPE_BCC = 3

Public Const ATTACHTYPE_DATA = 0
Public Const ATTACHTYPE_EOLE = 1
Public Const ATTACHTYPE_SOLE = 2


'-------------------------------------------------
'  Outline
'-------------------------------------------------
' PictureType
Public Const MSOUTLINE_PICTURE_CLOSED = 0
Public Const MSOUTLINE_PICTURE_OPEN = 1
Public Const MSOUTLINE_PICTURE_LEAF = 2

'Outline Control Error Constants
Public Const MSOUTLINE_BADPICFORMAT = 32000
Public Const MSOUTLINE_BADINDENTATION = 32001
Public Const MSOUTLINE_MEM = 32002
Public Const MSOUTLINE_PARENTNOTEXPANDED = 32003

Public Function hMid(strString As String, ByVal nStartPosition As Integer, Optional num As Integer = 300) As String
    Dim i                   As Integer
    Dim nHanLen             As Integer
    Dim strResult           As String
    '/1999.08.04    unicode에서 한글 영문 가져오는 방법 기존midb 를 hmid으로 처리
    
    'Optional ByVal AbsPos As Long = 1
    
    For i = 1 To Len(strString)
        If Asc(Mid(strString, i, 1)) < 0 Then
            nHanLen = nHanLen + 2
        Else
            nHanLen = nHanLen + 1
        End If
        
        If (nHanLen >= nStartPosition) And (nHanLen <= nStartPosition + num - 1) Then
            strResult = strResult + Mid(strString, i, 1)
        End If
    Next i
    
    hMid = strResult
    
End Function

Public Function hLeft(ByVal strCutString As String, nLengthCut As Long) As String
    Dim i                           As Long
    Dim nHangulCount                As Long
    Dim strResultString             As String
    
    nHangulCount = 0                                         '/1999.09.17    unicode에서 한글과 영문이 섞여 있는문자열에서 필요한 길이 만큼 가져 온다.
    For i = 1 To Len(strCutString)
        If Asc(Mid(strCutString, i, 1)) < 0 Then
            nHangulCount = nHangulCount + 2
        Else
            nHangulCount = nHangulCount + 1
        End If
        
        If nHangulCount <= nLengthCut Then     ' - 1
            strResultString = strResultString + Mid(strCutString, i, 1)
            If (nHangulCount - nLengthCut) = 0 Then Exit For
        ElseIf (nHangulCount - nLengthCut) = 1 Then     '마지막에 한글이 한자 걸릴때
            strResultString = strResultString & " "
            Exit For
        Else
            Exit For
        End If
        
    Next i
    hLeft = strResultString
    
End Function

Public Function hCut(ByVal strCutString As String, nLengthCut As Long) As String
    If nLengthCut - Len(strCutString) > 0 Then               '문자열의 길이가 자르려고 하는것 보다 길때
        hCut = hLeft(strCutString & Space(nLengthCut - Len(strCutString)), nLengthCut)
    Else
        hCut = hLeft(strCutString, nLengthCut)
    End If
    
End Function

