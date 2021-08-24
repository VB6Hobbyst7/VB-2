Attribute VB_Name = "VbSpread1"
'----------------------------------------------------------
'
' File: SSOCX.BAS
'
' Copyright (C) 1996 FarPoint Technologies.
' All rights reserved.
'
'----------------------------------------------------------
'function prototypes
Declare Function SpreadGetDataFillData Lib "Spread20.VBX" (SS As Control, Var As Variant, ByVal VType As Integer) As Integer
Declare Function SpreadSetDataFillData Lib "Spread20.VBX" (SS As Control, Var As Variant) As Integer
Declare Function SpreadSaveTabFile Lib "Spread20.VBX" (SS As Control, ByVal FileName As String) As Integer
Declare Function SpreadSetCellDirtyFlag Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal Dirty As Integer) As Integer
Declare Function SpreadGetCellDirtyFlag Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
Declare Function SpreadGetMultiSelItem Lib "Spread20.VBX" (SS As Control, ByVal SelPrev As Long) As Long

Declare Function SpreadAddCustomFunction Lib "Spread20.VBX" (hCtl As Control, ByVal lpszFunctionName As String, ByVal nParameterCnt As Integer) As Integer
Declare Function SpreadCFGetDoubleParam Lib "Spread20.VBX" (hCtl As Control, ByVal dParam As Integer) As Double
Declare Function SpreadCFGetLongParam Lib "Spread20.VBX" (hCtl As Control, ByVal dParam As Integer) As Long
Declare Function SpreadCFGetParamInfo Lib "Spread20.VBX" (hCtl As Control, ByVal dParam As Integer, wType As Integer, wStatus As Integer) As Integer
Declare Function SpreadCFGetStringParam Lib "Spread20.VBX" (hCtl As Control, ByVal dParam As Integer) As String
Declare Sub SpreadCFSetResult Lib "Spread20.VBX" (hCtl As Control, Var As Variant)
Declare Function SpreadColNumberToLetter Lib "Spread20.VBX" (ByVal lHeaderNumber As Long) As String
Declare Sub SpreadColWidthToTwips Lib "Spread20.VBX" (SS As Control, ByVal fColWidth As Single, lpTwips As Long)
Declare Sub SpreadGetBottomRightCell Lib "Spread20.VBX" (SS As Control, lpCol As Long, lpRow As Long)
Declare Sub SpreadGetCellFromScreenCoord Lib "Spread20.VBX" (SS As Control, lpCol As Long, lpRow As Long, ByVal X As Long, ByVal Y As Long)
Declare Function SpreadGetCellPos Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lpx As Long, lpy As Long, lpWidth As Long, lpHeight As Long) As Integer
Declare Sub SpreadGetClientArea Lib "Spread20.VBX" (SS As Control, lplWidth As Long, lplHeight As Long)
Declare Function SpreadGetColItemData Lib "Spread20.VBX" (SS As Control, ByVal Col As Long) As Long
Declare Sub SpreadGetFirstValidCell Lib "Spread20.VBX" (SS As Control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetItemData Lib "Spread20.VBX" (SS As Control) As Long
Declare Sub SpreadGetLastValidCell Lib "Spread20.VBX" (SS As Control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetRowItemData Lib "Spread20.VBX" (SS As Control, ByVal Row As Long) As Long
Declare Function SpreadGetText Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, Var As Variant) As Integer
Declare Function SpreadIsFormulaValid Lib "Spread20.VBX" (SS As Control, hszFormula As String) As Integer
Declare Function SpreadIsVisible Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal Partial As Integer) As Integer
Declare Function SpreadIsCellSelected Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
Declare Sub SpreadRowHeightToTwips Lib "Spread20.VBX" (SS As Control, ByVal Row As Long, ByVal fRowHeight As Single, lpTwips As Long)
Declare Sub SpreadSaveDesignInfo Lib "Spread20.VBX" (ByVal Their_hWnd As Integer, ByVal My_hWnd As Integer, ByVal finit As Integer)
Declare Sub SpreadSetColItemData Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal lpVar As Long)
Declare Sub SpreadSetItemData Lib "Spread20.VBX" (SS As Control, ByVal lpVar As Long)
Declare Sub SpreadSetRowItemData Lib "Spread20.VBX" (SS As Control, ByVal Row As Long, ByVal lpVar As Long)
Declare Sub SpreadSetText Lib "Spread20.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lpVar As Variant)
Declare Sub SpreadTwipsToColWidth Lib "Spread20.VBX" (SS As Control, ByVal Twips As Long, fColWidth As Single)
Declare Sub SpreadTwipsToRowHeight Lib "Spread20.VBX" (SS As Control, ByVal Row As Long, ByVal Twips As Long, fRowHeight As Single)

' Action property settings
Public Const SS_ACTION_ACTIVE_CELL = 0
Public Const SS_ACTION_GOTO_CELL = 1
Public Const SS_ACTION_SELECT_BLOCK = 2
Public Const SS_ACTION_CLEAR = 3
Public Const SS_ACTION_DELETE_COL = 4
Public Const SS_ACTION_DELETE_ROW = 5
Public Const SS_ACTION_INSERT_COL = 6
Public Const SS_ACTION_INSERT_ROW = 7
Global Const SS_ACTION_LOAD_SPREAD_SHEET = 8
Global Const SS_ACTION_SAVE_ALL = 9
Global Const SS_ACTION_SAVE_VALUES = 10
Public Const SS_ACTION_RECALC = 11
Public Const SS_ACTION_CLEAR_TEXT = 12
Public Const SS_ACTION_PRINT = 13
Public Const SS_ACTION_DESELECT_BLOCK = 14
Public Const SS_ACTION_DSAVE = 15
Public Const SS_ACTION_SET_CELL_BORDER = 16
Public Const SS_ACTION_ADD_MULTISELBLOCK = 17
Public Const SS_ACTION_GET_MULTI_SELECTION = 18
Public Const SS_ACTION_COPY_RANGE = 19
Public Const SS_ACTION_MOVE_RANGE = 20
Public Const SS_ACTION_SWAP_RANGE = 21
Public Const SS_ACTION_CLIPBOARD_COPY = 22
Public Const SS_ACTION_CLIPBOARD_CUT = 23
Public Const SS_ACTION_CLIPBOARD_PASTE = 24
Public Const SS_ACTION_SORT = 25
Public Const SS_ACTION_COMBO_CLEAR = 26
Public Const SS_ACTION_COMBO_REMOVE = 27
Public Const SS_ACTION_RESET = 28
Public Const SS_ACTION_SEL_MODE_CLEAR = 29
Public Const SS_ACTION_VMODE_REFRESH = 30
Global Const SS_ACTION_REFRESH_BOUND = 31
Public Const SS_ACTION_SMARTPRINT = 32

' SelectBlockOptions property settings
Public Const SS_SELBLOCKOPT_COLS = 1
Public Const SS_SELBLOCKOPT_ROWS = 2
Public Const SS_SELBLOCKOPT_BLOCKS = 4
Public Const SS_SELBLOCKOPT_ALL = 8

' DAutoSize property settings
Public Const SS_AUTOSIZE_NO = 0
Public Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Public Const SS_AUTOSIZE_BEST_GUESS = 2

' BackColorStyle property settings
Public Const SS_BACKCOLORSTYLE_OVERGRID = 0
Public Const SS_BACKCOLORSTYLE_UNDERGRID = 1

' CellType property settings
Public Const SS_CELL_TYPE_DATE = 0
Public Const SS_CELL_TYPE_EDIT = 1
Public Const SS_CELL_TYPE_FLOAT = 2
Public Const SS_CELL_TYPE_INTEGER = 3
Public Const SS_CELL_TYPE_PIC = 4
Public Const SS_CELL_TYPE_STATIC_TEXT = 5
Public Const SS_CELL_TYPE_TIME = 6
Public Const SS_CELL_TYPE_BUTTON = 7
Public Const SS_CELL_TYPE_COMBOBOX = 8
Public Const SS_CELL_TYPE_PICTURE = 9
Public Const SS_CELL_TYPE_CHECKBOX = 10
Public Const SS_CELL_TYPE_OWNER_DRAWN = 11

' CellBorderType property settings
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_OUTLINE = 16
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8

' CellBorderStyle property settings
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_DASH = 2
Public Const SS_BORDER_STYLE_DOT = 3
Public Const SS_BORDER_STYLE_DASH_DOT = 4
Public Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Public Const SS_BORDER_STYLE_BLANK = 6
Public Const SS_BORDER_STYLE_FINE_SOLID = 11
Public Const SS_BORDER_STYLE_FINE_DASH = 12
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

'Bound. auto col sizing
Global Const SS_BOUND_COL_NO_SIZE = 0
Global Const SS_BOUND_COL_MAX_SIZE = 1
Global Const SS_BOUND_COL_SMART_SIZE = 2

' ColHeaderDisplay and RowHeaderDisplay property settings
Public Const SS_HEADER_BLANK = 0
Public Const SS_HEADER_NUMBERS = 1
Public Const SS_HEADER_LETTERS = 2

' TypeCheckTextAlign property settings
Public Const SS_CHECKBOX_TEXT_LEFT = 0
Public Const SS_CHECKBOX_TEXT_RIGHT = 1

' CursorStyle property settings
Public Const SS_CURSOR_STYLE_USER_DEFINED = 0
Public Const SS_CURSOR_STYLE_DEFAULT = 1
Public Const SS_CURSOR_STYLE_ARROW = 2
Public Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Public Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Public Const SS_CURSOR_TYPE_DEFAULT = 0
Public Const SS_CURSOR_TYPE_COLRESIZE = 1
Public Const SS_CURSOR_TYPE_ROWRESIZE = 2
Public Const SS_CURSOR_TYPE_BUTTON = 3
Public Const SS_CURSOR_TYPE_GRAYAREA = 4
Public Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Public Const SS_CURSOR_TYPE_COLHEADER = 6
Public Const SS_CURSOR_TYPE_ROWHEADER = 7

' OperationMode property settings
Public Const SS_OP_MODE_NORMAL = 0
Public Const SS_OP_MODE_READONLY = 1
Public Const SS_OP_MODE_ROWMODE = 2
Public Const SS_OP_MODE_SINGLE_SELECT = 3
Public Const SS_OP_MODE_MULTI_SELECT = 4
Public Const SS_OP_MODE_EXT_SELECT = 5

' SortKeyOrder property settings
Public Const SS_SORT_ORDER_NONE = 0
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_SORT_ORDER_DESCENDING = 2

' SortBy property settings
Public Const SS_SORT_BY_ROW = 0
Public Const SS_SORT_BY_COL = 1

' UserResize property settings
Public Const SS_USER_RESIZE_COL = 1
Public Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Public Const SS_USER_RESIZE_DEFAULT = 0
Public Const SS_USER_RESIZE_ON = 1
Public Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Public Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' Position property settings
Public Const SS_POSITION_UPPER_LEFT = 0
Public Const SS_POSITION_UPPER_CENTER = 1
Public Const SS_POSITION_UPPER_RIGHT = 2
Public Const SS_POSITION_CENTER_LEFT = 3
Public Const SS_POSITION_CENTER_CENTER = 4
Public Const SS_POSITION_CENTER_RIGHT = 5
Public Const SS_POSITION_BOTTOM_LEFT = 6
Public Const SS_POSITION_BOTTOM_CENTER = 7
Public Const SS_POSITION_BOTTOM_RIGHT = 8

' ScrollBars property settings
Public Const SS_SCROLLBAR_NONE = 0
Public Const SS_SCROLLBAR_H_ONLY = 1
Public Const SS_SCROLLBAR_V_ONLY = 2
Public Const SS_SCROLLBAR_BOTH = 3

' PrintOrientation property settings
Public Const SS_PRINTORIENT_DEFAULT = 0
Public Const SS_PRINTORIENT_PORTRAIT = 1
Public Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Public Const SS_PRINT_ALL = 0
Public Const SS_PRINT_CELL_RANGE = 1
Public Const SS_PRINT_CURRENT_PAGE = 2
Public Const SS_PRINT_PAGE_RANGE = 3

' TypeButtonType property settings
Public Const SS_CELL_BUTTON_NORMAL = 0
Public Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeButtonAlign property settings
Public Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Public Const SS_CELL_BUTTON_ALIGN_TOP = 1
Public Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Public Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' ButtonDrawMode property settings
Public Const SS_BDM_ALWAYS = 0
Public Const SS_BDM_CURRENT_CELL = 1
Public Const SS_BDM_CURRENT_COLUMN = 2
Public Const SS_BDM_CURRENT_ROW = 4

' TypeDateFormat property settings
Public Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Public Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Public Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Public Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Public Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Public Const SS_CELL_EDIT_CASE_NO_CASE = 1
Public Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Public Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Public Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeTextAlignVert property settings
Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Public Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Public Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Public Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Public Const SS_CELL_TIME_24_HOUR_CLOCK = 1

'Unit type
Public Const SS_CELL_UNIT_NORMAL = 0
Public Const SS_CELL_UNIT_VGA = 1
Public Const SS_CELL_UNIT_TWIPS = 2

' TypeHAlign property settings
Public Const SS_CELL_H_ALIGN_LEFT = 0
Public Const SS_CELL_H_ALIGN_RIGHT = 1
Public Const SS_CELL_H_ALIGN_CENTER = 2

' EditEnterAction property settings
Public Const SS_CELL_EDITMODE_EXIT_NONE = 0
Public Const SS_CELL_EDITMODE_EXIT_UP = 1
Public Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Public Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Public Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Public Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Public Const SS_CELL_EDITMODE_EXIT_SAME = 7
Public Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' Custom function parameter type used with CFGetParamInfo method
Public Const SS_VALUE_TYPE_LONG = 0
Public Const SS_VALUE_TYPE_DOUBLE = 1
Public Const SS_VALUE_TYPE_STR = 2
Public Const SS_VALUE_TYPE_CELL = 3
Public Const SS_VALUE_TYPE_RANGE = 4

' Custom function parameter status used with CFGetParamInfo method
Public Const SS_VALUE_STATUS_OK = 0
Public Const SS_VALUE_STATUS_ERROR = 1
Public Const SS_VALUE_STATUS_EMPTY = 2
Global Const SS_VALUE_STATUS_CLEAR = 3
Global Const SS_VALUE_STATUS_NONE = 4

' Reference style settings used with GetRefStyle/SetRefStyle methods
Public Const SS_REFSTYLE_DEFAULT = 0
Public Const SS_REFSTYLE_A1 = 1
Public Const SS_REFSTYLE_R1C1 = 2

' Options used with Flags parameter of AddCustomFunctionExt method
Public Const SS_CUSTFUNC_WANTCELLREF = 1
Public Const SS_CUSTFUNC_WANTRANGEREF = 2

