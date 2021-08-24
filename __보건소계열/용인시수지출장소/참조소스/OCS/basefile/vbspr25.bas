Attribute VB_Name = "modSpr25"
'----------------------------------------------------------
'
' File: SSVBX.BAS
'
' Copyright (C) 1996 FarPoint Technologies.
' All rights reserved.
'
'----------------------------------------------------------

' Action property settings
Global Const SS_ACTION_ACTIVE_CELL = 0
Global Const SS_ACTION_GOTO_CELL = 1
Global Const SS_ACTION_SELECT_BLOCK = 2
Global Const SS_ACTION_CLEAR = 3
Global Const SS_ACTION_DELETE_COL = 4
Global Const SS_ACTION_DELETE_ROW = 5
Global Const SS_ACTION_INSERT_COL = 6
Global Const SS_ACTION_INSERT_ROW = 7
Global Const SS_ACTION_LOAD_SPREAD_SHEET = 8
Global Const SS_ACTION_SAVE_ALL = 9
Global Const SS_ACTION_SAVE_VALUES = 10
Global Const SS_ACTION_RECALC = 11
Global Const SS_ACTION_CLEAR_TEXT = 12
Global Const SS_ACTION_PRINT = 13
Global Const SS_ACTION_DESELECT_BLOCK = 14
Global Const SS_ACTION_DSAVE = 15
Global Const SS_ACTION_SET_CELL_BORDER = 16
Global Const SS_ACTION_ADD_MULTISELBLOCK = 17
Global Const SS_ACTION_GET_MULTI_SELECTION = 18
Global Const SS_ACTION_COPY_RANGE = 19
Global Const SS_ACTION_MOVE_RANGE = 20
Global Const SS_ACTION_SWAP_RANGE = 21
Global Const SS_ACTION_CLIPBOARD_COPY = 22
Global Const SS_ACTION_CLIPBOARD_CUT = 23
Global Const SS_ACTION_CLIPBOARD_PASTE = 24
Global Const SS_ACTION_SORT = 25
Global Const SS_ACTION_COMBO_CLEAR = 26
Global Const SS_ACTION_COMBO_REMOVE = 27
Global Const SS_ACTION_RESET = 28
Global Const SS_ACTION_SEL_MODE_CLEAR = 29
Global Const SS_ACTION_VMODE_REFRESH = 30
Global Const SS_ACTION_SMARTPRINT = 32

' SelectBlockOptions property settings
Global Const SS_SELBLOCKOPT_COLS = 1
Global Const SS_SELBLOCKOPT_ROWS = 2
Global Const SS_SELBLOCKOPT_BLOCKS = 4
Global Const SS_SELBLOCKOPT_ALL = 8

' BackColorStyle property settings
Global Const SS_BACKCOLORSTYLE_OVERGRID = 0
Global Const SS_BACKCOLORSTYLE_UNDERGRID = 1

' DAutoSize property settings
Global Const SS_AUTOSIZE_NO = 0
Global Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Global Const SS_AUTOSIZE_BEST_GUESS = 2

' CellType property settings
Global Const SS_CELL_TYPE_DATE = 0
Global Const SS_CELL_TYPE_EDIT = 1
Global Const SS_CELL_TYPE_FLOAT = 2
Global Const SS_CELL_TYPE_INTEGER = 3
Global Const SS_CELL_TYPE_PIC = 4
Global Const SS_CELL_TYPE_STATIC_TEXT = 5
Global Const SS_CELL_TYPE_TIME = 6
Global Const SS_CELL_TYPE_BUTTON = 7
Global Const SS_CELL_TYPE_COMBOBOX = 8
Global Const SS_CELL_TYPE_PICTURE = 9
Global Const SS_CELL_TYPE_CHECKBOX = 10
Global Const SS_CELL_TYPE_OWNER_DRAWN = 11

' CellBorderType property settings
Global Const SS_BORDER_TYPE_NONE = 0
Global Const SS_BORDER_TYPE_OUTLINE = 16
Global Const SS_BORDER_TYPE_LEFT = 1
Global Const SS_BORDER_TYPE_RIGHT = 2
Global Const SS_BORDER_TYPE_TOP = 4
Global Const SS_BORDER_TYPE_BOTTOM = 8

' CellBorderStyle property settings
Global Const SS_BORDER_STYLE_DEFAULT = 0
Global Const SS_BORDER_STYLE_SOLID = 1
Global Const SS_BORDER_STYLE_DASH = 2
Global Const SS_BORDER_STYLE_DOT = 3
Global Const SS_BORDER_STYLE_DASH_DOT = 4
Global Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Global Const SS_BORDER_STYLE_BLANK = 6
Global Const SS_BORDER_STYLE_FINE_SOLID = 11
Global Const SS_BORDER_STYLE_FINE_DASH = 12
Global Const SS_BORDER_STYLE_FINE_DOT = 13
Global Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Global Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Global Const SS_HEADER_BLANK = 0
Global Const SS_HEADER_NUMBERS = 1
Global Const SS_HEADER_LETTERS = 2

' TypeCheckTextAlign property settings
Global Const SS_CHECKBOX_TEXT_LEFT = 0
Global Const SS_CHECKBOX_TEXT_RIGHT = 1

' CursorStyle property settings
Global Const SS_CURSOR_STYLE_USER_DEFINED = 0
Global Const SS_CURSOR_STYLE_DEFAULT = 1
Global Const SS_CURSOR_STYLE_ARROW = 2
Global Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Global Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Global Const SS_CURSOR_TYPE_DEFAULT = 0
Global Const SS_CURSOR_TYPE_COLRESIZE = 1
Global Const SS_CURSOR_TYPE_ROWRESIZE = 2
Global Const SS_CURSOR_TYPE_BUTTON = 3
Global Const SS_CURSOR_TYPE_GRAYAREA = 4
Global Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Global Const SS_CURSOR_TYPE_COLHEADER = 6
Global Const SS_CURSOR_TYPE_ROWHEADER = 7

' OperationMode property settings
Global Const SS_OP_MODE_NORMAL = 0
Global Const SS_OP_MODE_READONLY = 1
Global Const SS_OP_MODE_ROWMODE = 2
Global Const SS_OP_MODE_SINGLE_SELECT = 3
Global Const SS_OP_MODE_MULTI_SELECT = 4
Global Const SS_OP_MODE_EXT_SELECT = 5

' SortKeyOrder property settings
Global Const SS_SORT_ORDER_NONE = 0
Global Const SS_SORT_ORDER_ASCENDING = 1
Global Const SS_SORT_ORDER_DESCENDING = 2

' SortBy property settings
Global Const SS_SORT_BY_ROW = 0
Global Const SS_SORT_BY_COL = 1

' UserResize property settings
Global Const SS_USER_RESIZE_COL = 1
Global Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Global Const SS_USER_RESIZE_DEFAULT = 0
Global Const SS_USER_RESIZE_ON = 1
Global Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Global Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Global Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Global Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' Position property settings
Global Const SS_POSITION_UPPER_LEFT = 0
Global Const SS_POSITION_UPPER_CENTER = 1
Global Const SS_POSITION_UPPER_RIGHT = 2
Global Const SS_POSITION_CENTER_LEFT = 3
Global Const SS_POSITION_CENTER_CENTER = 4
Global Const SS_POSITION_CENTER_RIGHT = 5
Global Const SS_POSITION_BOTTOM_LEFT = 6
Global Const SS_POSITION_BOTTOM_CENTER = 7
Global Const SS_POSITION_BOTTOM_RIGHT = 8

' ScrollBars property settings
Global Const SS_SCROLLBAR_NONE = 0
Global Const SS_SCROLLBAR_H_ONLY = 1
Global Const SS_SCROLLBAR_V_ONLY = 2
Global Const SS_SCROLLBAR_BOTH = 3

' PrintOrientation property settings
Global Const SS_PRINTORIENT_DEFAULT = 0
Global Const SS_PRINTORIENT_PORTRAIT = 1
Global Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Global Const SS_PRINT_ALL = 0
Global Const SS_PRINT_CELL_RANGE = 1
Global Const SS_PRINT_CURRENT_PAGE = 2
Global Const SS_PRINT_PAGE_RANGE = 3

' TypeButtonType property settings
Global Const SS_CELL_BUTTON_NORMAL = 0
Global Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeButtonAlign property settings
Global Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Global Const SS_CELL_BUTTON_ALIGN_TOP = 1
Global Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Global Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' ButtonDrawMode property settings
Global Const SS_BDM_ALWAYS = 0
Global Const SS_BDM_CURRENT_CELL = 1
Global Const SS_BDM_CURRENT_COLUMN = 2
Global Const SS_BDM_CURRENT_ROW = 4

' TypeDateFormat property settings
Global Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Global Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Global Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Global Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Global Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Global Const SS_CELL_EDIT_CASE_NO_CASE = 1
Global Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Global Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Global Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Global Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Global Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeTextAlignVert property settings
Global Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Global Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Global Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Global Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Global Const SS_CELL_TIME_24_HOUR_CLOCK = 1

' UnitType property settings
Global Const SS_CELL_UNIT_NORMAL = 0
Global Const SS_CELL_UNIT_VGA = 1
Global Const SS_CELL_UNIT_TWIPS = 2

' TypeHAlign property settings
Global Const SS_CELL_H_ALIGN_LEFT = 0
Global Const SS_CELL_H_ALIGN_RIGHT = 1
Global Const SS_CELL_H_ALIGN_CENTER = 2

' EditEnterAction property settings
Global Const SS_CELL_EDITMODE_EXIT_NONE = 0
Global Const SS_CELL_EDITMODE_EXIT_UP = 1
Global Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Global Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Global Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Global Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Global Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Global Const SS_CELL_EDITMODE_EXIT_SAME = 7
Global Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' Custom function parameter type used with CFGetParamInfo method
Global Const SS_VALUE_TYPE_LONG = 0
Global Const SS_VALUE_TYPE_DOUBLE = 1
Global Const SS_VALUE_TYPE_STR = 2
Global Const SS_VALUE_TYPE_CELL = 3
Global Const SS_VALUE_TYPE_RANGE = 4

' Custom function parameter status used with CFGetParamInfo method
Global Const SS_VALUE_STATUS_OK = 0
Global Const SS_VALUE_STATUS_ERROR = 1
Global Const SS_VALUE_STATUS_EMPTY = 2

' Reference style settings used with GetRefStyle/SetRefStyle methods
Global Const SS_REFSTYLE_DEFAULT = 0
Global Const SS_REFSTYLE_A1 = 1
Global Const SS_REFSTYLE_R1C1 = 2

' Options used with Flags parameter of AddCustomFunctionExt method
Global Const SS_CUSTFUNC_WANTCELLREF = 1
Global Const SS_CUSTFUNC_WANTRANGEREF = 2

' Function prototypes
Declare Function SpreadAddCustomFunction Lib "SSVBX25.VBX" (hCtl As Control, ByVal lpszFunctionName As String, ByVal nParameterCnt As Integer) As Integer
Declare Function SpreadAddCustomFunctionExt Lib "SSVBX25.VBX" (hCtl As Control, ByVal lpszFunctionName As String, ByVal nMinParamCnt As Integer, ByVal nMaxParamCnt As Integer, ByVal Flags As Long) As Integer
Declare Sub SpreadCFGetCellParam Lib "SSVBX25.VBX" (hCtl As Control, ByVal dParam As Integer, Col As Long, Row As Long)
Declare Function SpreadCFGetDoubleParam Lib "SSVBX25.VBX" (hCtl As Control, ByVal dParam As Integer) As Double
Declare Function SpreadCFGetLongParam Lib "SSVBX25.VBX" (hCtl As Control, ByVal dParam As Integer) As Long
Declare Function SpreadCFGetParamInfo Lib "SSVBX25.VBX" (hCtl As Control, ByVal dParam As Integer, wType As Integer, wStatus As Integer) As Integer
Declare Sub SpreadCFGetRangeParam Lib "SSVBX25.VBX" (hCtl As Control, ByVal dParam As Integer, Col As Long, Row As Long, Col2 As Long, Row2 As Long)
Declare Function SpreadCFGetStringParam Lib "SSVBX25.VBX" (hCtl As Control, ByVal dParam As Integer) As String
Declare Sub SpreadCFSetResult Lib "SSVBX25.VBX" (hCtl As Control, Var As Variant)
Declare Function SpreadColNumberToLetter Lib "SSVBX25.VBX" (ByVal lHeaderNumber As Long) As String
Declare Sub SpreadColWidthToTwips Lib "SSVBX25.VBX" (SS As Control, ByVal fColWidth As Single, lpTwips As Long)
Declare Sub SpreadGetBottomRightCell Lib "SSVBX25.VBX" (SS As Control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetCellDirtyFlag Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
Declare Sub SpreadGetCellFromScreenCoord Lib "SSVBX25.VBX" (SS As Control, lpCol As Long, lpRow As Long, ByVal x As Long, ByVal y As Long)
Declare Function SpreadGetCellPos Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lpx As Long, lpy As Long, lpWidth As Long, lpHeight As Long) As Integer
Declare Sub SpreadGetClientArea Lib "SSVBX25.VBX" (SS As Control, lplWidth As Long, lplHeight As Long)
Declare Function SpreadGetColItemData Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long) As Long
Declare Function SpreadGetCustomName Lib "SSVBX25.VBX" (SS As Control, ByVal sName As String) As String
'Declare Function SpreadGetDataConnectHandle Lib "SSVBX25.VBX" (SS As Control) As Integer
Declare Function SpreadGetDataFillData Lib "SSVBX25.VBX" (SS As Control, Var As Variant, ByVal VType As Integer) As Integer
'Declare Function SpreadGetDataSelectHandle Lib "SSVBX25.VBX" (SS As Control) As Integer
Declare Sub SpreadGetFirstValidCell Lib "SSVBX25.VBX" (SS As Control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetFormulaSync Lib "SSVBX25.VBX" (SS As Control) As Integer
Declare Function SpreadGetItemData Lib "SSVBX25.VBX" (SS As Control) As Long
Declare Function SpreadGetIteration Lib "SSVBX25.VBX" (SS As Control, MaxIterations As Integer, MaxChange As Double) As Integer
Declare Sub SpreadGetLastValidCell Lib "SSVBX25.VBX" (SS As Control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetMultiSelItem Lib "SSVBX25.VBX" (SS As Control, ByVal SelPrev As Long) As Long
Declare Function SpreadGetRefStyle Lib "SSVBX25.VBX" (SS As Control) As Integer
Declare Function SpreadGetRowItemData Lib "SSVBX25.VBX" (SS As Control, ByVal Row As Long) As Long
Declare Function SpreadGetText Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, Var As Variant) As Integer
Declare Function SpreadIsCellSelected Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
Declare Function SpreadIsFormulaValid Lib "SSVBX25.VBX" (SS As Control, hszFormula As String) As Integer
Declare Function SpreadIsVisible Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal Partial As Integer) As Integer
Declare Function SpreadLoadFromFile Lib "SSVBX25.VBX" (SS As Control, ByVal FileName As String) As Integer
Declare Function SpreadLoadTabFile Lib "SSVBX25.VBX" (SS As Control, ByVal FileName As String) As Integer
Declare Function SpreadQueryCustomName Lib "SSVBX25.VBX" (SS As Control, ByVal sPrevName As String) As String
Declare Sub SpreadRowHeightToTwips Lib "SSVBX25.VBX" (SS As Control, ByVal Row As Long, ByVal fRowHeight As Single, lpTwips As Long)
'Declare Sub SpreadSaveDesignInfo Lib "SSVBX25.VBX" (ByVal Their_hWnd As Integer, ByVal My_hWnd As Integer, ByVal finit As Integer)
Declare Function SpreadSaveTabFile Lib "SSVBX25.VBX" (SS As Control, ByVal FileName As String) As Integer
Declare Function SpreadSaveToFile Lib "SSVBX25.VBX" (SS As Control, ByVal FileName As String, ByVal DataOnly As Integer) As Integer
Declare Function SpreadSetCellDirtyFlag Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal Dirty As Integer) As Integer
Declare Sub SpreadSetColItemData Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal lpVar As Long)
Declare Function SpreadSetCustomName Lib "SSVBX25.VBX" (SS As Control, ByVal sName As String, ByVal sValue As String) As Integer
Declare Function SpreadSetDataFillData Lib "SSVBX25.VBX" (SS As Control, Var As Variant) As Integer
Declare Sub SpreadSetFormulaSync Lib "SSVBX25.VBX" (SS As Control, ByVal Sync As Integer)
Declare Sub SpreadSetItemData Lib "SSVBX25.VBX" (SS As Control, ByVal lpVar As Long)
Declare Sub SpreadSetIteration Lib "SSVBX25.VBX" (SS As Control, ByVal Iteration As Integer, ByVal MaxIterations As Integer, ByVal MaxChange As Double)
Declare Sub SpreadSetRefStyle Lib "SSVBX25.VBX" (SS As Control, ByVal RefStyle As Integer)
Declare Sub SpreadSetRowItemData Lib "SSVBX25.VBX" (SS As Control, ByVal Row As Long, ByVal lpVar As Long)
Declare Sub SpreadSetText Lib "SSVBX25.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lpVar As Variant)
Declare Sub SpreadTwipsToColWidth Lib "SSVBX25.VBX" (SS As Control, ByVal Twips As Long, fColWidth As Single)
Declare Sub SpreadTwipsToRowHeight Lib "SSVBX25.VBX" (SS As Control, ByVal Row As Long, ByVal Twips As Long, fRowHeight As Single)

