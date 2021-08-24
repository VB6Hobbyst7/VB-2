object vViewNames: TvViewNames
  Left = 0
  Top = 0
  ActiveControl = Grid
  BiDiMode = bdLeftToRight
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Edit view names'
  ClientHeight = 393
  ClientWidth = 855
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  DesignSize = (
    855
    393)
  PixelsPerInch = 96
  TextHeight = 16
  object Shape2: TShape
    Left = 0
    Top = 346
    Width = 855
    Height = 1
    Align = alTop
    Pen.Color = 13421772
    ExplicitTop = 8
  end
  object Shape1: TShape
    Left = 0
    Top = 0
    Width = 855
    Height = 1
    Align = alTop
    Pen.Color = 13421772
  end
  object Grid: TAdvStringGrid
    Left = 0
    Top = 130
    Width = 855
    Height = 216
    Cursor = crDefault
    Align = alTop
    BorderStyle = bsNone
    ColCount = 13
    Ctl3D = True
    DrawingStyle = gdsClassic
    FixedColor = 16316664
    RowCount = 9
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    ParentCtl3D = False
    ScrollBars = ssBoth
    TabOrder = 0
    HoverRowCells = [hcNormal, hcSelected]
    OnCustomCellDraw = GridCustomCellDraw
    OnCanEditCell = GridCanEditCell
    OnEditCellDone = GridEditCellDone
    OnEditChange = GridEditChange
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    ColumnHeaders.Strings = (
      ''
      '1'
      '2'
      '3'
      '4'
      '5'
      '6'
      '7'
      '8'
      '9'
      '10'
      '11'
      '12')
    ControlLook.FixedGradientHoverFrom = clGray
    ControlLook.FixedGradientHoverTo = clWhite
    ControlLook.FixedGradientDownFrom = clGray
    ControlLook.FixedGradientDownTo = clSilver
    ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownHeader.Font.Color = clWindowText
    ControlLook.DropDownHeader.Font.Height = -11
    ControlLook.DropDownHeader.Font.Name = 'Tahoma'
    ControlLook.DropDownHeader.Font.Style = []
    ControlLook.DropDownHeader.Visible = True
    ControlLook.DropDownHeader.Buttons = <>
    ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownFooter.Font.Color = clWindowText
    ControlLook.DropDownFooter.Font.Height = -11
    ControlLook.DropDownFooter.Font.Name = 'Tahoma'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    DefaultAlignment = taCenter
    DefaultEditor = edUpperCase
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'Tahoma'
    FilterDropDown.Font.Style = []
    FilterDropDown.TextChecked = 'Checked'
    FilterDropDown.TextUnChecked = 'Unchecked'
    FilterDropDownClear = '(All)'
    FilterEdit.TypeNames.Strings = (
      'Starts with'
      'Ends with'
      'Contains'
      'Not contains'
      'Equal'
      'Not equal'
      'Larger than'
      'Smaller than'
      'Clear')
    FixedRowHeight = 22
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -11
    FixedFont.Name = 'Tahoma'
    FixedFont.Style = []
    FloatFormat = '%.3f'
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HoverFixedCells = hfAll
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    Look = glStandard
    PrintSettings.DateFormat = 'dd/mm/yyyy'
    PrintSettings.Font.Charset = DEFAULT_CHARSET
    PrintSettings.Font.Color = clWindowText
    PrintSettings.Font.Height = -11
    PrintSettings.Font.Name = 'Tahoma'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'Tahoma'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'Tahoma'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'Tahoma'
    PrintSettings.FooterFont.Style = []
    PrintSettings.PageNumSep = '/'
    RowHeaders.Strings = (
      ''
      'A'
      'B'
      'C'
      'D'
      'E'
      'F'
      'G'
      'H')
    SearchFooter.Color = clBtnFace
    SearchFooter.FindNextCaption = 'Find &next'
    SearchFooter.FindPrevCaption = 'Find &previous'
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
    SearchFooter.Font.Style = []
    SearchFooter.HighLightCaption = 'Highlight'
    SearchFooter.HintClose = 'Close'
    SearchFooter.HintFindNext = 'Find next occurrence'
    SearchFooter.HintFindPrev = 'Find previous occurrence'
    SearchFooter.HintHighlight = 'Highlight occurrences'
    SearchFooter.MatchCaseCaption = 'Match case'
    SearchFooter.ResultFormat = '(%d of %d)'
    SelectionColor = clHighlight
    SelectionRectangle = True
    SelectionTextColor = clHighlightText
    ShowSelection = False
    SortSettings.DefaultFormat = ssAutomatic
    VAlignment = vtaCenter
    Version = '8.2.4.0'
    WordWrap = False
    ColWidths = (
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64)
    RowHeights = (
      22
      22
      22
      22
      22
      22
      22
      22
      22)
  end
  object ButtonCancel: TButton
    Left = 764
    Top = 357
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '&Cancel'
    ModalResult = 2
    TabOrder = 1
  end
  object ButtonOk: TButton
    Left = 683
    Top = 357
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '&OK'
    ModalResult = 1
    TabOrder = 2
    OnClick = ButtonOkClick
  end
  object Panel1: TPanel
    Left = 0
    Top = 1
    Width = 855
    Height = 129
    Align = alTop
    BevelOuter = bvNone
    Color = clWhite
    ParentBackground = False
    TabOrder = 3
    object Image1: TImage
      Left = 42
      Top = 8
      Width = 64
      Height = 64
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000400000
        00400806000000AA6971DE000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000002
        324944415478DAED993F2845511CC78FA22883C16030500603B12862A00C068A
        B21884B22823599445492619D85006C568B0791B83625094C1DD1814C32BAF28
        BEBFEEB9D1EBDDD739BD73EEAF9CDFB7BEBDF36EBFBABFF3B9FF7EBF73AA54E0
        AAE24E805B02803B016E0900EE04B82500B813E09600E04E805BB6003AE04683
        B86B385F74AC0D6ECE604E797D7EE70076E0458B243AE148FF5F82B732987CA2
        3378CC35800FB8D6227E0E3ED0E37BB83D4300C673B301F0AD7F23F83025A60B
        1E2F01E0096E81DFE16D8F939ED1E7F10A20070FA5C4CCC2FB650010BC568F00
        2EE04101200004800010001E01BCC0E7293154ED0D94014005D2A94700237093
        2F006F708345FC347CA4C73770B7C78917EB0BAE710D60155E83AB0D6223B84F
        C5770B6916DE5576956425DA83175C03203528B3BB202A71AC5E993552958A1E
        B357D36069872DE383BE03827F0704FF1508BE0E904A500008000120000440C0
        ABC2C1EF0B04BF33440A7A6FF05F4A0058C4CEC31BCAEC11C8C1132A7EEB9346
        55DC0E67F5086CC2EBAE013CABDF4ECB4453F0B11E5FC2BD194C3E5101AE730D
        20A903E8AADEA6C410A0E47357AA0EA0C4AE3C4E9CD61C92350BA90405800010
        000240007800F0A0E242A394FA555C30A501A00D8B658F0056D4EF67D839804F
        65B62992E82F804715374359CA39801378D2229EDAE13B3DA6567829C3C95337
        D8E31A005D7DDA7830ED05A2A263C32ABB5E80D6030AAE01FC4B0900EE04B825
        00B813E09600E04E805B02803B016E050FE00745E5EE41DABAEADD0000000049
        454E44AE426082}
    end
    object LabelDesc: TLabel
      Left = 112
      Top = 16
      Width = 727
      Height = 56
      AutoSize = False
      Caption = 
        'To edit the view names on the each cells, press enter or double ' +
        'click the cell and then edit it.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object Label1: TLabel
      Left = 51
      Top = 105
      Width = 381
      Height = 16
      Caption = 
        'This will appear on front of each cells. To apply prefix press e' +
        'nter.'
    end
    object Shape3: TShape
      Left = 0
      Top = 128
      Width = 855
      Height = 1
      Align = alBottom
      Pen.Color = 13421772
      ExplicitTop = 8
    end
    object EditPrefix: TEdit
      Left = 51
      Top = 78
      Width = 185
      Height = 24
      TabOrder = 0
      TextHint = 'New prefix'
      OnKeyDown = EditPrefixKeyDown
    end
  end
  object Translator1: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      '.Caption'
      'ButtonCancel.Caption'
      'ButtonOk.Caption'
      'EditPrefix.TextHint'
      'Label1.Caption'
      'LabelDesc.Caption')
    Translatables.Literals = (
      'FF050F61C5854864F9C7B2D465CA7BE8')
    Left = 416
    Top = 200
  end
end
