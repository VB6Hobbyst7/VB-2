object vCalcMtrl: TvCalcMtrl
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  BorderStyle = bsNone
  Caption = 'vCalcMtrl'
  ClientHeight = 452
  ClientWidth = 675
  Color = clBtnFace
  DoubleBuffered = True
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  OnCreate = FormCreate
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl: TRzPageControl
    AlignWithMargins = True
    Left = 0
    Top = 3
    Width = 672
    Height = 446
    Hint = ''
    Margins.Left = 0
    ActivePage = TabM2
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    Images = PngImageList1
    ParentFont = False
    SortTabMenu = False
    ShowFocusRect = False
    ShowShadow = False
    TabHeight = 25
    TabIndex = 0
    TabOrder = 0
    TabStyle = tsSquareCorners
    UseGradients = False
    FixedDimension = 25
    object TabM2: TRzTabSheet
      ImageIndex = 0
      Caption = 'In-Tubu(Nil, Antigen)'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      object GridM2: TAdvStringGrid
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 664
        Height = 410
        Cursor = crDefault
        Align = alClient
        BorderStyle = bsNone
        ColCount = 7
        Ctl3D = False
        DefaultColWidth = 71
        DrawingStyle = gdsClassic
        RowCount = 44
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        HoverRowCells = [hcNormal, hcSelected]
        OnGetCellColor = GridGetCellColor
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Tahoma'
        ActiveCellFont.Style = [fsBold]
        ColumnHeaders.Strings = (
          'Subject ID'
          'Nil'
          'TB Ag'
          'Mitogen'
          'TB Ag - Nil'
          'Mitogen - Nil'
          'Result')
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
        FixedColWidth = 62
        FixedRowHeight = 22
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = clWindowText
        FixedFont.Height = -11
        FixedFont.Name = 'Tahoma'
        FixedFont.Style = [fsBold]
        Flat = True
        FloatFormat = '%.2f'
        HoverButtons.Buttons = <>
        HoverButtons.Position = hbLeftFromColumnLeft
        HoverFixedCells = hfAll
        HTMLSettings.ImageFolder = 'images'
        HTMLSettings.ImageBaseName = 'img'
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
        ShowSelection = False
        SortSettings.DefaultFormat = ssAutomatic
        Version = '8.2.4.0'
        ColWidths = (
          62
          71
          71
          71
          71
          71
          71)
        RowHeights = (
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
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
    end
    object TabM3: TRzTabSheet
      ImageIndex = 1
      Caption = 'In-Tube(Nil, Antigen, Mitogen)'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object GridM3: TAdvStringGrid
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 664
        Height = 410
        Cursor = crDefault
        Align = alClient
        BorderStyle = bsNone
        ColCount = 7
        Ctl3D = False
        DefaultColWidth = 71
        DrawingStyle = gdsClassic
        RowCount = 44
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        HoverRowCells = [hcNormal, hcSelected]
        OnGetCellColor = GridGetCellColor
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Tahoma'
        ActiveCellFont.Style = [fsBold]
        ColumnHeaders.Strings = (
          'Subject ID'
          'Nil'
          'TB Ag'
          'Mitogen'
          'TB Ag - Nil'
          'Mitogen - Nil'
          'Result')
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
        FixedColWidth = 62
        FixedRowHeight = 22
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = clWindowText
        FixedFont.Height = -11
        FixedFont.Name = 'Tahoma'
        FixedFont.Style = [fsBold]
        Flat = True
        FloatFormat = '%.2f'
        HoverButtons.Buttons = <>
        HoverButtons.Position = hbLeftFromColumnLeft
        HoverFixedCells = hfAll
        HTMLSettings.ImageFolder = 'images'
        HTMLSettings.ImageBaseName = 'img'
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
        ShowSelection = False
        SortSettings.DefaultFormat = ssAutomatic
        Version = '8.2.4.0'
        ColWidths = (
          62
          71
          71
          71
          71
          104
          71)
        RowHeights = (
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
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
    end
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Left = 272
    Top = 152
  end
  object PngImageList1: TPngImageList
    Height = 24
    Width = 24
    PngImages = <
      item
        Background = clWindow
        Name = 'Award -05'
        PngImage.Data = {
          89504E470D0A1A0A0000000D4948445200000018000000180806000000E0773D
          F8000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000001D24944415478DAE5953128846118C79F2B034559D40D37503750
          D42917C35D91147506C560BB2BCA607083C144192483813A8372E5EA06571683
          A228068A2806E536CA0D06E50683E2FFF4FEBFEEED5CDD772775E5A95F7DEFFB
          3DDFFB7FDEE77D9EF7F3C81F9BE75F08D48128E8053D1CDF824B9004F9DF0804
          C00EE804F7E00A7C803ECEE5400C9C5623A08B9C71E118A3B6AD9DE2EA370A0E
          2A11A807378C36C868D78B7CE2143F027ED005DEDC0ACC81352EAE91F7839322
          9F01A6C607EEC03698772BB007BC205CE25D94A97104D47699B2A05B8147E634
          CE71C04A91978BD902BAE315D0043EDD08BC834DB0C0B193A2572994E524B8E0
          F30C488016FA9415D0037E16531DB6805653B284FF0698020D6E53B4CC6DB731
          2247E01C64E9B30A1E40A39843D61E99702BD0CA8F0EF99123609B730609461F
          B6525656402DC44553620EBBB8C6B557B49467C1B49832954A04D4C6415ACC95
          9001D7629A4FEFA588986A4AF26CA41A8166F0C43CDB96B7E68EC150B5024B60
          91CF1929F48556D897E5D72D3FEFAAB2029AE317A6449B4B2BC8E96C1F7796A7
          9F36E558A5025AA683621A2AC29D74F09DFE17F6C1085395669AB26E0534AA28
          D8B2E642DC859A9FD1E738D61D0E4B8926AC895F666D0B7C03531C6319CA3E5F
          9E0000000049454E44AE426082}
      end
      item
        Background = clWindow
        Name = 'Award -04'
        PngImage.Data = {
          89504E470D0A1A0A0000000D4948445200000018000000180806000000E0773D
          F8000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000001C04944415478DAE5953128846118C79F2B03455994E186BBBAE1
          1445B918CE2445312806DB5D51060383C144190C1645312857D40DAE6E31288A
          62A08862A328CA0D862BDF6050FCFFBD8FEEEB73F27EDF6590A77E75DFF7BDDF
          FB7FDEFFF33CDF85E49723F42F04AA400A748076BDBE04A720039C4A045AC106
          6806D7E00CBC824EBD570069701844809B1CE9C669CDDA1D7115E7BA01B0E347
          A01A5C68B609500B263D6B96D49E3D10032DA0682BC0CD167573661E01779E35
          51700FC2E00AAC83695B816DD008BACA3C4BA9359F028C4DB52C612B70A39E4E
          E935C5B2AEDF718F004FBC00EAC09B8DC00B5801339602E360153480671B0116
          F8514C77D858B40C46418DAD45F37AECA8661491EF8BCC0E63913923C3B60211
          7D69575FAA97F26D5A546B983D1BE2C45680910407604B4CB1BD3DCE59612B4F
          8031316D2A7E041843628ACB4F420E9C8B193E7E97FAC5143B2366D2258800AD
          79509FDDE1B8EEED839EA0027360567FE7A43417ECB077D7BA36F9FAADFA5180
          1E3FA925ECFD63294D76584FE6E83A0EE5A05F01764D371851BF7992267DC6FF
          853CE853ABB26AD3ADAD00B34A8135D7BDA49E8211D3EC0B7ACD13F68A29B8AF
          1A541C7F5FE0033DBF5F19AAC49B020000000049454E44AE426082}
      end>
    Left = 336
    Top = 152
    Bitmap = {}
  end
end
