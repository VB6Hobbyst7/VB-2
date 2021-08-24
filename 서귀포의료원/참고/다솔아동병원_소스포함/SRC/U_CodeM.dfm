object F_CodeM: TF_CodeM
  Left = 148
  Top = 217
  Width = 633
  Height = 610
  Caption = 'Code Setting'
  Color = clBtnFace
  Font.Charset = HANGEUL_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #44404#47548#52404
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 617
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    BorderStyle = bsSingle
    TabOrder = 0
    object btnClose: TSpeedButton
      Left = 412
      Top = 3
      Width = 81
      Height = 30
      Caption = #45803#44592
      OnClick = btnCloseClick
    end
    object btnView: TSpeedButton
      Left = 326
      Top = 3
      Width = 81
      Height = 30
      Caption = #51312#54924
      OnClick = btnViewClick
    end
    object Panel17: TPanel
      Left = 0
      Top = 0
      Width = 254
      Height = 37
      Align = alLeft
      BevelOuter = bvNone
      Caption = #44160#49324#53076#46300' '#49444#51221
      Color = 15658734
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #44404#47548#52404
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 553
    Width = 617
    Height = 19
    Panels = <>
  end
  object Panel2: TPanel
    Left = 0
    Top = 41
    Width = 617
    Height = 512
    Align = alClient
    BevelOuter = bvNone
    BorderStyle = bsSingle
    Caption = 'Panel2'
    TabOrder = 2
    object Panel3: TPanel
      Left = 0
      Top = 0
      Width = 360
      Height = 508
      Align = alClient
      BevelOuter = bvNone
      BorderStyle = bsSingle
      Caption = 'Panel3'
      TabOrder = 0
      object gdCodeM: TAdvStringGrid
        Left = 0
        Top = 0
        Width = 356
        Height = 504
        Cursor = crDefault
        Align = alClient
        BorderStyle = bsNone
        ColCount = 10
        Ctl3D = False
        DefaultRowHeight = 21
        FixedCols = 0
        RowCount = 2
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected]
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssBoth
        TabOrder = 0
        OnClick = gdCodeMClick
        OnClickCell = gdCodeMClickCell
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Tahoma'
        ActiveCellFont.Style = [fsBold]
        Bands.Active = True
        Bands.PrimaryColor = 16641503
        CellNode.TreeColor = clSilver
        ColumnHeaders.Strings = (
          'NO'
          #52376#48169#53076#46300
          #44160#49324#54032#45356
          #44160#49324#53076#46300
          #44160#49324#47749
          #50557#50612
          #51204#49569#53076#46300
          #49688#49888#53076#46300
          #49345#54620#52824
          #54616#54620#52824)
        ControlLook.ControlStyle = csWinXP
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
        ControlLook.DropDownFooter.Font.Name = 'MS Sans Serif'
        ControlLook.DropDownFooter.Font.Style = []
        ControlLook.DropDownFooter.Visible = True
        ControlLook.DropDownFooter.Buttons = <>
        Filter = <>
        FilterDropDown.Font.Charset = DEFAULT_CHARSET
        FilterDropDown.Font.Color = clWindowText
        FilterDropDown.Font.Height = -11
        FilterDropDown.Font.Name = 'MS Sans Serif'
        FilterDropDown.Font.Style = []
        FilterDropDownClear = '(All)'
        FixedColWidth = 41
        FixedFont.Charset = HANGEUL_CHARSET
        FixedFont.Color = clWindowText
        FixedFont.Height = -13
        FixedFont.Name = #44404#47548#52404
        FixedFont.Style = []
        Flat = True
        FloatFormat = '%.2f'
        PrintSettings.DateFormat = 'dd/mm/yyyy'
        PrintSettings.Font.Charset = DEFAULT_CHARSET
        PrintSettings.Font.Color = clWindowText
        PrintSettings.Font.Height = -11
        PrintSettings.Font.Name = 'MS Sans Serif'
        PrintSettings.Font.Style = []
        PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
        PrintSettings.FixedFont.Color = clWindowText
        PrintSettings.FixedFont.Height = -11
        PrintSettings.FixedFont.Name = 'MS Sans Serif'
        PrintSettings.FixedFont.Style = []
        PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
        PrintSettings.HeaderFont.Color = clWindowText
        PrintSettings.HeaderFont.Height = -11
        PrintSettings.HeaderFont.Name = 'MS Sans Serif'
        PrintSettings.HeaderFont.Style = []
        PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
        PrintSettings.FooterFont.Color = clWindowText
        PrintSettings.FooterFont.Height = -11
        PrintSettings.FooterFont.Name = 'MS Sans Serif'
        PrintSettings.FooterFont.Style = []
        PrintSettings.PageNumSep = '/'
        ScrollWidth = 16
        SearchFooter.FindNextCaption = 'Find &next'
        SearchFooter.FindPrevCaption = 'Find &previous'
        SearchFooter.Font.Charset = DEFAULT_CHARSET
        SearchFooter.Font.Color = clWindowText
        SearchFooter.Font.Height = -11
        SearchFooter.Font.Name = 'MS Sans Serif'
        SearchFooter.Font.Style = []
        SearchFooter.HighLightCaption = 'Highlight'
        SearchFooter.HintClose = 'Close'
        SearchFooter.HintFindNext = 'Find next occurence'
        SearchFooter.HintFindPrev = 'Find previous occurence'
        SearchFooter.HintHighlight = 'Highlight occurences'
        SearchFooter.MatchCaseCaption = 'Match case'
        SortSettings.Column = 0
        Version = '5.0.9.0'
        ColWidths = (
          41
          74
          100
          80
          183
          70
          64
          54
          56
          64)
      end
    end
    object Panel4: TPanel
      Left = 360
      Top = 0
      Width = 253
      Height = 508
      Align = alRight
      BorderStyle = bsSingle
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object btnAdd: TSpeedButton
        Left = 17
        Top = 12
        Width = 71
        Height = 29
        Caption = #49352#53076#46300
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        OnClick = btnAddClick
      end
      object btnDel: TSpeedButton
        Left = 163
        Top = 12
        Width = 71
        Height = 29
        Caption = #49325#51228
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        OnClick = btnDelClick
      end
      object Panel5: TPanel
        Left = 8
        Top = 107
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #44160#49324#53076#46300
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object Panel6: TPanel
        Left = 8
        Top = 134
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #44160#49324#47749
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 7
      end
      object Panel7: TPanel
        Left = 8
        Top = 214
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #50557#50612
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 9
      end
      object Panel8: TPanel
        Left = 8
        Top = 239
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = 'REF(Low)'
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 11
      end
      object Panel9: TPanel
        Left = 8
        Top = 264
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = 'REF(High)'
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 6
      end
      object edCode: TEdit
        Left = 91
        Top = 107
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 1
      end
      object edName: TEdit
        Left = 91
        Top = 134
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 2
      end
      object edAbbr: TEdit
        Left = 91
        Top = 214
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 4
      end
      object Panel10: TPanel
        Left = 8
        Top = 289
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #49692#49436
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 8
      end
      object Panel11: TPanel
        Left = 8
        Top = 187
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #51109#48708#49688#49888#53076#46300
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 10
      end
      object edUpCode: TEdit
        Left = 91
        Top = 187
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 3
      end
      object btnSave: TBitBtn
        Left = 90
        Top = 12
        Width = 71
        Height = 29
        Caption = #51200#51109
        TabOrder = 5
        OnClick = btnSaveClick
      end
      object edLow: TEdit
        Left = 91
        Top = 239
        Width = 48
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 12
        OnKeyPress = edAbbrKeyPress
      end
      object edHigh: TEdit
        Left = 91
        Top = 264
        Width = 48
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 13
        OnKeyPress = edAbbrKeyPress
      end
      object seSeq: TSpinEdit
        Left = 92
        Top = 289
        Width = 49
        Height = 21
        MaxValue = 0
        MinValue = 0
        TabOrder = 14
        Value = 0
      end
      object Panel12: TPanel
        Left = 8
        Top = 53
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #52376#48169#53076#46300
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 15
      end
      object edOrdcd: TEdit
        Left = 91
        Top = 53
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 16
      end
      object Panel13: TPanel
        Left = 8
        Top = 80
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #44160#49324#54032#45356
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 17
      end
      object edPanel: TEdit
        Left = 91
        Top = 80
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 18
      end
      object Panel14: TPanel
        Left = 8
        Top = 160
        Width = 84
        Height = 20
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = #51109#48708#49569#49888#53076#46300
        Color = 15588058
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        TabOrder = 19
      end
      object edIfcd: TEdit
        Left = 91
        Top = 160
        Width = 89
        Height = 20
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 20
      end
    end
  end
end
