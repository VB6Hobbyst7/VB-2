object vMain: TvMain
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  Caption = 'ELISA Report'
  ClientHeight = 680
  ClientWidth = 993
  Color = 14671839
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  GlassFrame.Enabled = True
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poScreenCenter
  StyleElements = [seFont]
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object SplitterProperty: TSplitter
    Left = 217
    Top = 145
    Width = 1
    Height = 516
    Color = 12566463
    ParentColor = False
    Visible = False
    StyleElements = [seFont]
  end
  object Ribbon: TUIRibbon
    Left = 0
    Top = 0
    Width = 993
    Height = 117
    ResourceName = 'APPLICATION'
    ActionManager = ActionManager
  end
  object RzStatusBar1: TRzStatusBar
    Left = 0
    Top = 661
    Width = 993
    Height = 19
    BorderInner = fsNone
    BorderOuter = fsNone
    BorderSides = [sdLeft, sdTop, sdRight, sdBottom]
    BorderWidth = 0
    TabOrder = 1
    object StatusTestDate: TRzFieldStatus
      Left = 0
      Top = 0
      Width = 91
      Height = 19
      Align = alLeft
      FieldLabel = 'Test Date'
      FieldLabelColor = clBtnShadow
      AutoSize = True
      Caption = ''
    end
    object StatusTestNum: TRzFieldStatus
      Left = 91
      Top = 0
      Width = 104
      Height = 19
      Align = alLeft
      FieldLabel = 'Test number'
      FieldLabelColor = clBtnShadow
      AutoSize = True
      Caption = ''
      ExplicitLeft = 85
    end
    object StatusKitBatch: TRzFieldStatus
      Left = 195
      Top = 0
      Width = 126
      Height = 19
      Align = alLeft
      FieldLabel = 'Kit-Batch number'
      FieldLabelColor = clBtnShadow
      AutoSize = True
      Caption = ''
    end
    object StatusOperator: TRzFieldStatus
      Left = 321
      Top = 0
      Width = 88
      Height = 19
      Align = alLeft
      FieldLabel = 'Operator'
      FieldLabelColor = clBtnShadow
      AutoSize = True
      Caption = ''
      ExplicitLeft = 299
    end
    object StatusVersion: TRzFieldStatus
      Left = 893
      Top = 0
      Height = 19
      Align = alRight
      FieldLabel = 'ver'
      Caption = ''
      ExplicitLeft = 993
      ExplicitHeight = 20
    end
  end
  object PageControl: TRzPageControl
    Left = 218
    Top = 145
    Width = 775
    Height = 516
    Hint = ''
    ActivePage = TabSheetRawData
    Align = alClient
    BoldCurrentTab = True
    ButtonColor = 14671839
    ButtonColorDisabled = 14671839
    ButtonSymbolColor = 14671839
    ButtonSymbolColorDisabled = 14671839
    UseColoredTabs = True
    FlatColor = 10263441
    HotTrackStyle = htsTabBar
    Images = PngImageList1
    Margin = 5
    ShowCardFrame = False
    SortTabMenu = False
    ShowShadow = False
    TabColors.HighlightBar = 7434609
    TabColors.Shadow = 14671839
    TabColors.Unselected = 14671839
    TabIndex = 0
    TabOrder = 2
    TabOrientation = toBottom
    TabStyle = tsSquareCorners
    UseGradients = False
    FixedDimension = 22
    object TabSheetRawData: TRzTabSheet
      Color = clWhite
      ImageIndex = 0
      Caption = 'Data'
      object SplitterFmt: TSplitter
        Left = 546
        Top = 0
        Width = 5
        Height = 491
        Align = alRight
        Color = 14671839
        ParentColor = False
        Visible = False
        ExplicitLeft = 532
        ExplicitHeight = 519
      end
      object ShapeFmt: TShape
        AlignWithMargins = True
        Left = 774
        Top = 3
        Width = 1
        Height = 486
        Margins.Left = 0
        Margins.Right = 0
        Margins.Bottom = 2
        Align = alRight
        Pen.Color = 12566463
        Visible = False
        ExplicitLeft = 241
        ExplicitTop = 128
        ExplicitHeight = 444
      end
      object PanelFmt: TPanel
        Left = 551
        Top = 0
        Width = 223
        Height = 491
        Align = alRight
        BevelOuter = bvNone
        TabOrder = 0
        Visible = False
      end
    end
    object TabSheetStdResult: TRzTabSheet
      Color = clWhite
      ImageIndex = 1
      Caption = 'Analyze Result'
      ExplicitWidth = 993
      ExplicitHeight = 0
      object PanelCalcStd: TPanel
        Left = 0
        Top = 0
        Width = 300
        Height = 491
        Align = alLeft
        TabOrder = 0
      end
      object PanelCalcMtrl: TPanel
        Left = 300
        Top = 0
        Width = 475
        Height = 491
        Align = alClient
        TabOrder = 1
        ExplicitWidth = 693
      end
    end
  end
  object PanelNoProp: TPanel
    Left = 0
    Top = 117
    Width = 993
    Height = 28
    Align = alTop
    BevelOuter = bvNone
    Color = 14671839
    ParentBackground = False
    TabOrder = 3
    Visible = False
    StyleElements = [seFont]
    object LabelProperty: THTMLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 987
      Height = 21
      Align = alClient
      AutoSizing = True
      Color = clBtnFace
      FocusControl = PanelNoProp
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      HTMLText.Strings = (
        
          '<B><font color="#E90309">No properties have been defined.</font>' +
          '</B> <a href="OnPropertiesClick">Click here to edit.</a>')
      ParentColor = False
      ParentFont = False
      Transparent = True
      URLColor = 12615680
      VAlignment = tvaCenter
      OnAnchorClick = LabelPropertyAnchorClick
      Version = '1.9.4.0'
      ExplicitLeft = 6
      ExplicitTop = 1
      ExplicitWidth = 873
    end
    object ShapeProperty: TShape
      Left = 0
      Top = 27
      Width = 993
      Height = 1
      Margins.Left = 0
      Margins.Right = 0
      Margins.Bottom = 2
      Align = alBottom
      Pen.Color = 12566463
      ExplicitTop = 3
      ExplicitWidth = 22
    end
  end
  object PanelProperty: TPanel
    Left = 0
    Top = 145
    Width = 217
    Height = 516
    Align = alLeft
    BevelOuter = bvNone
    TabOrder = 4
    Visible = False
  end
  object RzFormState: TRzFormState
    Enabled = False
    RegIniFile = svcOption.Reg
    Left = 240
    Top = 136
  end
  object ActionManager: TActionManager
    Left = 312
    Top = 136
    StyleName = 'Platform Default'
    object ActionRecentTest: TAction
      Caption = 'Recent Tests'
    end
    object ActionOpen: TAction
      Caption = '&Open'
      ShortCut = 16463
      OnExecute = ActionOpenExecute
    end
    object ActionSave: TAction
      Caption = '&Save'
      ShortCut = 16467
      OnExecute = ActionSaveExecute
    end
    object ActionSaveAs: TAction
      Caption = 'Save &As'
      OnExecute = ActionSaveAsExecute
    end
    object ActionSdb: TAction
      Caption = 'SD BIOSENSOR'
      OnExecute = ActionSdbExecute
    end
    object ActionOption: TAction
      Caption = 'O&ption'
      OnExecute = ActionOptionExecute
    end
    object ActionAbout: TAction
      Caption = 'Abou&t'
      OnExecute = ActionAboutExecute
    end
    object ActionExit: TAction
      Caption = 'E&xit'
      OnExecute = ActionExitExecute
    end
    object ActionTabHome: TAction
      Caption = 'Home'
    end
    object ActionGroupData: TAction
      Caption = 'Data'
    end
    object ActionPaste: TAction
      Caption = 'Paste'
      ShortCut = 16470
      OnExecute = ActionPasteExecute
    end
    object ActionTestProperty: TAction
      AutoCheck = True
      Caption = 'Property'
      OnExecute = ActionTestPropertyExecute
    end
    object ActionNewTest: TAction
      Caption = '&New Test'
      ShortCut = 16462
      OnExecute = ActionNewTestExecute
    end
    object ActionManualDataEntry: TAction
      Caption = 'Manual Data Entry'
    end
    object ActionManualTest: TAction
      Caption = 'Write-in test'
      OnExecute = ActionManualTestExecute
    end
    object ActionClearTest: TAction
      Caption = 'Clear All'
      ShortCut = 24622
      OnExecute = ActionClearTestExecute
    end
    object ActionGroupDataFmt: TAction
      Caption = 'Format'
      Enabled = False
    end
    object ActionFmtDefault: TAction
      Caption = 'Default'
      Enabled = False
      OnExecute = ActionSplitButton
    end
    object ActionFmt2Item: TAction
      Tag = 2
      Caption = 'In-Tubu (Nil, Antigen)'
      Enabled = False
      OnExecute = ActionFmtExecute
    end
    object ActionFmt3Item: TAction
      Tag = 3
      Caption = 'In-Tubu (Nil, Antigen, Mitogen)'
      Enabled = False
      OnExecute = ActionFmtExecute
    end
    object ActionFmtManual: TAction
      AutoCheck = True
      Caption = 'Manual'
      Enabled = False
      OnExecute = ActionFmtManualExecute
    end
    object ActionFmtLoad: TAction
      Caption = 'Load'
      OnExecute = ActionFmtLoadExecute
    end
    object ActionFmtSave: TAction
      Caption = 'Save'
      Enabled = False
      OnExecute = ActionFmtSaveExecute
    end
    object ActionFmtViewNames: TAction
      Caption = 'View Names'
      Enabled = False
      OnExecute = ActionFmtViewNamesExecute
    end
    object ActionGroupAnalyze: TAction
      Caption = 'Analyze'
      OnExecute = ActionGroupAnalyzeExecute
    end
    object ActionCalc: TAction
      Caption = 'Calculate'
      OnExecute = ActionCalcExecute
    end
    object ActionPrint: TAction
      Caption = 'Print'
      OnExecute = ActionPrintExecute
    end
    object ActionExportResult: TAction
      Caption = 'Result Export'
      OnExecute = ActionSplitButton
    end
    object ActionResultToClipboard: TAction
      Caption = 'To Clipboard'
      OnExecute = ActionResultToClipboardExecute
    end
    object ActionResultToCsv: TAction
      Caption = 'To CSV File'
      OnExecute = ActionResultToCsvExecute
    end
    object ActionExportData: TAction
      Caption = 'Export Data'
      OnExecute = ActionSplitButton
    end
    object ActionDataToClipboard: TAction
      Caption = 'To Clipboard'
      OnExecute = ActionDataToClipboardExecute
    end
    object ActionDataToCsv: TAction
      Caption = 'To CSV File'
      OnExecute = ActionDataToCsvExecute
    end
    object ActionTabResults: TAction
      Caption = 'Results'
    end
    object ActionTabTest: TAction
      Caption = 'Test'
      OnExecute = ActionTabTestExecute
    end
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      'ActionAbout.Caption'
      'ActionCalc.Caption'
      'ActionClearTest.Caption'
      'ActionDataToClipboard.Caption'
      'ActionDataToCsv.Caption'
      'ActionExit.Caption'
      'ActionExportData.Caption'
      'ActionExportResult.Caption'
      'ActionFmtDefault.Caption'
      'ActionFmtLoad.Caption'
      'ActionFmtManual.Caption'
      'ActionFmtSave.Caption'
      'ActionFmtViewNames.Caption'
      'ActionGroupAnalyze.Caption'
      'ActionGroupData.Caption'
      'ActionGroupDataFmt.Caption'
      'ActionManualDataEntry.Caption'
      'ActionManualTest.Caption'
      'ActionNewTest.Caption'
      'ActionOpen.Caption'
      'ActionOption.Caption'
      'ActionPaste.Caption'
      'ActionPrint.Caption'
      'ActionRecentTest.Caption'
      'ActionResultToClipboard.Caption'
      'ActionResultToCsv.Caption'
      'ActionSave.Caption'
      'ActionSaveAs.Caption'
      'ActionTabHome.Caption'
      'ActionTabResults.Caption'
      'ActionTabTest.Caption'
      'ActionTestProperty.Caption'
      'StatusKitBatch.FieldLabel'
      'StatusOperator.FieldLabel'
      'StatusTestDate.FieldLabel'
      'StatusTestNum.FieldLabel'
      'TabSheetRawData.Caption'
      'TabSheetStdResult.Caption')
    Translatables.Literals = (
      '13BFED02C36A61DEBEC426374F1DEF9D')
    OnAfterTranslate = TranslatorAfterTranslate
    Left = 384
    Top = 136
  end
  object OpenDialog: TOpenDialog
    Filter = 
      'Avaliable Format(*.tbf, *.qft)|*.tbf;*.qff|TB FERON Format (UNIC' +
      'ODE)|*.tbf|QuantiFERON Format (UNICODE)|*.qff'
    Left = 456
    Top = 136
  end
  object SaveDialog: TSaveDialog
    Filter = 
      'TB FERON Format (UNICODE)|*.tbf|QuantiFERON Format (UNICODE)|*.q' +
      'ff'
    Left = 528
    Top = 136
  end
  object TimerInit: TTimer
    Enabled = False
    Interval = 500
    OnTimer = TimerInitTimer
    Left = 592
    Top = 136
  end
  object PngImageList1: TPngImageList
    PngImages = <
      item
        Background = clWindow
        Name = 'Test16x16_1'
        PngImage.Data = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000000AB4944415478DA6364A01030E2100F01E2D940FC018B9C0010DB
          02F1155C068034F703F11D2076C4223F1F881702F1016C068034D703712A10B7
          936A800E109F873AFB07101F01E248520C7000E278204E24106E040D3808A5B1
          0190467B42061442431A1BF8000D60BC065C04E27C1C064C0462FD1110061B81
          D81A88E70271399A0120397F42063402B10410DF00620334031E3040522ACE94
          7898017B064206220C90247E06DD00B200006765381115E530EB000000004945
          4E44AE426082}
      end
      item
        Background = clWindow
        Name = 'Calculate16x16_1'
        PngImage.Data = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000000E54944415478DA6364A01030A2F14380B81BCA1601624B20BE42
          8A01C8200688F581B8945C033880F83A10AB02F11F720C0081D540BC1288D790
          63800610EF07E23340EC4B8E018781B816882703B127103F21C5801C203606E2
          44202E01621620EE20D60019A8D34D81F803104B00F16E20D625D680F5D0805B
          8126D60BC4470819004A48F15802CD078883A15EC2698000109F066257207E80
          A60E140697A1DEFA82CB80F9407C118827E008D87620BE0BC4731820C9DC0488
          77C00C70802AB065C09DEA54A081F902EADA0D405C093360330324CE2F30E007
          11407C02D98B8492324100007AF829116B9577C10000000049454E44AE426082}
      end>
    Left = 656
    Top = 136
    Bitmap = {}
  end
end
