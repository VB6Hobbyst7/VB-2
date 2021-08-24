object vReport: TvReport
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  Caption = 'Print Preview'
  ClientHeight = 658
  ClientWidth = 868
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnMouseWheel = FormMouseWheel
  PixelsPerInch = 96
  TextHeight = 14
  object frxPreview: TfrxPreview
    Left = 249
    Top = 0
    Width = 619
    Height = 658
    Align = alClient
    BackColor = 14671839
    BevelInner = bvNone
    BevelOuter = bvNone
    BorderStyle = bsNone
    OutlineVisible = False
    OutlineWidth = 120
    ThumbnailVisible = True
    UseReportHints = True
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 249
    Height = 658
    Align = alLeft
    BevelOuter = bvNone
    TabOrder = 1
    object ScrollBox1: TScrollBox
      Left = 0
      Top = 153
      Width = 249
      Height = 464
      Align = alClient
      BorderStyle = bsNone
      TabOrder = 0
      DesignSize = (
        249
        464)
      object Label3: TLabel
        Left = 7
        Top = 9
        Width = 65
        Height = 23
        Caption = 'Printer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label4: TLabel
        Left = 7
        Top = 107
        Width = 79
        Height = 23
        Caption = 'Settings'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LabelSubject: TLabel
        Left = 95
        Top = 202
        Width = 46
        Height = 14
        Alignment = taRightJustify
        Caption = 'Subject:'
      end
      object LabelPrinterProperty: TLabel
        Left = 148
        Top = 66
        Width = 95
        Height = 14
        Cursor = crHandPoint
        Alignment = taRightJustify
        Caption = 'Printer Properties'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = 13395510
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = [fsUnderline]
        ParentFont = False
        StyleElements = [seClient, seBorder]
        OnClick = LabelPrinterPropertyClick
      end
      object CheckImgPrint: TRzCheckBox
        Left = 7
        Top = 136
        Width = 237
        Height = 29
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        Caption = 'Print Standard Curve and Plate Formatting'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        State = cbUnchecked
        TabOrder = 0
        WordWrap = True
        OnClick = CheckImgPrintClick
      end
      object ComboSubID: TComboBox
        Left = 148
        Top = 198
        Width = 95
        Height = 22
        Style = csDropDownList
        Anchors = [akTop, akRight]
        Enabled = False
        TabOrder = 1
        OnChange = ComboSubIDChange
      end
      object ComboPrinter: TComboBox
        Left = 7
        Top = 38
        Width = 236
        Height = 22
        Style = csDropDownList
        TabOrder = 2
        OnClick = ComboPrinterClick
      end
      object ComboReportType: TComboBox
        Left = 7
        Top = 171
        Width = 236
        Height = 22
        Style = csDropDownList
        ItemIndex = 0
        TabOrder = 3
        Text = 'All Subjects (Group Report)'
        OnClick = ComboReportTypeClick
        Items.Strings = (
          'All Subjects (Group Report)'
          'All Subjects (Individual Report)'
          'Single Subject Report')
      end
    end
    object Panel2: TPanel
      Left = 0
      Top = 0
      Width = 249
      Height = 153
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      object Label1: TLabel
        Left = 7
        Top = 8
        Width = 51
        Height = 25
        Caption = 'Print'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -21
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label2: TLabel
        Left = 147
        Top = 49
        Width = 39
        Height = 14
        Alignment = taRightJustify
        Caption = 'Copies:'
      end
      object EditCopies: TSpinEdit
        Left = 192
        Top = 45
        Width = 51
        Height = 23
        MaxValue = 99
        MinValue = 1
        TabOrder = 0
        Value = 1
      end
      object ButtonPDF: TRzBitBtn
        Left = 8
        Top = 80
        Width = 133
        Height = 35
        Alignment = taLeftJustify
        Caption = 'Save as PDF'
        TabOrder = 1
        OnClick = ButtonPDFClick
        ImageIndex = 0
        Images = Imgs
        Layout = blGlyphRight
        Spacing = 0
      end
      object ButtonPrint: TRzBitBtn
        Left = 8
        Top = 39
        Width = 133
        Height = 35
        Alignment = taLeftJustify
        Caption = 'Print'
        TabOrder = 2
        OnClick = ButtonPrintClick
        ImageIndex = 1
        Images = Imgs
        Layout = blGlyphRight
        Spacing = 0
      end
    end
    object Panel3: TPanel
      Left = 0
      Top = 617
      Width = 249
      Height = 41
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 2
      DesignSize = (
        249
        41)
      object Button5: TButton
        Left = 7
        Top = 6
        Width = 75
        Height = 25
        Anchors = [akTop, akRight]
        Caption = '&Close'
        ModalResult = 1
        TabOrder = 0
      end
    end
  end
  object MemoM3: TMemo
    Left = 520
    Top = 248
    Width = 185
    Height = 89
    Lines.Strings = (
      
        '<b>Nil control must be '#8804' 8.0 IU/mL and Mitogen - Nil must be '#8805' 0' +
        '.5 IU/mL OR TB Antigen - nil must be '#8805' 0.35 IU/mL for a subject ' +
        'to have a valid ELISA Report result.</b>'
      ''
      
        'The Mitogen control generally elicits the greatest IFN-gamma res' +
        'ponse of the 3 samples from each subject. In some cases, the Mit' +
        'ogen control OD value will be above the limit of the microplate ' +
        'reader; this has no impact on the test interpretation. The IFN-g' +
        'amma level of the Nil control is considered background and is su' +
        'btracted from the TB Antigen and Mitogen results for that blood ' +
        'specimen. In clinical studies, less than 0.25% of subjects had I' +
        'FN-gamma levels of > 8.0 IU/mL for the Nil control.'
      ''
      
        'The cut-off for the ELISA Product test is 0.35 IU/mL above the N' +
        'il control (and TB Antigen minus Nil is '#8805' 25% of the Nil control' +
        ') for the TB Antigen stimulated plasma sample. Individuals displ' +
        'aying a response to the TB Antigen above this cut-off are likely' +
        ' to be infected with <i>M.tuberculosis</i>.'
      ''
      
        'The magnitude of the measured IFN-gamma level cannot be correlat' +
        'ed with stage or degree of infection, level of immune responsive' +
        'ness, or likelihood for progression to active disease. A positiv' +
        'e ELISA Report result does not necessarily indicate the presence' +
        ' of active tuberculosis disease. Other diagnostic procedures, su' +
        'ch as X-ray examination of the chest and microbiological examina' +
        'tion of sputum, should be used when TB disease is suspected.'
      ''
      
        'More detailed information can be found in the "Interpretation of' +
        ' Result'#39' section of then ELISA Product Package Insert.')
    ScrollBars = ssBoth
    TabOrder = 2
    Visible = False
  end
  object MemoM2: TMemo
    Left = 520
    Top = 153
    Width = 185
    Height = 89
    Lines.Strings = (
      
        '<b>Nil control must be '#8804' 8.0 IU/mL for a subject to have a valid' +
        ' ELISA Report result.</b>'
      ''
      
        'The IFN-gamma level of the Nil control is considered background ' +
        'and is subtracted from the TB Antigen result for that blood spec' +
        'imen. In clinical studies, less than 0.25% of subjects had IFN-g' +
        'amma levels of > 8.0 IU/mL for the Nil control.'
      ''
      
        'The cut-off for the ELISA Product test is 0.35 IU/mL above the N' +
        'il control (and TB Antigen minus Nil is '#8805' 25% of the Nil control' +
        ') for the TB Antigen stimulated plasma sample. Individuals displ' +
        'aying a response to the TB Antigen above this cut-off are likely' +
        ' to be infected with <i>M. tuberculosis</i>.'
      ''
      
        'The magnitude of the measured IFN-gamma level cannot be correlat' +
        'ed with stage or degree of infection, level of immune responsive' +
        'ness, or likelihood for progression to active disease. A positiv' +
        'e ELISA Report result does not necessarily indicate the presence' +
        ' of active tuberculosis disease. Other diagnostic procedures, su' +
        'ch as X-ray examination of the chest and microbiological examina' +
        'tion of sputum, should be used when TB disease is suspected.'
      ''
      
        'Where a poor response due to an immunosuppressive condition or m' +
        'edication is suspected, a Mitogen control may be used to monitor' +
        ' a subject'#39's capacity to procedure IFN-gamma. This can be achiev' +
        'ed by repreating the test incorporating a Mitogen control tube.'
      ''
      
        'More detailed information can be found in the "Interpretation of' +
        ' Result'#39' section of then ELISA Product Package Insert.')
    ScrollBars = ssBoth
    TabOrder = 3
    Visible = False
  end
  object frxReport: TfrxReport
    Version = '6.1.12'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    Preview = frxPreview
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.MDIChild = True
    PreviewOptions.ThumbnailVisible = True
    PreviewOptions.ShowCaptions = True
    PreviewOptions.Zoom = 1.000000000000000000
    PreviewOptions.ZoomMode = zmPageWidth
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    ReportOptions.CreateDate = 42828.407864976900000000
    ReportOptions.LastChange = 42909.487180555550000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      ''
      'end.')
    OnBeforePrint = frxReportBeforePrint
    OnGetValue = frxReportGetValue
    Left = 392
    Top = 32
    Datasets = <
      item
        DataSet = frxRawDataDataSet
        DataSetName = 'ODDataSet'
      end
      item
        DataSet = frxResultSet
        DataSetName = 'RangeDataSet'
      end
      item
        DataSet = frxStdDataSet
        DataSetName = 'StdDataSet'
      end>
    Variables = <>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object PageResult: TfrxReportPage
      PaperWidth = 210.000000000000000000
      PaperHeight = 297.000000000000000000
      PaperSize = 9
      LeftMargin = 10.001250000000000000
      RightMargin = 10.001250000000000000
      TopMargin = 10.001250000000000000
      BottomMargin = 10.001250000000000000
      Duplex = dmVertical
      Frame.Typ = []
      LargeDesignHeight = True
      object PageResultHeader: TfrxPageHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 154.000000000000000000
        Top = 18.897650000000000000
        Width = 718.101251175000000000
        Stretched = True
        object LabelTitle: TfrxMemoView
          AllowVectorExport = True
          Left = 229.417440000000000000
          Top = 10.000000000000000000
          Width = 486.488250000000000000
          Height = 34.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -32
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'ELISA Report Result')
          ParentFont = False
        end
        object PictureLogo: TfrxPictureView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Top = 1.763760000000001000
          Width = 200.000000000000000000
          Height = 80.000000000000000000
          Frame.Typ = []
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object LabelRunDate: TfrxMemoView
          AllowVectorExport = True
          Left = 240.440940000000000000
          Top = 49.204700000000000000
          Width = 168.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Run Date:')
          ParentFont = False
        end
        object LabelOperator: TfrxMemoView
          AllowVectorExport = True
          Left = 240.440940000000000000
          Top = 69.204700000000000000
          Width = 168.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Operator:')
          ParentFont = False
        end
        object LabelRunNumber: TfrxMemoView
          AllowVectorExport = True
          Left = 240.440940000000000000
          Top = 89.204700000000000000
          Width = 168.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Run Number:')
          ParentFont = False
        end
        object LabelKitBatchNumber: TfrxMemoView
          AllowVectorExport = True
          Left = 240.440940000000000000
          Top = 109.204700000000000000
          Width = 168.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Kit Batch Number:')
          ParentFont = False
        end
        object MemoRunDate: TfrxMemoView
          AllowVectorExport = True
          Left = 416.440940000000000000
          Top = 49.204700000000000000
          Width = 296.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[RunDate]')
          ParentFont = False
        end
        object MemoOperator: TfrxMemoView
          AllowVectorExport = True
          Left = 416.440940000000000000
          Top = 69.204700000000000000
          Width = 296.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[Operator]')
          ParentFont = False
        end
        object MemoRunNumber: TfrxMemoView
          AllowVectorExport = True
          Left = 416.440940000000000000
          Top = 89.204700000000000000
          Width = 296.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[RunNumber]')
          ParentFont = False
        end
        object MemoKitBatchNumber: TfrxMemoView
          AllowVectorExport = True
          Left = 416.440940000000000000
          Top = 109.204700000000000000
          Width = 296.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[KitBatchNumber]')
          ParentFont = False
        end
        object MemoResultDesc: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Top = 135.000000000000000000
          Width = 664.000000000000000000
          Height = 19.000000000000000000
          StretchMode = smActualHeight
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[ResultDesc]')
          ParentFont = False
        end
        object MemoExeVer: TfrxMemoView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Top = 88.000000000000000000
          Width = 200.000000000000000000
          Height = 24.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[ExeVer]')
          ParentFont = False
        end
        object SysMemoPage: TfrxSysMemoView
          AllowVectorExport = True
          Left = 624.000000000000000000
          Width = 88.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[PAGE#] / [TOTALPAGES#]')
          ParentFont = False
        end
      end
      object ResultData: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 17.000000000000000000
        Top = 313.700990000000000000
        Width = 718.101251175000000000
        DataSet = frxResultSet
        DataSetName = 'RangeDataSet'
        KeepHeader = True
        RowCount = 0
        object MemoSubjID: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Width = 184.000000000000000000
          Height = 17.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          GapX = 5.000000000000000000
          Memo.UTF8W = (
            '[SubjectID]')
          ParentFont = False
        end
        object MemoNil: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 208.000000000000000000
          Width = 78.000000000000000000
          Height = 17.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Nil]')
          ParentFont = False
        end
        object MemoTBAg: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 286.000000000000000000
          Width = 78.000000000000000000
          Height = 17.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[TBAg]')
          ParentFont = False
        end
        object MemoMitogen: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 364.000000000000000000
          Width = 78.000000000000000000
          Height = 17.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Mitogen]')
          ParentFont = False
        end
        object MemoDiffTBAgNil: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 442.000000000000000000
          Width = 78.000000000000000000
          Height = 17.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[DiffTBAgNil]')
          ParentFont = False
        end
        object MemoDiffMtgNil: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 520.000000000000000000
          Width = 78.000000000000000000
          Height = 17.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[DiffMtgNil]')
          ParentFont = False
        end
        object MemoResult: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 598.000000000000000000
          Width = 99.000000000000000000
          Height = 17.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftBottom]
          GapX = 5.000000000000000000
          Memo.UTF8W = (
            '[Result]')
          ParentFont = False
        end
        object Line8: TfrxLineView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Height = 17.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line9: TfrxLineView
          AllowVectorExport = True
          Left = 16.000000000000000000
          Height = 17.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line10: TfrxLineView
          AllowVectorExport = True
          Left = 704.000000000000000000
          Height = 17.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line11: TfrxLineView
          AllowVectorExport = True
          Left = 712.000000000000000000
          Height = 17.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
      end
      object ReportSummary: TfrxReportSummary
        FillType = ftBrush
        Frame.Typ = []
        Height = 848.000000000000000000
        Top = 517.795610000000000000
        Width = 718.101251175000000000
        object PictureGrid: TfrxPictureView
          AllowVectorExport = True
          Top = 440.000000000000000000
          Width = 712.000000000000000000
          Height = 408.000000000000000000
          Frame.Color = clNone
          Frame.Typ = []
          FillType = ftGradient
          Fill.EndColor = clNone
          Fill.GradientStyle = gsHorizontal
          Picture.Data = {
            0A54504E474F626A65637489504E470D0A1A0A0000000D49484452000000C800
            0000500802000000D30FADE7000000017352474200AECE1CE90000000467414D
            410000B18F0BFC6105000000097048597300000EC300000EC301C76FA8640000
            02264944415478DAEDDAB172DA401846515C2694A176ED9449CDDB43ED9471ED
            9A94B8CE100721AC95CC025F1CA4732ACFF2B36646773412E86EB55ACDE0DAEE
            844582B088101611C222425844088B086111212C22844584B0887827ACE572F9
            D19F90FFDA7ABD2EAE0B8B8B088B086111212C22844584B088101611C2224258
            44088B086111212C22844584B088101611C22262F461FD7A5AFFDCFCF96BF175
            F9F0E5A33FCE64088B086111212C228455F0F2FCE3F179DB5A1878EBD1EC6EEE
            D37E617EFFFDDBFDE7F377BE6DC2EA9D7FA3D0492793DDD86231DB6C4A61D5EC
            7CFB84D5D62AA5992EADBD196ECA386AE7A897AA9DC7405827CC36EBED568A8B
            ED588AC3A7EC3C0AC22A8C768E72E195BEE1C279AC72E7711056E3904467B2FB
            52FF702195AA9D4742588DA163DCA9E55A618DF69425ACEE64E1105F1656D5CE
            2321AC8633D63509AB5177255473F1EE1AEB608261D5DDBB15EFFEFABE6E7057
            D89862587DB33D0D1DA6DFFF82B46EE73110565BE5F7E3A51F69E6F3F9763B78
            26F3CDFBC0DB461AD6ACFE17BDA3F9DDBF98ED174EC97068E7DB36A1B0FA748F
            EB25CF20BC0C3DDCE0E98669877581D1DEE5D51A7D58FF96C70AF78475AED773
            53BB9F9E271BA64958E72A3DE5F7D7D4CF563BC2BA44E1024E54AF844584B088
            101611C222425844088B086111212C22844584B088101611C222425844088B08
            6111212C228445C49961C179844584B088101611C222425844088B086111212C
            22844584B088F80DA2FDA24E7198A7FD0000000049454E44AE426082}
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object PictureChart: TfrxPictureView
          AllowVectorExport = True
          Width = 592.000000000000000000
          Height = 432.000000000000000000
          Frame.Typ = []
          FillType = ftGradient
          Fill.EndColor = clNone
          Fill.GradientStyle = gsHorizontal
          Picture.Data = {
            0A54504E474F626A65637489504E470D0A1A0A0000000D49484452000000C800
            0000500802000000D30FADE7000000017352474200AECE1CE90000000467414D
            410000B18F0BFC6105000000097048597300000EC300000EC301C76FA8640000
            02264944415478DAEDDAB172DA401846515C2694A176ED9449CDDB43ED9471ED
            9A94B8CE100721AC95CC025F1CA4732ACFF2B36646773412E86EB55ACDE0DAEE
            844582B088101611C222425844088B086111212C22844584B0887827ACE572F9
            D19F90FFDA7ABD2EAE0B8B8B088B086111212C22844584B088101611C2224258
            44088B086111212C22844584B088101611C22262F461FD7A5AFFDCFCF96BF175
            F9F0E5A33FCE64088B086111212C228455F0F2FCE3F179DB5A1878EBD1EC6EEE
            D37E617EFFFDDBFDE7F377BE6DC2EA9D7FA3D0492793DDD86231DB6C4A61D5EC
            7CFB84D5D62AA5992EADBD196ECA386AE7A897AA9DC7405827CC36EBED568A8B
            ED588AC3A7EC3C0AC22A8C768E72E195BEE1C279AC72E7711056E3904467B2FB
            52FF702195AA9D4742588DA163DCA9E55A618DF69425ACEE64E1105F1656D5CE
            2321AC8633D63509AB5177255473F1EE1AEB608261D5DDBB15EFFEFABE6E7057
            D89862587DB33D0D1DA6DFFF82B46EE73110565BE5F7E3A51F69E6F3F9763B78
            26F3CDFBC0DB461AD6ACFE17BDA3F9DDBF98ED174EC97068E7DB36A1B0FA748F
            EB25CF20BC0C3DDCE0E98669877581D1DEE5D51A7D58FF96C70AF78475AED773
            53BB9F9E271BA64958E72A3DE5F7D7D4CF563BC2BA44E1024E54AF844584B088
            101611C222425844088B086111212C22844584B088101611C222425844088B08
            6111212C228445C49961C179844584B088101611C222425844088B086111212C
            22844584B088F80DA2FDA24E7198A7FD0000000049454E44AE426082}
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
      end
      object FooterResultData: TfrxFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 56.000000000000000000
        Top = 355.275820000000000000
        Width = 718.101251175000000000
        Child = frxReport.ChildSubReport
        Stretched = True
        object Line2: TfrxLineView
          AllowVectorExport = True
          Left = 16.000000000000000000
          Top = 4.000000000000000000
          Width = 688.000000000000000000
          Color = clBlack
          Frame.Typ = [ftTop]
        end
        object Line4: TfrxLineView
          AllowVectorExport = True
          Left = 712.000000000000000000
          Height = 12.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line6: TfrxLineView
          AllowVectorExport = True
          Left = 16.000000000000000000
          Height = 4.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line7: TfrxLineView
          AllowVectorExport = True
          Left = 704.000000000000000000
          Height = 4.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line3: TfrxLineView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Top = 12.000000000000000000
          Width = 704.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line5: TfrxLineView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Height = 12.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object LabelSig: TfrxMemoView
          AllowVectorExport = True
          Left = 16.000000000000000000
          Top = 24.000000000000000000
          Width = 104.000000000000000000
          Height = 19.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Signature')
          ParentFont = False
        end
        object MemoSig: TfrxMemoView
          AllowVectorExport = True
          Left = 120.000000000000000000
          Top = 24.000000000000000000
          Width = 230.000000000000000000
          Height = 19.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftBottom]
          ParentFont = False
        end
        object LabelDate: TfrxMemoView
          AllowVectorExport = True
          Left = 350.101251180000000000
          Top = 24.000000000000000000
          Width = 128.000000000000000000
          Height = 19.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Date')
          ParentFont = False
        end
        object MemoDate: TfrxMemoView
          AllowVectorExport = True
          Left = 478.101251180000000000
          Top = 24.000000000000000000
          Width = 230.000000000000000000
          Height = 19.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftBottom]
          ParentFont = False
        end
      end
      object HeaderResultData: TfrxHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 56.000000000000000000
        Top = 234.330860000000000000
        Width = 718.101251175000000000
        object LabelSubjectID: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Top = 38.000000000000030000
          Width = 184.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          GapX = 5.000000000000000000
          Memo.UTF8W = (
            'Subject ID')
          ParentFont = False
        end
        object LabelNil: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 208.000000000000000000
          Top = 38.000000000000030000
          Width = 78.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Nil')
          ParentFont = False
        end
        object LabelTBAg: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 286.000000000000000000
          Top = 38.000000000000030000
          Width = 78.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'TB Ag')
          ParentFont = False
        end
        object LabelMitogen: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 364.000000000000000000
          Top = 38.000000000000030000
          Width = 78.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Mitogen')
          ParentFont = False
        end
        object LabelDiffTBAgNil: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 442.000000000000000000
          Top = 38.000000000000030000
          Width = 78.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'TB Ag-Nil')
          ParentFont = False
        end
        object LabelDiffMtgNil: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 520.000000000000000000
          Top = 38.000000000000030000
          Width = 78.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Mitogen-Nil')
          ParentFont = False
        end
        object LabelResult: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 598.000000000000000000
          Top = 38.000000000000030000
          Width = 99.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftBottom]
          GapX = 5.000000000000000000
          Memo.UTF8W = (
            'Result')
          ParentFont = False
        end
        object LabelResultTitle: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Top = 19.000000000000000000
          Width = 648.000000000000000000
          Height = 19.000000000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Results (IU/mL)')
          ParentFont = False
        end
        object Line13: TfrxLineView
          AllowVectorExport = True
          Left = 16.000000000000000000
          Top = 16.000000000000000000
          Height = 40.000000000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Line12: TfrxLineView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Top = 8.000000000000000000
          Height = 48.000000000000000000
          Color = clBlack
          Frame.Typ = [ftLeft]
        end
        object Line15: TfrxLineView
          AllowVectorExport = True
          Left = 16.000000000000000000
          Top = 16.000000000000000000
          Width = 688.000000000000000000
          Color = clBlack
          Frame.Typ = [ftTop]
        end
        object Line16: TfrxLineView
          AllowVectorExport = True
          Left = 704.000000000000000000
          Top = 16.000000000000000000
          Height = 40.000000000000000000
          Color = clBlack
          Frame.Typ = [ftLeft]
        end
        object Line18: TfrxLineView
          AllowVectorExport = True
          Left = 712.000000000000000000
          Top = 8.000000000000000000
          Height = 48.000000000000000000
          Color = clBlack
          Frame.Typ = [ftLeft]
        end
        object Line14: TfrxLineView
          AllowVectorExport = True
          Left = 8.000000000000000000
          Top = 8.000000000000000000
          Width = 704.000000000000000000
          Color = clBlack
          Frame.Typ = [ftTop]
        end
      end
      object ChildSubReport: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 24.000000000000000000
        Top = 434.645950000000000000
        Width = 718.101251175000000000
        StartNewPage = True
        ToNRows = 0
        ToNRowsMode = rmCount
        object Subreport: TfrxSubreport
          Align = baWidth
          AllowVectorExport = True
          Top = 8.000000000000000000
          Width = 718.101251175000000000
          Height = 16.000000000000000000
          Page = frxReport.Page1
          PrintOnParent = True
        end
      end
    end
    object Page1: TfrxReportPage
      PaperWidth = 210.000000000000000000
      PaperHeight = 297.000000000000000000
      PaperSize = 9
      LeftMargin = 10.001250000000000000
      RightMargin = 10.001250000000000000
      TopMargin = 10.001250000000000000
      BottomMargin = 10.001250000000000000
      Duplex = dmVertical
      Frame.Typ = []
      LargeDesignHeight = True
      object HeaderStd: TfrxHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 16.000000000000000000
        Top = 18.897650000000000000
        Width = 718.101251175000000000
        ReprintOnNewPage = True
        object Memo1: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Std')
          ParentFont = False
        end
        object Memo2: TfrxMemoView
          AllowVectorExport = True
          Left = 120.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Conc')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          AllowVectorExport = True
          Left = 312.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Mean')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          AllowVectorExport = True
          Left = 216.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '% CV')
          ParentFont = False
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 408.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'QC Result')
          ParentFont = False
        end
      end
      object FooterStd: TfrxFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 26.000000000000000000
        Top = 94.488250000000000000
        Width = 718.101251175000000000
        object LabelIntercept: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Top = 3.000000000000000000
          Width = 96.000000000000000000
          Height = 18.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Intercept: ')
          ParentFont = False
        end
        object MemoIntercept: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 120.000000000000000000
          Top = 3.000000000000000000
          Width = 96.000000000000000000
          Height = 18.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[Intercept]')
          ParentFont = False
        end
        object LabelSlope: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 216.000000000000000000
          Top = 3.000000000000000000
          Width = 96.000000000000000000
          Height = 18.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Slope: ')
          ParentFont = False
        end
        object LabelCorreCoef: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 408.000000000000000000
          Top = 3.000000000000000000
          Width = 160.000000000000000000
          Height = 18.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Correlation Coefficient: ')
          ParentFont = False
        end
        object MemoCorreCoef: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 568.000000000000000000
          Top = 3.000000000000000000
          Width = 96.000000000000000000
          Height = 18.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[CorrelCoef]')
          ParentFont = False
        end
        object MemoSlope: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 312.000000000000000000
          Top = 3.000000000000000000
          Width = 96.000000000000000000
          Height = 18.000000000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[Slope]')
          ParentFont = False
        end
      end
      object MasterDataStd: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 16.000000000000000000
        Top = 56.692950000000000000
        Width = 718.101251175000000000
        DataSet = frxStdDataSet
        DataSetName = 'StdDataSet'
        KeepFooter = True
        KeepHeader = True
        RowCount = 0
        object MemoConc: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 120.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[Conc]')
          ParentFont = False
        end
        object MemoCV: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 312.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[CV]')
          ParentFont = False
        end
        object MemoMean: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 216.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[Mean]')
          ParentFont = False
        end
        object MemoQCResult: TfrxMemoView
          Align = baLeft
          AllowVectorExport = True
          Left = 408.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[QCResult]')
          ParentFont = False
        end
        object Memo19: TfrxMemoView
          AllowVectorExport = True
          Left = 24.000000000000000000
          Width = 96.000000000000000000
          Height = 16.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[Std]')
          ParentFont = False
        end
      end
      object FooterOD: TfrxFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 366.000000000000000000
        Top = 249.448980000000000000
        Width = 718.101251175000000000
        Stretched = True
        object Line1: TfrxLineView
          AllowVectorExport = True
          Top = 8.000000000000000000
          Width = 720.000000000000000000
          Color = clBlack
          Frame.Style = fsDash
          Frame.Typ = [ftTop]
        end
        object MemoMatrialNote: TfrxMemoView
          AllowVectorExport = True
          Top = 16.000000000000000000
          Width = 712.000000000000000000
          Height = 64.000000000000000000
          StretchMode = smActualHeight
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          LineSpacing = 0.200000000000000000
          Memo.UTF8W = (
            'ELISA Report results are interpreted as follows:'
            ''
            
              '<b>NOTE: Diagnosing or excluding tuberculosis disease, and asses' +
              'sing the probability of LTBI, Requires a combination of epidemio' +
              'logical, historical, medical, and diagnostic findings that shoul' +
              'd be taken into account when interpreting ELISA Report results.<' +
              '/b>')
          ParentFont = False
        end
        object M2Memo68: TfrxMemoView
          AllowVectorExport = True
          Top = 84.000000000000000000
          Width = 75.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Nil'
            '(IU/mL)')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo69: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 84.000000000000000000
          Width = 187.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'TB Antigen minus Nil'
            '(IU/mL)')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo70: TfrxMemoView
          AllowVectorExport = True
          Left = 362.000000000000000000
          Top = 84.000000000000000000
          Width = 220.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Report/Intepretation')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo71: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 84.000000000000000000
          Width = 100.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ELISA Report'
            'Result')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo73: TfrxMemoView
          AllowVectorExport = True
          Top = 118.000000000000000000
          Width = 75.000000000000000000
          Height = 54.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8804' 8.0')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo74: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 118.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '< 0.35')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo76: TfrxMemoView
          AllowVectorExport = True
          Left = 362.000000000000000000
          Top = 118.000000000000000000
          Width = 220.000000000000000000
          Height = 36.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '<I>M.tuberculosis</i> infection NOT likely')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo77: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 118.000000000000000000
          Width = 100.000000000000000000
          Height = 36.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Negative')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo78: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 136.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.35 and < 25% of Nil value')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo80: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 154.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.35 and '#8805' 25% of Nil value')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo86: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 172.000000000000000000
          Width = 100.000000000000000000
          Height = 36.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Indeterminate')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo87: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 154.000000000000000000
          Width = 100.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Positive')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo88: TfrxMemoView
          AllowVectorExport = True
          Left = 362.000000000000000000
          Top = 154.000000000000000000
          Width = 220.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '<i>M.tuberculosis</i> infection likely')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo89: TfrxMemoView
          AllowVectorExport = True
          Left = 362.000000000000000000
          Top = 172.000000000000000000
          Width = 220.000000000000000000
          Height = 36.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Result are indeterminate for'
            'TB Antigen responsiveness')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo90: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 172.000000000000000000
          Width = 187.000000000000000000
          Height = 36.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Any')
          ParentFont = False
          VAlign = vaCenter
        end
        object M2Memo92: TfrxMemoView
          AllowVectorExport = True
          Top = 172.000000000000000000
          Width = 75.000000000000000000
          Height = 36.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '>8.0')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo42: TfrxMemoView
          AllowVectorExport = True
          Top = 208.000000000000000000
          Width = 75.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Nil'
            '(IU/mL)')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo43: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 208.000000000000000000
          Width = 187.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'TB Antigen minus Nil'
            '(IU/mL)')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo44: TfrxMemoView
          AllowVectorExport = True
          Left = 494.000000000000000000
          Top = 208.000000000000000000
          Width = 220.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Report/Intepretation')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo45: TfrxMemoView
          AllowVectorExport = True
          Left = 394.000000000000000000
          Top = 208.000000000000000000
          Width = 100.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ELISA Report'
            'Result')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo46: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 208.000000000000000000
          Width = 132.000000000000000000
          Height = 34.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Mitogen minus Nil'
            '(IU/mL)')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo47: TfrxMemoView
          AllowVectorExport = True
          Top = 242.000000000000000000
          Width = 75.000000000000000000
          Height = 90.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8804' 8.0')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo48: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 242.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '< 0.35')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo49: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 242.000000000000000000
          Width = 132.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.5')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo50: TfrxMemoView
          AllowVectorExport = True
          Left = 494.000000000000000000
          Top = 242.000000000000000000
          Width = 220.000000000000000000
          Height = 36.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '<I>M.tuberculosis</i> infection NOT likely')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo51: TfrxMemoView
          AllowVectorExport = True
          Left = 394.000000000000000000
          Top = 242.000000000000000000
          Width = 100.000000000000000000
          Height = 36.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Negative')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo52: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 260.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.35 and < 25% of Nil value')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo53: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 260.000000000000000000
          Width = 132.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.5')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo54: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 278.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.35 and '#8805' 25% of Nil value')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo55: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 278.000000000000000000
          Width = 132.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Any')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo56: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 296.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '< 0.35')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo57: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 296.000000000000000000
          Width = 132.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '< 0.5')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo58: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 314.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            #8805' 0.35 and < 25% of Nil value')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo59: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 314.000000000000000000
          Width = 132.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '< 0.5')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo61: TfrxMemoView
          AllowVectorExport = True
          Left = 394.000000000000000000
          Top = 296.000000000000000000
          Width = 100.000000000000000000
          Height = 54.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Indeterminate')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo62: TfrxMemoView
          AllowVectorExport = True
          Left = 394.000000000000000000
          Top = 278.000000000000000000
          Width = 100.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Positive')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo63: TfrxMemoView
          AllowVectorExport = True
          Left = 494.000000000000000000
          Top = 278.000000000000000000
          Width = 220.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '<i>M.tuberculosis</i> infection likely')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo64: TfrxMemoView
          AllowVectorExport = True
          Left = 494.000000000000000000
          Top = 296.000000000000000000
          Width = 220.000000000000000000
          Height = 54.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Result are indeterminate for'
            'TB Antigen responsiveness')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo65: TfrxMemoView
          AllowVectorExport = True
          Left = 75.000000000000000000
          Top = 332.000000000000000000
          Width = 187.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Any')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo66: TfrxMemoView
          AllowVectorExport = True
          Left = 262.000000000000000000
          Top = 332.000000000000000000
          Width = 132.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'Any')
          ParentFont = False
          VAlign = vaCenter
        end
        object M3Memo67: TfrxMemoView
          AllowVectorExport = True
          Top = 332.000000000000000000
          Width = 75.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '>8.0')
          ParentFont = False
          VAlign = vaCenter
        end
        object MemoMtrlDesc: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 350.000000000000000000
          Width = 718.101251175000000000
          Height = 16.000000000000000000
          StretchMode = smActualHeight
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          LineSpacing = 1.000000000000000000
          Memo.UTF8W = (
            '[MaterialDesc]')
          ParentFont = False
        end
      end
      object HeaderOD: TfrxHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 42.000000000000000000
        Top = 143.622140000000000000
        Width = 718.101251175000000000
        ReprintOnNewPage = True
        object LavelRawData: TfrxMemoView
          AllowVectorExport = True
          Width = 664.000000000000000000
          Height = 19.000000000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Raw Data(OD)')
          ParentFont = False
        end
        object LabelOD0: TfrxMemoView
          AllowVectorExport = True
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          ParentFont = False
        end
        object LabelOD1: TfrxMemoView
          AllowVectorExport = True
          Left = 54.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '1')
          ParentFont = False
        end
        object LabelOD2: TfrxMemoView
          AllowVectorExport = True
          Left = 108.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '2')
          ParentFont = False
        end
        object LabelOD3: TfrxMemoView
          AllowVectorExport = True
          Left = 162.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '3')
          ParentFont = False
        end
        object LabelOD4: TfrxMemoView
          AllowVectorExport = True
          Left = 216.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '4')
          ParentFont = False
        end
        object LabelOD5: TfrxMemoView
          AllowVectorExport = True
          Left = 270.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '5')
          ParentFont = False
        end
        object LabelOD6: TfrxMemoView
          AllowVectorExport = True
          Left = 324.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '6')
          ParentFont = False
        end
        object LabelOD7: TfrxMemoView
          AllowVectorExport = True
          Left = 378.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '7')
          ParentFont = False
        end
        object LabelOD8: TfrxMemoView
          AllowVectorExport = True
          Left = 432.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '8')
          ParentFont = False
        end
        object LabelOD9: TfrxMemoView
          AllowVectorExport = True
          Left = 486.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '9')
          ParentFont = False
        end
        object LabelOD10: TfrxMemoView
          AllowVectorExport = True
          Left = 540.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '10')
          ParentFont = False
        end
        object LabelOD11: TfrxMemoView
          AllowVectorExport = True
          Left = 594.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '11')
          ParentFont = False
        end
        object LabelOD12: TfrxMemoView
          AllowVectorExport = True
          Left = 648.000000000000000000
          Top = 24.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '12')
          ParentFont = False
        end
      end
      object MasterDataOD: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 20.000000000000000000
        Top = 207.874150000000000000
        Width = 718.101251175000000000
        DataSet = frxRawDataDataSet
        DataSetName = 'ODDataSet'
        KeepFooter = True
        KeepHeader = True
        RowCount = 0
        object MemoOD0: TfrxMemoView
          AllowVectorExport = True
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD00]')
          ParentFont = False
        end
        object MemoOD1: TfrxMemoView
          AllowVectorExport = True
          Left = 54.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD01]')
          ParentFont = False
        end
        object MemoOD2: TfrxMemoView
          AllowVectorExport = True
          Left = 108.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD02]')
          ParentFont = False
        end
        object MemoOD3: TfrxMemoView
          AllowVectorExport = True
          Left = 162.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD03]')
          ParentFont = False
        end
        object MemoOD4: TfrxMemoView
          AllowVectorExport = True
          Left = 216.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD04]')
          ParentFont = False
        end
        object MemoOD5: TfrxMemoView
          AllowVectorExport = True
          Left = 270.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD05]')
          ParentFont = False
        end
        object MemoOD6: TfrxMemoView
          AllowVectorExport = True
          Left = 324.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD06]')
          ParentFont = False
        end
        object MemoOD7: TfrxMemoView
          AllowVectorExport = True
          Left = 378.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD07]')
          ParentFont = False
        end
        object MemoOD8: TfrxMemoView
          AllowVectorExport = True
          Left = 432.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD08]')
          ParentFont = False
        end
        object MemoOD9: TfrxMemoView
          AllowVectorExport = True
          Left = 486.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD09]')
          ParentFont = False
        end
        object MemoOD10: TfrxMemoView
          AllowVectorExport = True
          Left = 540.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD10]')
          ParentFont = False
        end
        object MemoOD11: TfrxMemoView
          AllowVectorExport = True
          Left = 594.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD11]')
          ParentFont = False
        end
        object MemoOD12: TfrxMemoView
          AllowVectorExport = True
          Left = 648.000000000000000000
          Width = 54.000000000000000000
          Height = 18.000000000000000000
          AllowHTMLTags = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[OD12]')
          ParentFont = False
        end
      end
    end
  end
  object Imgs: TPngImageList
    Height = 32
    Width = 32
    PngImages = <
      item
        Background = clWindow
        Name = 'Adobe-Acrobat_32x32_4'
        PngImage.Data = {
          89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7A
          F4000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000001CD4944415478DAED96C12B047114C7DFD62A6A0F0E0EAB2865D5
          0AB17120A9510E94839250B61C284228CA4591552EDAC3BA888B288A72F11FD8
          E2E0A038382829872D0E8E0E0EBEAFF7266252B3667E53DA579FF630ED6F3EBF
          F77EEFCD2F440147A820F05F04C2E014EC8193200446C03EC881F22004164033
          180295E0C9A400A77F095483515004DE4D0AC474E7EDC002256EFEEC85401598
          057D24A9EF302D1025E980569002CBA605381E341335E03E08814392FABB6A41
          AF04A29A8162500BEE4C0BA4C11C78059B24E7C09800B7E00D38239982166830
          2510011724A96F035D24672101AEFD16E0E9774CD27A9D24752F03CF600BCCF8
          2910D69D76831E90FDF2EC1CC449BE076F5E0894EA82317D31C7A0BE7C11EC92
          1C3E3BC6C00E1806477F11E085A6483E2A5C4F3E60139AE6EFF1026E9547B046
          529244BE02DC52FC61992799ED1649AB35E96ED749A65D3D68D4DFB8C33A4970
          908F4086E444E734F515BABB6D4DB9536D232AD102EA3483DC1D576095A44D5D
          95A05717E49D6635CD6EC21E4EE3A05F3733E924EFD7A5D41ECF7C3F4CD26759
          074C09704C9394934BB04132377E5C5AFDBE96F3CEB91C7C78F94C5C920C2E63
          021C16496BF2D4E40E5A312DF06B14040217F800C0E75321FC72D3C400000000
          49454E44AE426082}
      end
      item
        Background = clFuchsia
        Name = 'Print32x32_4'
        PngImage.Data = {
          89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7A
          F4000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000001034944415478DA63641860C038EA80C1EE80E3406C41A11D5780
          58975C07FC07E203407C904CCBED81D8019F3DC438A011881BC87400485F3DA9
          0E90016215287B3F102F00E285643A201E881380D811CABF03C44F0839E03510
          8B90692121F00388390939E03F85BE26141A2876E2720025F18E0B80CCC3480F
          A30E1875002107B43350A728CE25D70120F60328260728403123250EA056514C
          B1039603B104099647027106351D40D71038C000A982EB911C40AA23607A6066
          60AD9AB139E03E0324E1C000CC010E243AE000920360E00D108B1272007A6850
          C3015469903490E800981EAA39801C303C1C708061001BA5E839821C8091F249
          7100CDC1A80306DC01005F0C5E2155587C3D0000000049454E44AE426082}
      end
      item
        Background = clWindow
        Name = 'Print-WF_Enable_32x32_4'
        PngImage.Data = {
          89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7A
          F4000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA864000001184944415478DA63641860C038EA80C1ED807F0CC7812A2C28B2
          E13FC3150626065DF21CF01F0C0F00551D24D3727BA05E0720C6690F61073030
          3402553590E90090BE7AD21CF08F410628AA02E5ED071AB200C85F48A603E281
          7A13802C4728FF0E303A9E1072C06BA0A80859161276D00FA00338F13B0012EF
          E4FB9A5068A0450776075012EFB81D0032AF7ED401A30E20CD01FFC1D89E42AB
          4145712EB90E00C10740FE03322D57009AA300B6946C0750AB28A6D801FF80C5
          3229800958FC52D501740D0144155CCF809C204901083D10337054CDD82AA3FB
          E084830030073890E88003480E8079EE0D306A44F13B003D34A8E100AA344828
          8982A1D5221A740E18D04629668E20C71118299F7807D0018C3A60C01D000046
          6EA521893EEEB90000000049454E44AE426082}
      end
      item
        Background = clWindow
        Name = 'Print-WF_Red_32x32_4'
        PngImage.Data = {
          89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7A
          F4000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
          6FA8640000011B4944415478DAED963F0E823014C65FDDDC757370707577034E
          A217F00CE84DF424E0E4051C4D7470C30338895F059246B1D0570298B4C90BB4
          E9EBFBF57DFD27A8E3221C40AF019E44477458D80448894E03A2390B00CE292C
          46A70333B8075F5F68E25402E0B345A70D1340FA85460048FB048DB3BC1A6190
          1DEA7B26C012BE2BFC0679FD0C396E5500091A479C8035801E00186A0172DDD9
          B3AECAC6A71CA50064A1BB06408E173A000760049066E659C69647F19A0B20B7
          E49532E394A9C84CB001A8A1A3D81A00276464121C690F1A05683503CA151C2A
          3046108ACF7B8C5F5773D96574910B47692A007C438058012826778734632D80
          460E3640230F121B09FEEB45D43B804E1FA5253B8203F1B5F26B03B4511C40E7
          002F8D6EA521E4F0CA9F0000000049454E44AE426082}
      end>
    Left = 272
    Top = 168
    Bitmap = {}
  end
  object frxResultSet: TfrxUserDataSet
    UserName = 'RangeDataSet'
    Left = 392
    Top = 88
  end
  object frxStdDataSet: TfrxUserDataSet
    UserName = 'StdDataSet'
    Left = 392
    Top = 144
  end
  object frxRawDataDataSet: TfrxUserDataSet
    UserName = 'ODDataSet'
    Left = 392
    Top = 200
  end
  object frxPDFExport: TfrxPDFExport
    UseFileCache = True
    ShowProgress = True
    OverwritePrompt = False
    DataOnly = False
    OpenAfterExport = False
    PrintOptimized = False
    Outline = False
    Background = False
    HTMLTags = True
    Quality = 95
    Transparency = False
    Author = 'ELISA Report, SD BIOSENSOR'
    Subject = 'FastReport PDF export'
    ProtectionFlags = [ePrint, eModify, eCopy, eAnnot]
    HideToolbar = False
    HideMenubar = False
    HideWindowUI = False
    FitWindow = False
    CenterWindow = False
    PrintScaling = False
    PdfA = False
    DisableMultiCharGliph = False
    Left = 392
    Top = 264
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      '.Caption'
      'Button5.Caption'
      'ButtonPDF.Caption'
      'ButtonPrint.Caption'
      'CheckImgPrint.Caption'
      'ComboReportType.Items[0]'
      'ComboReportType.Items[1]'
      'ComboReportType.Items[2]'
      'Label1.Caption'
      'Label2.Caption'
      'Label3.Caption'
      'Label4.Caption'
      'LabelPrinterProperty.Caption'
      'LabelSubject.Caption'
      'MemoM2.Lines.Text'
      'MemoM3.Lines.Text')
    Left = 272
    Top = 232
  end
end
