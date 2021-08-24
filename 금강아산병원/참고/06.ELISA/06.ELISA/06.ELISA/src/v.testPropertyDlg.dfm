object vTestPropertyDlg: TvTestPropertyDlg
  Left = 0
  Top = 0
  ActiveControl = EditTestNum
  BiDiMode = bdLeftToRight
  BorderStyle = bsDialog
  Caption = 'Test property'
  ClientHeight = 267
  ClientWidth = 432
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
    432
    267)
  PixelsPerInch = 96
  TextHeight = 16
  object ButtonCreate: TButton
    Left = 247
    Top = 232
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '&OK'
    Enabled = False
    ModalResult = 1
    TabOrder = 0
    OnClick = ButtonCreateClick
  end
  object Button2: TButton
    Left = 328
    Top = 232
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '&Cancel'
    ModalResult = 2
    TabOrder = 1
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 432
    Height = 224
    Align = alTop
    BevelOuter = bvNone
    Color = clWhite
    ParentBackground = False
    TabOrder = 2
    object Label1: TLabel
      Left = 67
      Top = 93
      Width = 60
      Height = 16
      Alignment = taRightJustify
      Caption = 'Test Date:'
    end
    object Image1: TImage
      Left = 8
      Top = 8
      Width = 64
      Height = 64
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000400000
        00400806000000AA6971DE000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000001
        5A4944415478DAEDDAAD4E03411885E16D8240222BCB1D6090247007482EA10E
        1048022478300483C061B8803A1088CA3A24700795080467D21165B2FC244CF6
        4C76DE939C64C566FA7D4FB7553B682ACFC03D803B00B807700700F700EE00E0
        1EC01D00DC03B8D315C0817A91F1BC7775A69EA993D201722FBF9C0F75336214
        09902E1FBEB9698673D7D48D787DAC9E9708D0F6CDBFAAEB19CEDE561FE275F8
        199C9606902E1F1ED5955A00D2E50FD57D75540340DBF297EA4B0D00DF2DDFD4
        00F0D3F2BD074897BF57AF927BEED4611F01F6E2727F4DEF00C2079F00B0C8AD
        FAF6CBFDF3E6EB7F43AF0076D4C70CCB01000000000000000000000000000000
        000000000000000000000000000000000000E5018CD5EB781DDED59B7704B0FC
        9A9C1560A43EABAB1D2DDE962DF5C90510B2DB2C9E8261C78B87A7ED48BDF9CF
        21BC2CED1EC01D00DC03B803807B007700700FE00E00EE01DCF904BCE8B841DC
        15A7520000000049454E44AE426082}
    end
    object Label5: TLabel
      Left = 80
      Top = 15
      Width = 334
      Height = 49
      AutoSize = False
      Caption = 
        'Type the information of the test which contains Test Date, Test ' +
        'Number, Kit batch Number and an Operator.'
      WordWrap = True
    end
    object EditTestNum: TLabeledEdit
      Left = 152
      Top = 119
      Width = 249
      Height = 24
      Alignment = taRightJustify
      EditLabel.Width = 79
      EditLabel.Height = 16
      EditLabel.Caption = '&Test Number:'
      LabelPosition = lpLeft
      LabelSpacing = 24
      TabOrder = 0
      OnChange = OnEditChange
      OnKeyDown = OnEditKeyDown
    end
    object EditOperator: TLabeledEdit
      Left = 152
      Top = 179
      Width = 249
      Height = 24
      EditLabel.Width = 56
      EditLabel.Height = 16
      EditLabel.Caption = 'O&perator:'
      LabelPosition = lpLeft
      LabelSpacing = 24
      TabOrder = 1
      OnChange = OnEditChange
      OnKeyDown = OnEditKeyDown
    end
    object EditBatchNum: TLabeledEdit
      Left = 152
      Top = 149
      Width = 249
      Height = 24
      Alignment = taRightJustify
      EditLabel.Width = 103
      EditLabel.Height = 16
      EditLabel.Caption = '&Kit Batch Number:'
      LabelPosition = lpLeft
      LabelSpacing = 24
      TabOrder = 2
      OnChange = OnEditChange
      OnKeyDown = OnEditKeyDown
    end
    object DatePicker: TCalendarPicker
      Left = 152
      Top = 88
      Width = 249
      Height = 25
      CalendarHeaderInfo.DaysOfWeekFont.Charset = DEFAULT_CHARSET
      CalendarHeaderInfo.DaysOfWeekFont.Color = clWindowText
      CalendarHeaderInfo.DaysOfWeekFont.Height = -13
      CalendarHeaderInfo.DaysOfWeekFont.Name = 'Segoe UI'
      CalendarHeaderInfo.DaysOfWeekFont.Style = []
      CalendarHeaderInfo.Font.Charset = DEFAULT_CHARSET
      CalendarHeaderInfo.Font.Color = clWindowText
      CalendarHeaderInfo.Font.Height = -20
      CalendarHeaderInfo.Font.Name = 'Segoe UI'
      CalendarHeaderInfo.Font.Style = []
      Color = clWindow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGray
      Font.Height = -13
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      TextHint = 'select a date'
    end
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      '.Caption'
      'Button2.Caption'
      'ButtonCreate.Caption'
      'DatePicker.TextHint'
      'EditBatchNum.EditLabel.Caption'
      'EditOperator.EditLabel.Caption'
      'EditTestNum.EditLabel.Caption'
      'Label1.Caption'
      'Label5.Caption')
    Left = 208
    Top = 136
  end
end
