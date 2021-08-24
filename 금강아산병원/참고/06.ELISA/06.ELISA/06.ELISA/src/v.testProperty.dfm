object vTestProperty: TvTestProperty
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  BorderStyle = bsNone
  Caption = 'vTestProperty'
  ClientHeight = 422
  ClientWidth = 235
  Color = 14671839
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  StyleElements = [seFont, seBorder]
  OnCreate = FormCreate
  DesignSize = (
    235
    422)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    AlignWithMargins = True
    Left = 8
    Top = 75
    Width = 227
    Height = 13
    Margins.Left = 8
    Margins.Right = 0
    Margins.Bottom = 0
    Align = alTop
    Caption = 'Test Date:'
    ExplicitWidth = 51
  end
  object PanelTitle: TPanel
    AlignWithMargins = True
    Left = 8
    Top = 3
    Width = 227
    Height = 25
    Margins.Left = 8
    Margins.Right = 0
    Align = alTop
    Alignment = taLeftJustify
    BevelOuter = bvNone
    Caption = 'Test Property'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBtnShadow
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    DesignSize = (
      227
      25)
    object ButtonClose: TPngSpeedButton
      Left = 204
      Top = 2
      Width = 23
      Height = 22
      Anchors = [akTop, akRight]
      Flat = True
      OnClick = ButtonCloseClick
      PngImage.Data = {
        89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
        61000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
        6FA864000000864944415478DA6364A0103052D380E9403C11886F10D02301C4
        DD401C8B6E8006102F07E2483C8680346F06E242203E82CD0BF80CC1D08C2B0C
        B019825533BE404436E4032ECDF80C4036E40F2ECD840C90801AC00075C90B52
        0C8069CE84F2A7E3328491806658206AE032849108CD0CF80C612452334E43C8
        49CA2043EAA18650373391050049232A11ECE012B00000000049454E44AE4260
        82}
      ExplicitLeft = 272
    end
  end
  object EditTestNum: TLabeledEdit
    Left = 8
    Top = 138
    Width = 227
    Height = 21
    Margins.Right = 0
    Anchors = [akLeft, akTop, akRight]
    EditLabel.Width = 65
    EditLabel.Height = 13
    EditLabel.BiDiMode = bdLeftToRight
    EditLabel.Caption = '&Test Number:'
    EditLabel.ParentBiDiMode = False
    TabOrder = 1
    OnChange = DatePickerChange
    OnKeyDown = EditKeyDown
  end
  object EditBatchNum: TLabeledEdit
    Left = 8
    Top = 181
    Width = 227
    Height = 21
    Margins.Right = 0
    Anchors = [akLeft, akTop, akRight]
    EditLabel.Width = 86
    EditLabel.Height = 13
    EditLabel.Caption = '&Kit Batch Number:'
    TabOrder = 2
    OnChange = DatePickerChange
    OnKeyDown = EditKeyDown
  end
  object EditOperator: TLabeledEdit
    Left = 8
    Top = 227
    Width = 227
    Height = 21
    Margins.Right = 0
    Anchors = [akLeft, akTop, akRight]
    EditLabel.Width = 48
    EditLabel.Height = 13
    EditLabel.Caption = 'O&perator:'
    TabOrder = 3
    OnChange = DatePickerChange
    OnKeyDown = EditKeyDown
  end
  object TPanel
    Left = 0
    Top = 31
    Width = 235
    Height = 41
    Margins.Right = 0
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 4
    object ButtonSave: TButton
      Left = 8
      Top = 3
      Width = 121
      Height = 25
      Caption = 'Save to change'
      Enabled = False
      TabOrder = 0
      OnClick = ButtonSaveClick
    end
  end
  object DatePicker: TCalendarPicker
    Left = 8
    Top = 91
    Width = 227
    Height = 25
    Margins.Right = 0
    Anchors = [akLeft, akTop, akRight]
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
    TabOrder = 5
    TextHint = 'select a date'
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      'ButtonSave.Caption'
      'EditBatchNum.EditLabel.Caption'
      'EditOperator.EditLabel.Caption'
      'EditTestNum.EditLabel.Caption'
      'Label1.Caption'
      'PanelTitle.Caption')
    Left = 104
    Top = 200
  end
end
