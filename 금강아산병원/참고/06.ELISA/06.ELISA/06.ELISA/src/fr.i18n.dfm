object frI18n: TfrI18n
  Left = 0
  Top = 0
  Width = 313
  Height = 52
  TabOrder = 0
  object GridPanel1: TGridPanel
    Left = 0
    Top = 0
    Width = 313
    Height = 52
    Align = alClient
    BevelOuter = bvNone
    ColumnCollection = <
      item
        Value = 30.067973929340680000
      end
      item
        Value = 69.932026070659320000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = Label1
        Row = 0
      end
      item
        Column = 1
        Control = ComboLang
        Row = 0
      end
      item
        Column = 0
        Control = Label2
        Row = 1
      end
      item
        Column = 1
        Control = EditDateFmt
        Row = 1
      end>
    RowCollection = <
      item
        Value = 50.000000000000000000
      end
      item
        Value = 50.000000000000000000
      end>
    TabOrder = 0
    object Label1: TLabel
      AlignWithMargins = True
      Left = 40
      Top = 3
      Width = 51
      Height = 20
      Align = alRight
      Caption = 'Language:'
      Layout = tlCenter
      ExplicitHeight = 13
    end
    object ComboLang: TCultureBox
      AlignWithMargins = True
      Left = 97
      Top = 3
      Width = 213
      Height = 22
      Align = alClient
      Color = clWindow
      DisplayName = cnNativeLanguageName
      Flags = svcI18n.FlagImg
      Items.Cultures = (
        '*')
      Localizer = svcI18n.Localizer
      TabOrder = 0
      OnSelect = ComboLangSelect
    end
    object Label2: TLabel
      AlignWithMargins = True
      Left = 29
      Top = 29
      Width = 62
      Height = 20
      Align = alRight
      Caption = 'Date format:'
      Layout = tlCenter
      ExplicitHeight = 13
    end
    object EditDateFmt: TButtonedEdit
      AlignWithMargins = True
      Left = 97
      Top = 29
      Width = 213
      Height = 21
      Margins.Bottom = 2
      Align = alClient
      Images = svcImg.x16
      ReadOnly = True
      RightButton.DropDownMenu = PopupDateFmt
      RightButton.ImageIndex = 0
      RightButton.Visible = True
      TabOrder = 1
    end
  end
  object PopupDateFmt: TPopupMenu
    Alignment = paCenter
    AutoHotkeys = maManual
    AutoLineReduction = maManual
    Left = 136
    Top = 3
    object MenuItemYYYYMMDD: TMenuItem
      Caption = 'YYYY-MM-DD'
      GroupIndex = 1
      RadioItem = True
      OnClick = MenuItemDateFmtClick
    end
    object MenuItemMMDDYYYY: TMenuItem
      Tag = 1
      Caption = 'MM-DD-YYYY'
      GroupIndex = 1
      RadioItem = True
      OnClick = MenuItemDateFmtClick
    end
    object MenuItemDDMMYYYY: TMenuItem
      Tag = 2
      Caption = 'DD-MM-YYYY'
      GroupIndex = 1
      RadioItem = True
      OnClick = MenuItemDateFmtClick
    end
  end
end
