object vI18nDlg: TvI18nDlg
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  BorderStyle = bsDialog
  Caption = 'ELISA Report'
  ClientHeight = 393
  ClientWidth = 597
  Color = clBtnFace
  DoubleBuffered = True
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object PanelContents: TRelativePanel
    Left = 0
    Top = 0
    Width = 597
    Height = 393
    ControlCollection = <
      item
        Control = LabelI18nDesc
        AlignBottomWithPanel = False
        AlignHorizontalCenterWithPanel = False
        AlignLeftWithPanel = False
        AlignRightWithPanel = False
        AlignTopWithPanel = False
        AlignVerticalCenterWithPanel = False
        Below = PanelLogo
      end
      item
        Control = PanelLogo
        AlignBottomWithPanel = False
        AlignHorizontalCenterWithPanel = True
        AlignLeftWithPanel = False
        AlignRightWithPanel = False
        AlignTopWithPanel = False
        AlignVerticalCenterWithPanel = False
      end
      item
        Control = PanelI18n
        AlignBottomWith = LabelI18nDesc
        AlignBottomWithPanel = True
        AlignHorizontalCenterWithPanel = False
        AlignLeftWithPanel = False
        AlignRightWith = LabelI18nDesc
        AlignRightWithPanel = True
        AlignTopWithPanel = False
        AlignVerticalCenterWithPanel = False
        Below = LabelI18nDesc
      end
      item
        Control = PanelButton
        AlignBottomWithPanel = False
        AlignHorizontalCenterWithPanel = False
        AlignLeftWithPanel = False
        AlignRightWithPanel = False
        AlignTopWithPanel = False
        AlignVerticalCenterWithPanel = False
        Below = PanelI18n
      end>
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    OnResize = PanelContentsResize
    DesignSize = (
      597
      393)
    object LabelI18nDesc: TLabel
      Left = 262
      Top = 159
      Width = 201
      Height = 16
      Anchors = []
      Caption = 'Please select the following options:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object PanelLogo: TPanel
      Left = 40
      Top = 64
      Width = 516
      Height = 89
      Anchors = [akLeft, akRight]
      TabOrder = 2
      object LabelModuleName: TLabel
        Left = 1
        Top = 1
        Width = 514
        Height = 87
        Align = alClient
        Alignment = taCenter
        Caption = 'ELISA Report'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -48
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
        Layout = tlCenter
        WordWrap = True
        ExplicitWidth = 322
        ExplicitHeight = 58
      end
    end
    object PanelI18n: TPanel
      Left = 150
      Top = 181
      Width = 313
      Height = 50
      Anchors = []
      TabOrder = 0
    end
    object PanelButton: TPanel
      Left = 313
      Top = 237
      Width = 150
      Height = 33
      Anchors = []
      TabOrder = 1
      object ButtonOk: TButton
        Tag = 1
        AlignWithMargins = True
        Left = 71
        Top = 4
        Width = 75
        Height = 25
        Align = alRight
        Caption = '&OK'
        ModalResult = 1
        TabOrder = 0
        OnClick = ButtonOkClick
      end
    end
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      'ButtonOk.Caption'
      'LabelI18nDesc.Caption')
    Left = 80
    Top = 320
  end
  object PopupDateFmt: TPopupMenu
    Alignment = paCenter
    AutoHotkeys = maManual
    AutoLineReduction = maManual
    Left = 136
    Top = 320
    object MenuItemYYYYMMDD: TMenuItem
      Caption = 'YYYY-MM-DD'
      GroupIndex = 1
      RadioItem = True
    end
    object MenuItemMMDDYYYY: TMenuItem
      Tag = 1
      Caption = 'MM-DD-YYYY'
      GroupIndex = 1
      RadioItem = True
    end
    object MenuItemDDMMYYYY: TMenuItem
      Tag = 2
      Caption = 'DD-MM-YYYY'
      GroupIndex = 1
      RadioItem = True
    end
  end
end
