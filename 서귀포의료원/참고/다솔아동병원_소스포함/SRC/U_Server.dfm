object F_Server: TF_Server
  Left = 192
  Top = 124
  Width = 1305
  Height = 675
  Caption = 'F_Server'
  Color = clBtnFace
  Font.Charset = HANGEUL_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #44404#47548#52404
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 12
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1289
    Height = 41
    Align = alTop
    TabOrder = 0
    object Button1: TButton
      Left = 80
      Top = 8
      Width = 75
      Height = 25
      Caption = #53244#47532#51312#54924
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 160
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Excute'
      TabOrder = 1
      OnClick = Button2Click
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 618
    Width = 1289
    Height = 19
    Panels = <>
  end
  object Panel2: TPanel
    Left = 0
    Top = 41
    Width = 1289
    Height = 577
    Align = alClient
    Caption = 'Panel1'
    TabOrder = 2
    object Panel3: TPanel
      Left = 1
      Top = 201
      Width = 1287
      Height = 375
      Align = alClient
      TabOrder = 0
      object DBGrid1: TDBGrid
        Left = 1
        Top = 1
        Width = 1285
        Height = 373
        Align = alClient
        DataSource = DataSource1
        ImeName = 'Microsoft IME 2010'
        Options = [dgEditing, dgAlwaysShowEditor, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
        TabOrder = 0
        TitleFont.Charset = HANGEUL_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = #44404#47548#52404
        TitleFont.Style = []
      end
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 1287
      Height = 200
      Align = alTop
      TabOrder = 1
      object mmSQL: TMemo
        Left = 1
        Top = 1
        Width = 1285
        Height = 198
        Align = alClient
        ImeName = 'Microsoft IME 2010'
        TabOrder = 0
      end
    end
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 784
    Top = 16
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 729
    Top = 274
  end
end
