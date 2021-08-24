object svcI18n: TsvcI18n
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 300
  Width = 409
  object Localizer: TLocalizer
    URI = 'elisarprt.i18n'
    Options = [loAutoSetectLanguage, loAdjustApplicationBiDiMode, loAdjustFormatSettings, loUseNativeDigits, loUseNativeCalendar]
    Left = 64
    Top = 40
  end
  object FlagImg: TFlagImageList
    Left = 64
    Top = 152
  end
  object Reg: TRzRegIniFile
    PathType = ptRegistry
    Left = 192
    Top = 40
  end
  object Translator: TTranslator
    Localizer = Localizer
    Translatables.Literals = (
      '537C66B24EF5C83B7382CDC3F34885F2'
      '7CBB885AA1164B390A0BC050A64E1812'
      '03727AC48595A24DAED975559C944A44')
    Left = 64
    Top = 96
  end
end
