object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 291
  Top = 196
  Height = 245
  Width = 637
  object qryC1: TADOQuery
    Parameters = <>
    Left = 18
    Top = 76
  end
  object qryC2: TADOQuery
    Parameters = <>
    Left = 20
    Top = 130
  end
  object qryC3: TADOQuery
    Parameters = <>
    Left = 60
    Top = 68
  end
  object qrySUp: TADOQuery
    Parameters = <>
    Left = 264
    Top = 48
  end
  object qryUp1: TADOQuery
    Parameters = <>
    Left = 20
    Top = 12
  end
  object qryUpOne: TADOQuery
    Parameters = <>
    Left = 60
    Top = 12
  end
  object qrySUp1: TADOQuery
    Parameters = <>
    Left = 264
    Top = 104
  end
  object qryV: TADOQuery
    Parameters = <>
    Left = 68
    Top = 130
  end
  object spUp: TADOStoredProc
    ProcedureName = 'PR_CPL_INSERT_CPL0891'
    Parameters = <
      item
        Name = 'I_USER_ID'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'I_JANGBI_CODE'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'I_SPECIMEN_SER'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'I_JANGBI_OUT_CODE'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'I_CPL_RESULT'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'I_RESULT_DATE'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'I_RESULT_SEQ'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end>
    Left = 376
    Top = 128
  end
  object conHosp: TADOConnection
    ConnectionString = 
      'Provider=msdaora.1;Data Source=schora2;User ID=medi;Password=med' +
      'i;Persist Security Info=True'
    Provider = 'msdaora.1'
    Left = 496
    Top = 136
  end
  object qrySOrder: TADOQuery
    Parameters = <
      item
        Name = 'in_spcid'
        Size = -1
        Value = Null
      end
      item
        Name = 'in_examcode'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'SELECT PRSNRSLT AS Result  '
      '      ,PRSNMRNO AS PatNo   '
      '      ,PRSNLBNO AS LabNo   '
      '      ,PRSNORNO AS OrdNo   '
      '      ,PRSNORSQ AS OrdSeq  '
      '      ,PRSNVSDT AS AcptDt  '
      '      ,PRSNCODE AS ExamCode'
      '      ,PRSNSUBC AS ItemCode'
      'FROM MPSDTA.PRSNUMBM'
      'where PRSNLBNO = :in_spcid     '
      '  And PRSNCODE = :in_examcode')
    Left = 408
    Top = 56
  end
end
