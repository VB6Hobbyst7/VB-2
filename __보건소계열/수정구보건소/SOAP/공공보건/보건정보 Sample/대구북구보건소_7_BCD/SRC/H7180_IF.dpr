program H7180_IF;

uses
  FastMM4,
  SyncObjs,
  Forms,
  U_Main in 'U_Main.pas' {F_Main},
  GlobalVar in 'GlobalVar.pas',
  SetDataBase in 'SetDataBase.pas',
  U_DM in 'U_DM.pas' {DM: TDataModule},
  U_CodeInfo in 'U_CodeInfo.pas',
  U_IFClass in 'U_IFClass.pas',
  U_Server in 'U_Server.pas',
  U_CommSet in 'U_CommSet.pas' {F_CommSet},
  U_TEST in 'U_TEST.pas' {F_Test},
  U_CODE_SET in 'U_CODE_SET.pas' {F_CodeSet};

{$R *.res}

begin
  RegisterExpectedMemoryLeak(TCriticalSection ,1);
  Application.Initialize;
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TF_Main, F_Main);
  Application.CreateForm(TF_Test, F_Test);
  Application.CreateForm(TF_CodeSet, F_CodeSet);
  Application.Run;
end.
