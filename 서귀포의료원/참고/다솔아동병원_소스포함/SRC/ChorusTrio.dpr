program ChorusTrio;

uses
  FastMM4,
  controls,
  SyncObjs,
  Forms,
  U_Main in 'U_Main.pas' {F_Main},
  GlobalVar in 'GlobalVar.pas',
  SetDataBase in 'SetDataBase.pas',
  U_DM in 'U_DM.pas' {DM: TDataModule},
  U_CodeInfo in 'U_CodeInfo.pas',
  U_IFClass in 'U_IFClass.pas',
  U_CodeM in 'U_CodeM.pas' {F_CodeM},
  U_Server in 'U_Server.pas' {F_Server};

{$R *.res}

begin
  RegisterExpectedMemoryLeak(TCriticalSection ,1);

  Application.Initialize;
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TF_Main, F_Main);
  Application.Run;

end.
