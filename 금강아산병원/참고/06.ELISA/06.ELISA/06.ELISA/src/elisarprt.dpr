program elisarprt;

uses
  Winapi.Windows,
  mUtils.Windows,
  Vcl.Forms,
  v.main in 'v.main.pas' {vMain},
  svc in 'svc.pas',
  svc.option in 'svc.option.pas' {svcOption: TDataModule},
  v.i18nDlg in 'v.i18nDlg.pas' {vI18nDlg},
  svc.i18n in 'svc.i18n.pas' {svcI18n: TDataModule},
  m.i18n in 'm.i18n.pas',
  svc.img in 'svc.img.pas' {svcImg: TDataModule},
  v.testPropertyDlg in 'v.testPropertyDlg.pas' {vTestPropertyDlg},
  v.eula in 'v.eula.pas' {vEula},
  v.rawdata in 'v.rawdata.pas' {vRawdata},
  m.test in 'm.test.pas',
  svc.test in 'svc.test.pas',
  Vcl.Themes,
  Vcl.Styles,
  v.viewNames in 'v.viewNames.pas' {vViewNames},
  m.rawdata in 'm.rawdata.pas',
  v.rawdataFmt in 'v.rawdataFmt.pas' {vRawdataFmt},
  v.testProperty in 'v.testProperty.pas' {vTestProperty},
  ribbon.main in 'ribbon\ribbon.main.pas',
  m.calc.std in 'm.calc.std.pas',
  v.calc.std in 'v.calc.std.pas' {vCalcStd},
  v.calc.mtrl in 'v.calc.mtrl.pas' {vCalcMtrl},
  v.report in 'v.report.pas' {vReport},
  v.option in 'v.option.pas' {vOption},
  fr.i18n in 'fr.i18n.pas' {frI18n: TFrame},
  v.i18n in 'v.i18n.pas' {vI18n},
  v.progressDlg in 'v.progressDlg.pas' {vProgressDlg},
  v.about in 'v.about.pas' {vAbout},
  CurveFit in 'curvefit\Source\CurveFit.pas',
  m.curvefit in 'curvefit\source\m.curvefit.pas',
  m.calc.mtrl in 'm.calc.mtrl.pas',
  m.block in 'm.block.pas';

{$R *.res}

var
  Mutex: THandle = 0;
begin
  Mutex := CreateMutex(nil, True, 'elisarprt_mutex');
  try
    if (Mutex = 0 ) or (GetLastError <> 0) then
      TerminateProcess('elisarprt.exe');

    Application.Initialize;
    Application.MainFormOnTaskbar := True;
    TStyleManager.TrySetStyle('Light');
  Application.CreateForm(TsvcI18n, svcI18n);
  Application.CreateForm(TsvcOption, svcOption);
  Application.CreateForm(TsvcImg, svcImg);
  Application.CreateForm(TvMain, vMain);
  Application.CreateForm(TvI18n, vI18n);
  Application.Run;
  finally
    if Mutex <> 0 then
      CloseHandle(Mutex);
  end;
end.
