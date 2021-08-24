unit v.main;

interface

uses
  mvw.vForm, Spring.Collections,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, RzForms, Vcl.ComCtrls, RzTabs, UIRibbonCommands, UIRibbon, Vcl.PlatformDefaultStyleActnCtrls,
  System.Actions, Vcl.ActnList, Vcl.ActnMan, i18nCore, i18nLocalizer, Vcl.ExtCtrls, RzShellDialogs, RzStatus, RzPanel,
  System.ImageList, Vcl.ImgList, PngImageList, Vcl.StdCtrls, HTMLabel, Vcl.ExtDlgs;

type
  TvMain = class(TvForm)
    RzFormState: TRzFormState;
    Ribbon: TUIRibbon;
    Translator: TTranslator;
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    ActionManager: TActionManager;
    ActionOpen: TAction;
    ActionSave: TAction;
    ActionSaveAs: TAction;
    ActionSdb: TAction;
    ActionOption: TAction;
    ActionAbout: TAction;
    ActionExit: TAction;
    ActionNewTest: TAction;
    ActionRecentTest: TAction;
    ActionTabHome: TAction;
    ActionGroupData: TAction;
    ActionPaste: TAction;
    ActionManualDataEntry: TAction;
    ActionGroupDataFmt: TAction;
    ActionFmtDefault: TAction;
    ActionFmtManual: TAction;
    ActionFmtLoad: TAction;
    ActionFmtSave: TAction;
    ActionFmtViewNames: TAction;
    ActionGroupAnalyze: TAction;
    ActionCalc: TAction;
    ActionPrint: TAction;
    ActionExportResult: TAction;
    ActionExportData: TAction;
    ActionFmt3Item: TAction;
    ActionFmt2Item: TAction;
    ActionTestProperty: TAction;
    ActionTabResults: TAction;
    ActionTabTest: TAction;
    RzStatusBar1: TRzStatusBar;
    StatusTestDate: TRzFieldStatus;
    StatusTestNum: TRzFieldStatus;
    StatusKitBatch: TRzFieldStatus;
    StatusOperator: TRzFieldStatus;
    TimerInit: TTimer;
    StatusVersion: TRzFieldStatus;
    ActionClearTest: TAction;
    PageControl: TRzPageControl;
    TabSheetRawData: TRzTabSheet;
    TabSheetStdResult: TRzTabSheet;
    SplitterFmt: TSplitter;
    PanelFmt: TPanel;
    PngImageList1: TPngImageList;
    ActionResultToClipboard: TAction;
    ActionResultToCsv: TAction;
    ActionDataToClipboard: TAction;
    ActionDataToCsv: TAction;
    PanelCalcStd: TPanel;
    PanelCalcMtrl: TPanel;
    PanelNoProp: TPanel;
    LabelProperty: THTMLabel;
    ShapeProperty: TShape;
    ActionManualTest: TAction;
    PanelProperty: TPanel;
    ShapeFmt: TShape;
    SplitterProperty: TSplitter;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

    procedure TimerInitTimer(Sender: TObject);

    procedure ActionSdbExecute(Sender: TObject);
    procedure ActionAboutExecute(Sender: TObject);
    procedure ActionOptionExecute(Sender: TObject);
    procedure ActionExitExecute(Sender: TObject);

    procedure ActionNewTestExecute(Sender: TObject);
    procedure ActionClearTestExecute(Sender: TObject);
    procedure ActionFmtExecute(Sender: TObject);
    procedure ActionFmtManualExecute(Sender: TObject);
    procedure ActionPasteExecute(Sender: TObject);
    procedure ActionFmtViewNamesExecute(Sender: TObject);
    procedure ActionGroupAnalyzeExecute(Sender: TObject);
    procedure ActionTestPropertyExecute(Sender: TObject);
    procedure ActionCalcExecute(Sender: TObject);

    procedure ActionPrintExecute(Sender: TObject);
    procedure ActionSaveExecute(Sender: TObject);
    procedure ActionSaveAsExecute(Sender: TObject);
    procedure ActionOpenExecute(Sender: TObject);
    procedure ActionFmtSaveExecute(Sender: TObject);
    procedure ActionFmtLoadExecute(Sender: TObject);
    procedure ActionExportResultExecute(Sender: TObject);
    procedure ActionResultToClipboardExecute(Sender: TObject);
    procedure ActionResultToCsvExecute(Sender: TObject);
    procedure ActionDataToCsvExecute(Sender: TObject);
    procedure ActionDataToClipboardExecute(Sender: TObject);
    procedure ActionManualTestExecute(Sender: TObject);

    procedure ActionSplitButton(Sender: TObject);

    procedure ActionTabTestExecute(Sender: TObject);
    procedure TranslatorAfterTranslate(Sender: TObject);
    procedure LabelPropertyAnchorClick(Sender: TObject; Anchor: string);
  private
    FDataFileName: String;
    procedure ShowNoProperyPanel;
    procedure EnableFmtUIs(const AEnabled: Boolean);
    procedure EnableCalcUIs;
    procedure UpdateStatusBar;
    function QueryCanDeleteUnsavedData: Boolean;
    procedure ClearTest(const AClearWithProperties: Boolean = True);
  public
  end;

var
  vMain: TvMain;

implementation

{$R *.dfm}

uses
  m.test, m.rawdata,
  svc,
  v.TestProperty, v.TestPropertyDlg, v.i18nDlg, v.eula, v.rawdata, v.ViewNames, v.rawdataFmt, v.calc.std,
  v.calc.mtrl, v.option, v.report, v.about,

  mUtils.Windows, CodeSiteLogging, mCodeSiteHelper, GraphUtil, UIRibbonUtils, System.DateUtils, mDateTimeHelper,
  System.Math, System.StrUtils, System.UITypes, Vcl.Themes, System.IOUtils, mIOUtils
  ;

function SetThreadPreferredUILanguages(dwFlags: DWORD; pwszLanguagesBuffer: PWideChar; pulNumLanguages: PULONG): BOOL; stdcall; external kernel32 name 'SetThreadPreferredUILanguages';

procedure TvMain.ActionAboutExecute(Sender: TObject);
begin
  TvAbout.Open
end;

procedure TvMain.ActionCalcExecute(Sender: TObject);
begin
  if dataContainer.HasStd then
  begin
    FDataFileName := '';
    vCalcStd.Initialize;
    vCalcMtrl.Initailize;
    EnableCalcUIs;
    PageControl.ActivePage := TabSheetStdResult;
  end;
end;

procedure TvMain.ActionClearTestExecute(Sender: TObject);
begin
  if not QueryCanDeleteUnsavedData then
    Exit;

  ClearTest(False);
end;

procedure TvMain.ActionDataToClipboardExecute(Sender: TObject);
begin
  vCalcStd.ExportToClipboard;
end;

procedure TvMain.ActionDataToCsvExecute(Sender: TObject);
begin
  SaveDialog.InitialDir := option.ExportDataPath;
  SaveDialog.DefaultExt := option.SCsvExt;
  SaveDialog.Filter := option.SCsvFilter;
  SaveDialog.FileName := '';
  if SaveDialog.Execute then
  begin
    option.ExportDataPath := TDirectory.GetParent(SaveDialog.FileName);
    vCalcStd.ExportToFile(SaveDialog.FileName);
  end;
end;

procedure TvMain.ActionExitExecute(Sender: TObject);
begin
  Close;
end;

procedure TvMain.ActionExportResultExecute(Sender: TObject);
begin
  //
end;

procedure TvMain.ActionFmtExecute(Sender: TObject);
begin
  option.MaterialCnt := TAction(Sender).Tag;
  case option.MaterialCnt of
    2: dataContainer.AssignDefault(cmNil_Antigen);
    3: dataContainer.AssignDefault(cmNil_Antigen_Mitogen);
  end;
  EnableCalcUIs;
  PageControl.ActivePage := TabSheetRawData;
  Application.ProcessMessages;;
  vRawdata.BuildGridBmp;
end;

procedure TvMain.ActionFmtLoadExecute(Sender: TObject);
begin
  PageControl.ActivePage := TabSheetRawData;
  OpenDialog.InitialDir := option.LoadFmtPath;
  OpenDialog.DefaultExt := option.STBFmtExt;
  OpenDialog.Filter := option.SOpenDlgTBFmtFilter;
  OpenDialog.FileName := '';
  if OpenDialog.Execute then
  begin
    option.LoadFmtPath := TDirectory.GetParent(OpenDialog.FileName);
    dataContainer.LoadFromFile(OpenDialog.FileName);
    UpdateStatusBar;
    EnableFmtUIs(True);
    EnableCalcUIs;
  end;
end;

procedure TvMain.EnableCalcUIs;
begin
  ActionCalc.Enabled := dataContainer.HasStd and dataContainer.CanCalc;
  TabSheetStdResult.TabVisible := stdCalc.HasResult;
  ActionSave.Enabled := stdCalc.HasResult;
  ActionSaveAs.Enabled := stdCalc.HasResult;
  ActionPrint.Enabled := stdCalc.HasResult;
  ActionExportResult.Enabled := stdCalc.HasResult;
  ActionExportData.Enabled := stdCalc.HasResult;
end;

procedure TvMain.EnableFmtUIs(const AEnabled: Boolean);
begin
  ActionSave.Enabled := dataContainer.HasData;
  ActionSaveAs.Enabled := dataContainer.HasData;
  ActionGroupDataFmt.Enabled := AEnabled;
  ActionFmtDefault.Enabled := AEnabled;
  ActionFmtManual.Enabled := AEnabled;
  ActionFmt3Item.Enabled := AEnabled;
  ActionFmt2Item.Enabled := AEnabled;
  ActionFmtSave.Enabled := AEnabled;
  ActionFmtViewNames.Enabled := AEnabled;
end;

procedure TvMain.ActionFmtManualExecute(Sender: TObject);
var
  LAction: TAction absolute Sender;
begin
  dataFmter.Enabled := LAction.Checked;

  PanelFmt.Visible := LAction.Checked;
  ShapeFmt.Visible := LAction.Checked;
  SplitterFmt.Visible := LAction.Checked;

  ShapeFmt.Left := SplitterFmt.Left - ShapeFmt.Width -1;
  SplitterFmt.Left := PanelFmt.Left - SplitterFmt.Width -1;
  EnableCalcUIs;
  PageControl.ActivePage := TabSheetRawData;
end;

procedure TvMain.ActionFmtSaveExecute(Sender: TObject);
begin
  PageControl.ActivePage := TabSheetRawData;

  SaveDialog.InitialDir := option.SaveFmtPath;
  SaveDialog.DefaultExt := '.tbf';
  if SaveDialog.Execute then
  begin
    dataContainer.SaveToFile(SaveDialog.FileName);
    option.SaveFmtPath := TDirectory.GetParent(SaveDialog.FileName);
    EnableFmtUIs(True);
  end;
end;

procedure TvMain.ActionFmtViewNamesExecute(Sender: TObject);
begin
  TvViewNames.Open;
end;

procedure TvMain.ActionGroupAnalyzeExecute(Sender: TObject);
var
  LAction: TAction absolute Sender;
begin
  CodeSite.Send(LAction.Caption);
end;

procedure TvMain.ActionManualTestExecute(Sender: TObject);
begin
  if not QueryCanDeleteUnsavedData then
    Exit;

  ClearTest(False);

  vRawdata.WriteInTest;
  EnableFmtUIs(True);
  PageControl.ActivePage := TabSheetRawData;
end;

procedure TvMain.ActionNewTestExecute(Sender: TObject);
begin
  if not QueryCanDeleteUnsavedData then
    Exit;

  ClearTest;
  TvTestPropertyDlg.Open;
end;

procedure TvMain.ActionOpenExecute(Sender: TObject);
begin
  OpenDialog.InitialDir := option.OpenDataPath;
  OpenDialog.DefaultExt := option.SDataExt;
  OpenDialog.Filter := option.SOpenDlgDataFilter;
  OpenDialog.FileName := '';
  if OpenDialog.Execute then
  begin
    option.OpenDataPath := TDirectory.GetParent(OpenDialog.FileName);
    dataContainer.LoadFromFile(OpenDialog.FileName);
    ShowNoProperyPanel;
    UpdateStatusBar;
    EnableFmtUIs(True);
    EnableCalcUIs;
    ActionCalc.Execute;
    PageControl.ActivePage := TabSheetRawData;
    Application.ProcessMessages;
    vRawdata.BuildGridBmp;
  end;
end;

procedure TvMain.ActionOptionExecute(Sender: TObject);
begin
  TvOption.Open;
end;

procedure TvMain.ActionPasteExecute(Sender: TObject);
var
  LSuccess: Boolean;
begin
  if not QueryCanDeleteUnsavedData then
    Exit;

  ClearTest(False);

  LSuccess := vRawdata.Paste;
  EnableFmtUIs(LSuccess);
  if LSuccess then
    PageControl.ActivePage := TabSheetRawData;
end;

procedure TvMain.ActionPrintExecute(Sender: TObject);
begin
  if stdCalc.HasResult then
    TvReport.Open;
end;

procedure TvMain.ActionResultToClipboardExecute(Sender: TObject);
begin
  vCalcMtrl.ExportToClilpboard;
end;

procedure TvMain.ActionResultToCsvExecute(Sender: TObject);
begin
  SaveDialog.InitialDir := option.ExportDataPath;
  SaveDialog.DefaultExt := option.SCsvExt;
  SaveDialog.Filter := option.SCsvFilter;
  SaveDialog.FileName := '';
  if SaveDialog.Execute then
  begin
    option.ExportDataPath := TDirectory.GetParent(SaveDialog.FileName);
    vCalcMtrl.ExportToCsvFile(SaveDialog.FileName);
  end;
end;

procedure TvMain.ActionSaveAsExecute(Sender: TObject);
begin
  SaveDialog.InitialDir := option.OpenDataPath;
  SaveDialog.DefaultExt := option.SDataExt;
  SaveDialog.Filter := option.SSaveDlgDataFilter;
  if not SaveDialog.Execute then
    Exit;

  FDataFileName := SaveDialog.FileName;
  option.OpenDataPath := TDirectory.GetParent(FDataFileName);
  dataContainer.SaveToFile(FDataFileName, True);
  EnableFmtUIs(True);
end;

procedure TvMain.ActionSaveExecute(Sender: TObject);
begin
  SaveDialog.InitialDir := option.OpenDataPath;
  SaveDialog.DefaultExt := option.SDataExt;
  SaveDialog.Filter := option.SSaveDlgDataFilter;
  if FDataFileName.IsEmpty then
  begin
    if not SaveDialog.Execute then
      Exit;

    FDataFileName := SaveDialog.FileName;
  end;
  option.OpenDataPath := TDirectory.GetParent(FDataFileName);
  dataContainer.SaveToFile(FDataFileName, True);
  EnableFmtUIs(True);
end;

procedure TvMain.ActionSdbExecute(Sender: TObject);
begin
  if Translator.CurrentCulture.Country.Code2.Equals('KR') then
    ShellExecuteFile('http://www.sdbiosensor.com')
  else
    ShellExecuteFile('http://en.sdbiosensor.com');
end;

procedure TvMain.ActionSplitButton(Sender: TObject);
begin
  // 주석을 지우지 마세요. 해당 메뉴가 비활성화 됩니다.
end;

procedure TvMain.ActionTabTestExecute(Sender: TObject);
var
  LAction: TAction absolute Sender;
begin
  CodeSite.Send(LAction.Caption);
end;

procedure TvMain.ActionTestPropertyExecute(Sender: TObject);
var
  LAction: TAction absolute Sender;
begin
  vTestProperty.Initialize;
  PanelProperty.Visible := LAction.Checked;
  SplitterProperty.Visible := LAction.Checked;
  SplitterProperty.Left := PanelProperty.Left + PanelProperty.Width +1;

  UpdateStatusBar;
end;

procedure TvMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := True;
  if stdCalc.HasResult and FDataFileName.IsEmpty then
    if MessageDlg(Translator.GetText('All contents will be deleted. Are you sure?'), mtWarning, [mbYes, mbNo], 0, mbNo) = mrNo then
      CanClose := False;
end;

procedure TvMain.FormCreate(Sender: TObject);
begin
  //TStyleManager.Engine.UnRegisterStyleHook(TCustomTabControl, TTabControlStyleHook);
  //SetThreadPreferredUILanguages(MUI_LANGUAGE_NAME, 'en-US', nil);

//  Ribbon.BackgroundHsbColor := HsbToHsbColor(0, 0, 245);
//  Ribbon.HighlightHsbColor := HsbToHsbColor(0, 0, 230);
//  Ribbon.TextHsbColor := HsbToHsbColor(0, 0, 158);

  StatusVersion.Caption := ExeVersion;

  dataContainer.OnPropertyChange := procedure
    begin
      UpdateStatusBar;
    end;

  vRawdata := TvForm.PlaceOn<TvRawData>(TabsheetRawData);
  vRawdata.OnProperyClick := procedure
    begin
      ActionTestProperty.Execute;
      ShowNoProperyPanel;
    end;
  vRawData.OnGridChange := procedure
    begin
      EnableCalcUIs;
    end;

  vRawdataFmt := TvForm.PlaceOn<TvRawdataFmt>(PanelFmt);
  vRawdataFmt.OnCloseClick := procedure
    begin
      ActionFmtManual.Execute;
      ShowNoProperyPanel;
    end;
  vRawdataFmt.OnClearSamples := procedure
  begin
    vRawData.ClearSamples;
    EnableCalcUIs;
  end;
  vRawdataFmt.OnClearStds := procedure
  begin
    vRawData.ClearStds;
    EnableCalcUIs;
  end;

  vTestProperty := TvForm.PlaceOn<TvTestProperty>(PanelProperty);
  vTestProperty.OnCloseClick := procedure
    begin
      ActionTestProperty.Execute;
    end;

  PageControl.ActivePageIndex := 0;
  PanelCalcStd.BevelOuter := bvNone;
  PanelCalcMtrl.BevelOuter := bvNone;
  vCalcStd := TvForm.PlaceOn<TvCalcStd>(PanelCalcStd);
  vCalcMtrl := TvForm.PlaceOn<TvCalcMtrl>(PanelCalcMtrl);

//  option.EulaAgreed := False;
//  if not option.EulaAgreed then
//    if not TvEula.Open then
//    begin
//      Application.Terminate;
//      Exit;
//    end;
//
//  option.i18nAssigned := False;
//  if not option.i18nAssigned then
//    TvI18nDlg.Open;
//  TimerInit.Enabled := True;

  EnableCalcUIs;
end;

procedure TvMain.FormShow(Sender: TObject);
begin
  FocusControl(vRawdata.ActiveControl);
  TimerInit.Enabled := True;
end;

procedure TvMain.LabelPropertyAnchorClick(Sender: TObject; Anchor: string);
begin
  if Anchor.Equals('OnPropertiesClick') then
  begin
    PanelNoProp.Visible := False;
    ActionTestProperty.Execute;
  end;
end;

function TvMain.QueryCanDeleteUnsavedData: Boolean;
begin
  Result := True;
  if StdCalc.HasResult and FDataFileName.IsEmpty then
    if MessageDlg(Translator.GetText('All contents will be deleted. Are you sure?'), mtWarning, [mbYes, mbNo], 0, mbNo) = mrNo then
      Exit(False);
end;

procedure TvMain.ClearTest(const AClearWithProperties: Boolean);
begin
  FDataFileName := '';
  dataContainer.Clear(AClearWithProperties);
  stdCalc.Clear;
  vCalcStd.Clear;
  vCalcMtrl.Clear;
  PageControl.ActivePage := TabSheetRawData;
  EnableFmtUIs(False);
  EnableCalcUIs;
end;

procedure TvMain.ShowNoProperyPanel;
begin
  PanelNoProp.Visible := not dataContainer.HasProperties and not ActionTestProperty.Checked;
end;

procedure TvMain.TimerInitTimer(Sender: TObject);
begin
  TimerInit.Enabled := False;

  //option.EulaAgreed := False;
  if not option.EulaAgreed then
    if not TvEula.Open then
    begin
      Application.Terminate;
      Exit;
    end;

  //option.i18nAssigned := False;
  if not option.i18nAssigned then
  begin
    //TvI18nDlg.Open;
    TvOption.Open;
  end;

  TvTestPropertyDlg.Open;
  ShowNoProperyPanel;
end;

procedure TvMain.TranslatorAfterTranslate(Sender: TObject);
begin
  dataContainer.AssignFmtSet(FormatSettings);
end;

procedure TvMain.UpdateStatusBar;
var
  LHasPror: Boolean;
begin
  LHasPror := dataContainer.HasProperties;
  with dataContainer.Properties do
  begin
    StatusTestDate.Caption := IfThen(LHasPror and (RunDate > 0), RunDate.ToString(FormatSettings.ShortDateFormat));
    StatusTestNum.Caption := IfThen(LHasPror, RunNumber);
    StatusKitBatch.Caption := IfThen(LHasPror, KitbatchNumber);
    StatusOperator.Caption := IfThen(LHasPror, &Operator);
  end;
  ShowNoProperyPanel;
end;

end.
