unit v.report;

interface

uses
  m.rawData,

  mvw.vForm, Spring.Collections,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, frxClass, frxPreview, Vcl.StdCtrls, Vcl.Samples.Spin, Vcl.ExtCtrls, RzButton,
  RzRadChk, System.ImageList, Vcl.ImgList, PngImageList, Vcl.Menus, RzCmboBx, frxExportPDF, i18nCore, i18nLocalizer,
  frxExportBaseDialog;

type
  TvReport = class(TvDlg)
    frxReport: TfrxReport;
    frxPreview: TfrxPreview;
    Panel1: TPanel;
    ScrollBox1: TScrollBox;
    Label3: TLabel;
    Label4: TLabel;
    CheckImgPrint: TRzCheckBox;
    LabelSubject: TLabel;
    ComboSubID: TComboBox;
    Panel2: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    EditCopies: TSpinEdit;
    Panel3: TPanel;
    Button5: TButton;
    Imgs: TPngImageList;
    ButtonPDF: TRzBitBtn;
    ButtonPrint: TRzBitBtn;
    frxResultSet: TfrxUserDataSet;
    frxStdDataSet: TfrxUserDataSet;
    frxRawDataDataSet: TfrxUserDataSet;
    MemoM3: TMemo;
    MemoM2: TMemo;
    ComboPrinter: TComboBox;
    ComboReportType: TComboBox;
    frxPDFExport: TfrxPDFExport;
    Translator: TTranslator;
    LabelPrinterProperty: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);

    procedure frxReportBeforePrint(Sender: TfrxReportComponent);
    procedure frxReportGetValue(const VarName: string; var Value: Variant);

    procedure LabelPrinterPropertyClick(Sender: TObject);
    procedure ButtonPrintClick(Sender: TObject);
    procedure ButtonPDFClick(Sender: TObject);
    procedure CheckImgPrintClick(Sender: TObject);
    procedure ComboPrinterClick(Sender: TObject);
    procedure ComboReportTypeClick(Sender: TObject);
    procedure ComboSubIDChange(Sender: TObject);
  private
    FReportBuf: TBytesStream;
    FResultData: TfrxMasterData;
    FChildSubReport: TfrxChild;
    FHeaderResultData: TfrxHeader;
    FReportSummary: TfrxReportSummary;
    //FPageResult: TfrxReportPage;
    FSysMemoPage: TfrxSysMemoView;
    procedure BuildPrinterItems;
    procedure BuildSubIdItems;
    function MtrlIdx: Integer;
  public
    procedure Initialize;
    class procedure Open;
  end;

implementation

{$R *.dfm}

uses
  svc, v.calc.mtrl, v.calc.std, v.rawData,

  Vcl.Printers, mPrinter, System.StrUtils, mUtils.Windows, System.Math, mComboBoxHelper, System.DateUtils,
  mDateTimeHelper, System.IOUtils, mIOUtils
  ;

const
  NAllGroupReport = 0; NAllIndivisualReport = 1; NSingleReport = 2;

procedure TvReport.BuildPrinterItems;
var
  i: Integer;
begin
  ComboPrinter.Items.BeginUpdate;
  try
    ComboPrinter.Items.Clear;
    for i := 0 to Printer.Printers.Count -1 do
      ComboPrinter.Items.Add(Printer.Printers[i].Replace('&', '', [rfReplaceAll]));
    ComboPrinter.ItemIndex := Printer.PrinterIndex;
    ComboPrinter.SetDropdownCount;
    ComboPrinter.DropdownListAutoWidth;
  finally
    ComboPrinter.Items.EndUpdate;
  end;
end;

procedure TvReport.ButtonPDFClick(Sender: TObject);
const
  SExt = '.pdf';
var
  i: Integer;
  LFileName: String;
begin
  LFileName := Now.ToString(i18n.DateFmt) + IfThen(CheckImgPrint.Checked, '_Calc');
  LFileName := TPath.Combine(option.PDFPath, LFileName);
  LFileName := TPath.MakeUniqueFileName(LFileName);

  frxPDFExport.ShowDialog := False;
  frxPDFExport.Author := dataContainer.Properties.&Operator;
  frxPDFExport.Subject := dataContainer.Properties.KitBatchNumber;
  frxPDFExport.Title := frxPDFExport.Subject + frxPDFExport.Author;
  frxPDFExport.CreationTime := Now;
  //frxPDFExport.EmbeddedFonts := CheckBoxPDFEmbeddeFont.Checked;
  //frxPDFExport.OpenAfterExport := DlgOpen;
  case ComboReportType.ItemIndex of
    NAllGroupReport,
    NSingleReport  :
    begin
      frxPDFExport.ShowProgress := True;
      frxPDFExport.FileName := LFileName + SExt;
      frxReport.Export(frxPDFExport);
    end;

    NAllIndivisualReport:
    begin
      for i := 0 to ComboSubID.Items.Count -1 do
      begin
        frxPDFExport.FileName := Format('%s%s%s', [LFileName, ComboSubID.Items[i], SExt]);
        frxPDFExport.PageNumbers := i.ToString;
        frxReport.Export(frxPDFExport);
      end;
    end;
  end;
  Close;
end;

procedure TvReport.ButtonPrintClick(Sender: TObject);
begin
  frxReport.PrintOptions.Copies := EditCopies.Value;
  frxReport.Print;
end;

procedure TvReport.FormCreate(Sender: TObject);
begin
  Height := Application.MainForm.Height -30;
  Width := Application.MainForm.Width -30;

  FReportBuf := TBytesStream.Create;
  frxReport.SaveToStream(FReportBuf);
  (frxReport.FindObject('PageResult') as TfrxReportPage).Font := Application.DefaultFont;

  frxPreview.Thumbnail.BackColor := $00DFDFDF;//
  frxPreview.Thumbnail.FrameColor := frxPreview.BackColor;

  BuildPrinterItems;
  BuildSubIdItems;
  Initialize;
end;

procedure TvReport.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FReportBuf);
end;

procedure TvReport.FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
var
  LPoint: TPoint;
begin
  LPoint := frxPreview.ScreenToClient(MousePos);
  if frxPreview.Thumbnail.ClientRect.Contains(LPoint) then
  begin
    if WheelDelta < 0 then
      frxPreview.Prior
    else
      frxPreview.Next;//.SetPosition();
  end
  else if frxPreview.ClientRect.Contains(LPoint) then
    frxPreview.MouseWheelScroll(WheelDelta, Shift, MousePos)
end;

procedure TvReport.frxReportBeforePrint(Sender: TfrxReportComponent);
const
  SLabelOD = 'LabelOD'; NLabelOD = Length(SLabelOD);
  SMemoOD  = 'MemoOD'; NMemoOD = Length(SMemoOD);
var
  LBuf: TBytesStream;
  LName: String;
  i: Integer;
begin
  LName := Sender.Name;
  if LName.EQuals('LabelMitogen') or LName.EQuals('LabelDiffMtgNil') or
     LName.EQuals('MemoMitogen')or LName.EQuals('MemoDiffMtgNil') then
    Sender.Visible := vCalcMtrl.CrtrMtrl = cmNil_Antigen_Mitogen
  //----------------------------------------------------------------------- RawData (OD)
  else if LName.StartsWith(SLabelOD) then
  begin
    i := LName.Substring(NLabelOD).ToInteger;
    if i > IfThen(vCalcMtrl.CrtrMtrl = cmNil_Antigen, 1, 2) then
      Sender.Visible := (ComboReportType.ItemIndex = NAllGroupReport) and (i < dataContainer.ColCount +1{13°³, 0..12});

    if ComboReportType.ItemIndex <> NAllGroupReport then
      (Sender as TfrxMemoView).Text := Format('[%s]', [LName]);

    Sender.Left := i * 54;
    Sender.Width := 54;
  end
  else if LName.StartsWith(SMemoOD) then
  begin
    i := LName.Substring(NMemoOD).ToInteger;
    if ComboReportType.ItemIndex <> NAllGroupReport then
      Sender.Visible := i in [1..IfThen(vCalcMtrl.CrtrMtrl = cmNil_Antigen, 2, 3)];
    Sender.Left := Max(0, i + IfThen(ComboReportType.ItemIndex <> NAllGroupReport, -1)) * 54;
    Sender.Width := 54;
  end
  //----------------------------------------------------------------------- Mtrl Desc
  else if LName.Equals('LabelSig') or LName.Equals('MemoSig') or LName.Equals('LabelDate') or LName.Equals('MemoDate') then
    Sender.Visible := ComboReportType.ItemIndex = NAllGroupReport
  else if LName.Equals('MemoMtrlDesc') then
    Sender.Top := IfThen(vCalcMtrl.CrtrMtrl = cmNil_Antigen, 220, 236)
  else if String(Sender.Name).StartsWith('M2Memo') then
    Sender.Visible := vCalcMtrl.CrtrMtrl = cmNil_Antigen
  else if String(Sender.Name).StartsWith('M3Memo') then
  begin
    Sender.Visible := vCalcMtrl.CrtrMtrl = cmNil_Antigen_Mitogen;
    if Sender.Visible then
      Sender.Top := Sender.Top - 120;
  end
  //------------------------------------------------------------------------ Imgs
  else if Sender.Name = 'PictureGrid' then
  begin
    LBuf := TBytesStream.Create;
    try
      vRawData.GridAsBmp.SaveToStream(LBuf);
      LBuf.Position := 0;
      (Sender as TfrxPictureView).LoadPictureFromStream(LBuf);
    finally
      FreeAndNil(LBuf);
    end;
  end
  else if LName.Equals('PictureChart') then
  begin
    LBuf := TBytesStream.Create;
    try
      vCalcStd.ChartAsBmp(Floor(Sender.Width), Floor(Sender.Height)).SaveToStream(LBuf);
      LBuf.Position := 0;
      (Sender as TfrxPictureView).LoadPictureFromStream(LBuf);
    finally
      FreeAndNil(LBuf);
    end;
  end;
end;

procedure TvReport.frxReportGetValue(const VarName: string; var Value: Variant);
begin
  //------------------------------------------------------------------PageHader
  if VarName.Equals('ExeVer') then
    Value := 'Version: ' + ExeVersion
  else if VarName.Equals('RunDate') then
    Value := dataContainer.Properties.AsRunDateStr // '2017-04-07'
  else if VarName.Equals('Operator') then
    Value := dataContainer.Properties.&Operator //'TestOp'
  else if VarName.Equals('RunNumber') then
    Value := dataContainer.Properties.RunNumber
  else if VarName.Equals('KitBatchNumber') then
    Value := dataContainer.Properties.KitBatchNumber
  else if VarNAme.Equals('MaterialDesc') then
    Value := IfThen(vCalcMtrl.CrtrMtrl = cmNil_Antigen, MemoM2.Text, MemoM3.Text)
  else if VarNAme.Equals('ResultDesc') then
    Value := i18n.StdResult(stdCalc.Valid)
  //-------------------------------------------------------------------MtrlResult
  else if VarName.Equals('SubjectID') then
    Value := vCalcMtrl.Cells['Subject ID', MtrlIdx +1]
  else if VarName.Equals('Nil') then
    Value := vCalcMtrl.Cells['Nil', MtrlIdx +1]
  else if VarName.Equals('TBAg') then
    Value := vCalcMtrl.Cells['TB Ag', MtrlIdx +1]
  else if VarName.Equals('Mitogen') then
    Value := vCalcMtrl.Cells['Mitogen', MtrlIdx +1]
  else if VarName.Equals('DiffTBAgNil') then
    Value := vCalcMtrl.Cells['TB Ag-Nil', MtrlIdx +1]
  else if VarName.Equals('DiffMtgNil') then
    Value := vCalcMtrl.Cells['Mitogen-Nil', MtrlIdx +1]
  else if VarName.Equals('Result') then
    Value := vCalcMtrl.Cells['Result', MtrlIdx +1]
  //-------------------------------------------------------------------StdResult
  // StdTable
  else if VarName.Equals('Std') then
    Value := vCalcStd.Cells[0, frxStdDataSet.RecNo +1]
  else if VarName.Equals('Conc') then
    Value := vCalcStd.Cells[1, frxStdDataSet.RecNo +1]
  else if VarName.Equals('CV') then
    Value := vCalcStd.Cells[2, frxStdDataSet.RecNo +1]
  else if VarName.Equals('Mean') then
    Value := vCalcStd.Cells[3, frxStdDataSet.RecNo +1]
  else if VarName.Equals('QCResult') then
    Value := vCalcStd.Cells[4, frxStdDataSet.RecNo +1]
  // Formula
  else if VarName.Equals('Intercept') then
    Value := vCalcStd.Intercept + '     '
  else if VarName.Equals('Slope') then
    Value := vCalcStd.Slope+ '     '
  else if VarName.Equals('CorrelCoef') then
    Value := vCalcStd.CorreCoef + '     '
  // Raw Data(OD)
  //  Header
  else if VarName.Equals('LabelOD0') then
    Value := 'Nil'
  else if VarName.Equals('LabelOD1') then
    Value := 'Tb Ag'
  else if VarName.Equals('LabelOD2') then
    Value := 'Mitogen'
  //  Table
  else if VarName.Equals('OD00') then
    Value := vRawData.CellAsHtmls[0, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD01') then
    Value := vRawData.CellAsHtmls[1, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD02') then
    Value := vRawData.CellAsHtmls[2, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD03') then
    Value := vRawData.CellAsHtmls[3, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD04') then
    Value := vRawData.CellAsHtmls[4, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD05') then
    Value := vRawData.CellAsHtmls[5, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD06') then
    Value := vRawData.CellAsHtmls[6, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD07') then
    Value := vRawData.CellAsHtmls[7, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD08') then
    Value := vRawData.CellAsHtmls[8, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD09') then
    Value := vRawData.CellAsHtmls[9, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD10') then
    Value := vRawData.CellAsHtmls[10, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD11') then
    Value := vRawData.CellAsHtmls[11, frxRawDataDataSet.RecNo]
  else if VarName.Equals('OD12') then
    Value := vRawData.CellAsHtmls[12, frxRawDataDataSet.RecNo]

  else
    Value := VarName;
end;

procedure TvReport.Initialize;
begin
  //--------------------------------------------------------------------------------------------------Init Report
  FReportBuf.Position := 0;
  frxReport.LoadFromStream(FReportBuf);

  //FPageResult := frxReport.FindObject('PageResult') as TfrxReportPage;
  //FPageResult.ResetPageNumbers := ComboReportType.ItemIndex = NAllIndivisualReport;

  FSysMemoPage := frxReport.FindObject('SysMemoPage') as TfrxSysMemoView;
  FSysMemoPage.Visible := ComboReportType.ItemIndex = NAllGroupReport;
  //--------------------------------------------------------------------------------------------------Report Ctrls
  FHeaderResultData := frxReport.FindObject('HeaderResultData') as TfrxHeader;
  FHeaderResultData.ReprintOnNewPage := ComboReportType.ItemIndex = NAllIndivisualReport;

  FResultData := frxReport.FindObject('ResultData') as TfrxMasterData;
  FResultData.StartNewPage := ComboReportType.ItemIndex = NAllIndivisualReport;
  FResultData.FooterAfterEach := ComboReportType.ItemIndex = NAllIndivisualReport;

  FChildSubReport := frxReport.FindObject('ChildSubReport') as TfrxChild;
  FChildSubReport.StartNewPage := ComboReportType.ItemIndex = NAllGroupReport;

  FReportSummary := frxReport.FindObject('ReportSummary') as TfrxReportSummary;
  FReportSummary.StartNewPage := CheckImgPrint.Checked;
  FReportSummary.Visible := CheckImgPrint.Checked;
  //--------------------------------------------------------------------------------------------------DataSet
  frxResultSet.RangeEnd := reCount;
  frxResultSet.RangeEndCount := IfThen(ComboReportType.ItemIndex <> NSingleReport, vCalcMtrl.MtrlCnt, 1);

  frxStdDataSet.RangeEnd := reCount;
  frxStdDataSet.RangeEndCount := 4;

  frxRawDataDataSet.RangeEnd := reCount;
  frxRawDataDataSet.RangeEndCount := IfThen(ComboReportType.ItemIndex = NAllGroupReport, 8, 1);
  //--------------------------------------------------------------------------------------------------BuildPreview
  frxReport.PreviewOptions.ZoomMode := zmPageWidth;
  frxReport.ShowProgress := True;
  if frxReport.PrepareReport then
    frxReport.ShowPreparedReport;
end;

procedure TvReport.LabelPrinterPropertyClick(Sender: TObject);
begin
  if ComboPrinter.ItemSelected then
    TPrinterProperties.Open(ComboPrinter.ItemIndex);
end;

function TvReport.MtrlIdx: Integer;
begin
  Result := IfThen(ComboReportType.ItemIndex <> NSingleReport, frxResultSet.RecNo, ComboSubID.ItemIndex)
end;

class procedure TvReport.Open;
var
  LForm: TvReport;
begin
  LForm := TvReport.Create(nil);
  try
    LForm.ShowModal;
  finally
    FreeAndNil(LForm);
  end;
end;

procedure TvReport.BuildSubIdItems;
var
  LItem: string;
begin
  ComboSubID.Items.BeginUpdate;
  for LItem in dataContainer.IDArray do
    ComboSubID.Items.Add(LItem);
  ComboSubID.SetDropdownCount(10);
  ComboSubID.DropdownListAutoWidth;
  ComboSubID.ItemIndex := 0;
  ComboSubID.Items.EndUpdate;
end;

procedure TvReport.CheckImgPrintClick(Sender: TObject);
var
  LChk: TCheckBox absolute Sender;
begin
  LChk.Enabled := False;
  Initialize;
  TThread.Queue(nil, procedure
    begin
      LChk.Enabled := True;
    end);
end;

procedure TvReport.ComboReportTypeClick(Sender: TObject);
begin
  ComboReportType.Enabled := False;
  Initialize;
  TThread.Queue(nil, procedure
    begin
      ComboSubID.Enabled := ComboReportType.ItemIndex = NSingleReport;
      ComboReportType.Enabled := True;
    end);
end;

procedure TvReport.ComboPrinterClick(Sender: TObject);
begin
  if ComboPrinter.ItemSelected then
    frxReport.PrintOptions.Printer := ComboPrinter.ItemText;
end;

procedure TvReport.ComboSubIDChange(Sender: TObject);
begin
  if not ComboSubID.ItemSelected then
    Exit;

  ComboSubID.Enabled := False;
  Initialize;
  TThread.Queue(nil, procedure begin ComboSubID.Enabled := True; end);
end;

end.
