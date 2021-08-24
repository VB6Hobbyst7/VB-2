unit v.progressDlg;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls;

type
  TvProgressDlg = class(TForm)
    Progress: TProgressBar;
    LabelProgress: TLabel;
  private
    procedure UpdateInfo;
  public
    class procedure Open(const AMax: Integer);
    class procedure SetepIt;
    class procedure Close;
  end;

implementation

{$R *.dfm}

var
  LForm: TvProgressDlg = nil;

{ TvProgressDlg }

class procedure TvProgressDlg.Close;
begin
  if Assigned(LForm) then
    LForm.Close;
end;

class procedure TvProgressDlg.Open(const AMax: Integer);
begin
  LForm := TvProgressDlg.Create(nil);
  try
    LForm.Progress.Max := AMax;
    LForm.Progress.Position := 0;
    LForm.UpdateInfo;
    LForm.ShowModal;
  finally
    FreeAndNil(LForm);
  end;
end;

class procedure TvProgressDlg.SetepIt;
begin
  if not Assigned(LForm) then
    Exit;

  LForm.Progress.StepIt;
end;

procedure TvProgressDlg.UpdateInfo;
begin
  LabelProgress.Caption := Format('%d/%d', [Progress.Position, Progress.Max]);
end;

end.
