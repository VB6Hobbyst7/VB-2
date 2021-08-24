unit svc.img;

interface

uses
  System.SysUtils, System.Classes, System.ImageList, Vcl.ImgList, Vcl.Controls, PngImageList;

type
  TsvcImg = class(TDataModule)
    x16: TPngImageList;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  svcImg: TsvcImg;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

end.
