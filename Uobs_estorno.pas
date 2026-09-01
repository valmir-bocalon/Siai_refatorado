unit Uobs_estorno;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, StdCtrls,   dxButton;

type
  TFrmObs_estorno = class(TForm)
    XBanner6: TXBanner;
    gbestorno: TGroupBox;
    XBanner13: TXBanner;
    obs_est: TMemo;
    DXBFechar: TdxButton;
    procedure DXBFecharClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmObs_estorno: TFrmObs_estorno;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

procedure TFrmObs_estorno.DXBFecharClick(Sender: TObject);
begin
  close;
end;


procedure TFrmObs_estorno.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
