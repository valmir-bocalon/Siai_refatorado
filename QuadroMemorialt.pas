unit QuadroMemorialt;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, StdCtrls, FnpNumericEdit, RLRichText;

type
  TFrm_QuadroMemorial = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLLabel1: TRLLabel;
    npg: TFnpNumericEdit;
    RLSubDetail7: TRLSubDetail;
    RLBand4: TRLBand;
    RLDBRichText1: TRLDBRichText;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_QuadroMemorial: TFrm_QuadroMemorial;

implementation

uses tabelas, funcoes, Venda, uRuntimeFields;

{$R *.dfm}


procedure TFrm_QuadroMemorial.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
