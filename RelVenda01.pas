unit RelVenda01;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport;

type
  TFrm_RelVenda01 = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLLabel1: TRLLabel;
    RLBand2: TRLBand;
    RLDBText1: TRLDBText;
    RLSubDetail1: TRLSubDetail;
    RLLabel2: TRLLabel;
    RLBand3: TRLBand;
    RLLabel3: TRLLabel;
    RLGroup1: TRLGroup;
    RLLabel4: TRLLabel;
    RLBand4: TRLBand;
    RLLabel5: TRLLabel;
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelVenda01: TFrm_RelVenda01;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}


procedure TFrm_RelVenda01.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
