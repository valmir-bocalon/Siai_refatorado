unit RelPrice;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLXLSFilter,
  RLRichFilter;

type
  TFrm_Relprice = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLLabel1: TRLLabel;
    RLLabel3: TRLLabel;
    RLLabel4: TRLLabel;
    RLSystemInfo1: TRLSystemInfo;
    RLSystemInfo2: TRLSystemInfo;
    RLLabel5: TRLLabel;
    RLLabel6: TRLLabel;
    RLLabel7: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLLabel8: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLLabel2: TRLLabel;
    RLDBText2: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel13: TRLLabel;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText9: TRLDBText;
    RLPDFFilter1: TRLPDFFilter;
    RLRichFilter1: TRLRichFilter;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Relprice: TFrm_Relprice;
  nm:integer;

implementation

uses tabelas, funcoes, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrm_Relprice.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFrm_Relprice.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_Relprice.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.


