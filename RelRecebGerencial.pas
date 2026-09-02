unit RelRecebGerencial;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLFilters, RLPDFFilter, RLRichFilter, RLXLSFilter;

type
  TFrmRelRecebGerencial = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLLabel1: TRLLabel;
    RLLabel2: TRLLabel;
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
    RLBand2: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLLabel15: TRLLabel;
    RLBand3: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLSubDetail1: TRLSubDetail;
    RLBand4: TRLBand;
    RLDBText9: TRLDBText;
    RLDBMemo1: TRLDBMemo;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelRecebGerencial: TFrmRelRecebGerencial;

implementation

uses RelRecebimento, Funcoes, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelRecebGerencial.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
  varposicao : string;
begin
  RLBand2.Font.Style := [];
  RLLabel15.Font.Style  := [];
  if FrmRelRecebimento.CDSPlanoclassificacao.Value = 'A' Then Begin
    RLBand2.Font.Style := [fsBold];
    RLLabel15.Font.Style  := [fsBold];
  end;
  varposicao:='';
  if Length(FrmRelRecebimento.CDSPlanomascara.Value) = 1 then
    varposicao:='';
  if Length(FrmRelRecebimento.CDSPlanomascara.Value) = 3 then
    varposicao:=': ';
  if Length(FrmRelRecebimento.CDSPlanomascara.Value) = 6 then
    varposicao:=':  : ';
  if Length(FrmRelRecebimento.CDSPlanomascara.Value) = 10 then
    varposicao:=':  :  : ';
  if Length(FrmRelRecebimento.CDSPlanomascara.Value) = 15 then
    varposicao:=':  :  :  : ';
  RLLabel15.Caption := varposicao+masccontabil(FrmRelRecebimento.CDSPlanomascara.Value);
end;

procedure TFrmRelRecebGerencial.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelRecebGerencial.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
