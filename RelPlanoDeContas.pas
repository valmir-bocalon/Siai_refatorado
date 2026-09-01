unit RelPlanoDeContas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLFilters, RLPDFFilter, RLRichFilter, RLXLSFilter;

type
  TFrmRelPlanoDeContas = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLLabel1: TRLLabel;
    RLSystemInfo1: TRLSystemInfo;
    RLSystemInfo2: TRLSystemInfo;
    RLSystemInfo3: TRLSystemInfo;
    RLLabel2: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLBand2: TRLBand;
    RLLabel3: TRLLabel;
    RLLabel4: TRLLabel;
    RLLabel5: TRLLabel;
    RLLabel6: TRLLabel;
    RLLabel7: TRLLabel;
    RLBand3: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLLabel8: TRLLabel;
    RLDBText5: TRLDBText;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelPlanoDeContas: TFrmRelPlanoDeContas;

implementation

uses Tabelas, funcoes, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelPlanoDeContas.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel9.Caption;
  if Length(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) = 1 then
    RLLabel9.Caption:='';
  if Length(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) = 3 then
    RLLabel9.Caption:=': ';
  if Length(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) = 6 then
    RLLabel9.Caption:=':  : ';
  if Length(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) = 10 then
    RLLabel9.Caption:=':  :  : ';
  if Length(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').Text) = 15 then
    RLLabel9.Caption:=':  :  :  : ';
  RLLabel9.Caption :=RLLabel9.Caption+masccontabil(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').AsString);

end;

procedure TFrmRelPlanoDeContas.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel10.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelPlanoDeContas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
