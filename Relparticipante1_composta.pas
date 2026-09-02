unit Relparticipante1_composta;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelparticipante_composta = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
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
    XNPrincipal: TXNumEdit;
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLLabel11: TRLLabel;
    RLBand2: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLLabel10: TRLLabel;
    RLLabel9: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLDBText6: TRLDBText;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLDBText7: TRLDBText;
    RLLabel17: TRLLabel;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLLabel18: TRLLabel;
    RLDBText10: TRLDBText;
    RLLabel19: TRLLabel;
    RLDBText11: TRLDBText;
    RLLabel20: TRLLabel;
    RLDBText12: TRLDBText;
    RLLabel21: TRLLabel;
    RLDBText13: TRLDBText;
    RLLabel22: TRLLabel;
    RLDBText14: TRLDBText;
    RLLabel23: TRLLabel;
    RLDBText15: TRLDBText;
    RLLabel24: TRLLabel;
    RLDBText16: TRLDBText;
    RLLabel25: TRLLabel;
    RLDBText17: TRLDBText;
    RLLabel26: TRLLabel;
    RLDBText18: TRLDBText;
    RLLabel27: TRLLabel;
    RLDBText19: TRLDBText;
    RLDraw1: TRLDraw;
    RLDraw2: TRLDraw;
    RLLabel28: TRLLabel;
    RLDBText20: TRLDBText;
    RLSubDetailConjuge: TRLSubDetail;
    RLBandConjugeCabecalho: TRLBand;
    RLBandConjuge: TRLBand;
    RLLabelConjugeId: TRLLabel;
    RLLabelConjugeNome: TRLLabel;
    RLLabelConjugeDoc1: TRLLabel;
    RLLabelConjugeDoc2: TRLLabel;
    RLDBTextConjugeId: TRLDBText;
    RLDBTextConjugeNome: TRLDBText;
    RLDBTextConjugeDoc2: TRLDBText;
    RLDBTextConjugeDoc1: TRLDBText;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelparticipante_composta: TFrmRelparticipante_composta;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, Participante, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelparticipante_composta.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand2.Color  := clWhite;
  if Frm_Participante.ZqParticipantes.recno mod 2 = 0 then
    RLBand2.Color  := $00EAEAEA;
  if Trim(Frm_Participante.ZqParticipantes.FieldByName('nomecidade').AsString) <> '' then
    RLDBText8.DataField := 'nomecidade'
  else
    RLDBText8.DataField := 'nomecidadenatural';
  if Trim(Frm_Participante.ZqParticipantes.FieldByName('nomeestado').AsString) <> '' then
    RLDBText9.DataField := 'nomeestado'
  else
    RLDBText9.DataField := 'nomeestadonatural';
end;

procedure TFrmRelparticipante_composta.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Frm_Participante=nil then
     Frm_Participante := TFrm_Participante.Create(self);
  FrmRelparticipante_composta.RLReport1.DataSource:=Frm_Participante.DS_Participantes;
  if DM_Tabelas.ZQCidade.Active then
     DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;
  if not Frm_Participante.ZQConjuge.Active then
     Frm_Participante.ZQConjuge.Open;
end;

procedure TFrmRelparticipante_composta.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelparticipante_composta.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
