unit RelReceb02_quitado2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport, Data.DB,
  Datasnap.DBClient;

type
  TFrmRelReceb02_quitado2 = class(TForm)
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
    RLBand5: TRLBand;
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLLabel10: TRLLabel;
    RLDBText2: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLDraw7: TRLDraw;
    RLDraw1: TRLDraw;
    RLBand6: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLBand7: TRLBand;
    RLSystemInfo5: TRLSystemInfo;
    RLDraw2: TRLDraw;
    RLLabel15: TRLLabel;
    RLLabel14: TRLLabel;
    RLSystemInfo6: TRLSystemInfo;
    RLDraw3: TRLDraw;
    RLLabel16: TRLLabel;
    RLDBText6: TRLDBText;
    cds_recebimento: TClientDataSet;
    cds_recebimentoidrecebimento_proposta: TLargeintField;
    cds_recebimentocliente: TIntegerField;
    cds_recebimentousuario: TIntegerField;
    cds_recebimentoDt_Entrada: TDateField;
    cds_recebimentoDt_Vencimento: TDateField;
    cds_recebimentoValor: TFloatField;
    cds_recebimentoObserv: TMemoField;
    cds_recebimentoVrDoc: TFloatField;
    cds_recebimentosaldo: TFloatField;
    cds_recebimentoRefBaixa: TIntegerField;
    cds_recebimentorefvinda: TIntegerField;
    cds_recebimentocontabil: TIntegerField;
    cds_recebimentoempresa: TIntegerField;
    cds_recebimentocustodaparcela: TFloatField;
    cds_recebimentoadversa: TIntegerField;
    cds_recebimentonumordem: TIntegerField;
    cds_recebimentoidloteamento: TIntegerField;
    cds_recebimentovenda_idvenda: TIntegerField;
    cds_recebimentosq: TLargeintField;
    cds_recebimentoData_reajuste: TDateField;
    cds_recebimentojuros: TFloatField;
    cds_recebimentodescontos: TFloatField;
    cds_recebimentoData_Quitacao: TDateField;
    DS_cds_recebimento: TDataSource;
    cds_recebimentodocumento: TStringField;
    cds_recebimentoordem: TStringField;
    cds_recebimentoTipDoc: TStringField;
    cds_recebimentomarcar: TStringField;
    cds_recebimentoorigem: TStringField;
    cds_recebimentorecpag: TStringField;
    cds_recebimentoquadralote: TStringField;
    cds_recebimentonumboleto: TStringField;
    cds_recebimentoSubstituicao: TStringField;
    cds_recebimentosomar: TStringField;
    cds_recebimentonomeadversa: TStringField;
    cds_recebimentoReajustado: TStringField;
    cds_recebimentoProximo_Reajuste: TStringField;
    cds_recebimentoParcelas_fixas: TStringField;
    cds_recebimentosld: TFloatField;
    cds_recebimentoApelido: TStringField;
    cds_recebimentonome_parte: TStringField;
    cds_recebimentoImovel: TIntegerField;
    cds_recebimentoHistorico: TMemoField;
    RLDBImage2: TRLDBImage;
    procedure RLBand5BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand5AfterPrint(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02_quitado2: TFrmRelReceb02_quitado2;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, RelRecebimento, RelVenda, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelReceb02_quitado2.RLBand5AfterPrint(Sender: TObject);
begin
  DM_Tabelas.ZQLoteamento.close;
  DM_Tabelas.ZQLoteamento.SQL.Clear;
  DM_Tabelas.ZQLoteamento.SQL.Add('select * from loteamento');
  DM_Tabelas.ZQLoteamento.Open;
end;

procedure TFrmRelReceb02_quitado2.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  Frm_RelVenda.ZQVenda2.Close;
  cds_recebimento.Close;

end;

procedure TFrmRelReceb02_quitado2.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand6.Color  := clWhite;
  if cds_recebimento.recno mod 2 = 0 then
     RLBand6.Color  := $00EAEAEA;
  if cds_recebimentoData_Quitacao.Value>0 then
     RLDBText4.DataField:='Data_Quitacao'
  else
     RLDBText4.DataField:='dt_entrada';

end;

procedure TFrmRelReceb02_quitado2.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
  DM_Tabelas.ZQLoteamento.close;
  DM_Tabelas.ZQLoteamento.SQL.Clear;
  DM_Tabelas.ZQLoteamento.SQL.Add('select * from loteamento');
  DM_Tabelas.ZQLoteamento.SQL.Add('where idloteamento='+quotedstr(cds_recebimentoidloteamento.Text));
  DM_Tabelas.ZQLoteamento.Open;
end;


procedure TFrmRelReceb02_quitado2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
