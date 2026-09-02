unit RelRecibodeQuita;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, RLFilters, RLPDFFilter, RLRichFilter, RLXLSFilter;

type
  TFrmRelRecibodeQuita = class(TForm)
    RLReport1: TRLReport;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel17: TRLLabel;
    RLBand1: TRLBand;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
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
    RLLabel19: TRLLabel;
    RLSubDetail1: TRLSubDetail;
    RLBand4: TRLBand;
    RLLabel20: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLBand3: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLSubDetail2: TRLSubDetail;
    RLBand6: TRLBand;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLBand5: TRLBand;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    RLLabel23: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel26: TRLLabel;
    RLLabel27: TRLLabel;
    RLLabel28: TRLLabel;
    RLDBText14: TRLDBText;
    RLLabel29: TRLLabel;
    RLSubDetail3: TRLSubDetail;
    RLBand7: TRLBand;
    RLDBText15: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText18: TRLDBText;
    RLDBText19: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText21: TRLDBText;
    RLDBText22: TRLDBText;
    RLBand8: TRLBand;
    RLLabel30: TRLLabel;
    RLLabel31: TRLLabel;
    RLLabel32: TRLLabel;
    RLLabel33: TRLLabel;
    RLLabel34: TRLLabel;
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLLabel38: TRLLabel;
    RLSubDetail4: TRLSubDetail;
    RLBand9: TRLBand;
    RLBand10: TRLBand;
    RLSystemInfo5: TRLSystemInfo;
    RLSystemInfo6: TRLSystemInfo;
    RLLabel45: TRLLabel;
    RLLabel46: TRLLabel;
    RLSystemInfo9: TRLSystemInfo;
    RLLabel49: TRLLabel;
    RLSystemInfo10: TRLSystemInfo;
    RLLabel47: TRLLabel;
    RLDBImage2: TRLDBImage;
    RLLabel48: TRLLabel;
    RLLabel51: TRLLabel;
    RLLabel52: TRLLabel;
    RLLabel53: TRLLabel;
    RLLabel54: TRLLabel;
    RLLabel55: TRLLabel;
    RLBand11: TRLBand;
    RLMemo1: TRLMemo;
    RLLabel56: TRLLabel;
    RLSubDetail5: TRLSubDetail;
    RLBand12: TRLBand;
    RLDBText28: TRLDBText;
    RLDBText29: TRLDBText;
    RLDBText30: TRLDBText;
    RLDBText35: TRLDBText;
    RLBand13: TRLBand;
    RLLabel57: TRLLabel;
    RLLabel61: TRLLabel;
    RLLabel64: TRLLabel;
    RLLabel65: TRLLabel;
    RLLabel58: TRLLabel;
    RLMemo2: TRLMemo;
    RLBand14: TRLBand;
    RLLabel59: TRLLabel;
    RLLabel60: TRLLabel;
    RLLabel62: TRLLabel;
    RLLabel63: TRLLabel;
    RLLabel66: TRLLabel;
    RLLabel67: TRLLabel;
    RLMemo3: TRLMemo;
    RLLabel2: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel39: TRLLabel;
    RLLabel40: TRLLabel;
    RLBand15: TRLBand;
    RLLabel41: TRLLabel;
    RLLabel42: TRLLabel;
    ERef_Baixa: TEdit;
    DS_forma: TDataSource;
    ZQForma: TZQuery;
    RLDBText4: TRLDBText;
    RLDBText23: TRLDBText;
    RLDBText24: TRLDBText;
    RLDBText25: TRLDBText;
    DS_tit: TDataSource;
    ZQTit: TZQuery;
    DS_Rateio: TDataSource;
    ZQRateio: TZQuery;




    RLLabel14: TRLLabel;
    RLDBText26: TRLDBText;
    RLLabel43: TRLLabel;
    RLDBText27: TRLDBText;
    RLLabel44: TRLLabel;
    RLLabel50: TRLLabel;
    RLLabel68: TRLLabel;
    RLDBText31: TRLDBText;
    RLBand17: TRLBand;
    RLLabel70: TRLLabel;
    RLLabel72: TRLLabel;
    RLLabel69: TRLLabel;
    RLLabel73: TRLLabel;
    RLLabel71: TRLLabel;
    RLDBText32: TRLDBText;

































































    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure ZQRateioCalcFields(DataSet: TDataSet);
    procedure ZQTitCalcFields(DataSet: TDataSet);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelRecibodeQuita: TFrmRelRecibodeQuita;

implementation

uses ImpRecibo, tabelas, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelRecibodeQuita.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if FrmImpRecibo=nil then
     FrmImpRecibo:=TFrmImpRecibo.Create(self);
  RLLabel17.Caption := FrmImpRecibo.Nomecli.Text;    // DM_Tabelas.ZQRecebimentonomecli.Value;
  RLLabel51.Caption := RLLabel17.Caption;
  RLLabel71.Caption := RLLabel17.Caption;
  RLLabel44.Caption := inttostr(ZQTit.RecordCount)+' Título(s)';
  RLLabel50.Caption := RLLabel44.Caption;
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelRecibodeQuita.ZQRateioCalcFields(DataSet: TDataSet);
begin
  ZQRateio.FieldByName('planodescgeral').AsString := ZQRateio.FieldByName('contabil').Text+' - '+ZQRateio.FieldByName('planodescricao').AsString;
end;

procedure TFrmRelRecibodeQuita.ZQTitCalcFields(DataSet: TDataSet);
begin
  ZQTit.FieldByName('valtit').AsFloat := ZQTit.FieldByName('valor').AsFloat + ZQTit.FieldByName('descon').AsFloat - ZQTit.FieldByName('juros').AsFloat;
end;


procedure TFrmRelRecibodeQuita.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelRecibodeQuita, 'ZQForma', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQForma', 'ZQFormavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecibodeQuita, 'ZQTit', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitvaltit', 'valtit', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitnomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQTit', 'ZQTitnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecibodeQuita, 'ZQRateio', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateionomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioplanodescricao', 'planodescricao', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioplanodescgeral', 'planodescgeral', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateiorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecibodeQuita, 'ZQRateio', 'ZQRateionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.

