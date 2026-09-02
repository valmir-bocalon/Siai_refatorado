unit RelReceb03;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLParser, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter, StdCtrls, XNum, Mask, XEdit, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, DBClient, FnpNumericEdit;

type
  TFrmRelReceb03 = class(TForm)
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
    RLExpressionParser1: TRLExpressionParser;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLReport1: TRLReport;
    XNsomabaixa: TXNumEdit;
    XVr_parcela: TXNumEdit;
    XNJuros: TXNumEdit;
    XNPrincipal: TXNumEdit;
    ZQcontabil: TZQuery;
    DScontabil: TDataSource;
    ZQParticip: TZQuery;
    DS_Particip: TDataSource;
    RLGroup1: TRLGroup;
    RLSubDetail2: TRLSubDetail;
    RLBand6: TRLBand;
    RLLabel31: TRLLabel;
    RLLabel32: TRLLabel;
    RLBand8: TRLBand;
    RLDBText21: TRLDBText;
    RLDBText22: TRLDBText;
    RLSubDetail8: TRLSubDetail;
    RLBand5: TRLBand;
    RLDBResult1: TRLDBResult;
    RLBand4: TRLBand;
    RLLabel10: TRLLabel;
    RLDBText2: TRLDBText;
    Datacabecalho2: TDataSource;
    DataCabecalho: TDataSource;
    DS_Empree: TDataSource;
    ZQEmpree: TZQuery;
    RLLabel18: TRLLabel;
    RLLabel19: TRLLabel;
    RLSubDetail1: TRLSubDetail;
    RLBand7: TRLBand;
    RLLabel20: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    RLLabel23: TRLLabel;
    RLBand9: TRLBand;
    RLDBText5: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBMemo2: TRLDBMemo;
    RLBand2: TRLBand;
    RLLabel11: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel9: TRLLabel;
    RLLabel30: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel33: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel12: TRLLabel;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    ZQparcelas: TZQuery;
    Dataparcelas: TDataSource;
    RLDraw1: TRLDraw;
    RLLabel24: TRLLabel;
    ZQrecebimento_historico_excluidos: TZQuery;
    DS_ZQrecebimento_historico_excluidos: TDataSource;






























    RLDBText9: TRLDBText;
    RLLabel25: TRLLabel;
























































































































































































































    RLLabel26: TRLLabel;

    totalrecebido: TFnpNumericEdit;
    RLBand10: TRLBand;
    RLLabel27: TRLLabel;
    RLDraw2: TRLDraw;
    RLDBResult2: TRLDBResult;
    RLDraw3: TRLDraw;
    RLDraw4: TRLDraw;
    RLDraw5: TRLDraw;
    RLDBResult3: TRLDBResult;
    RLDBResult4: TRLDBResult;
    RLDBResult5: TRLDBResult;
    XNDescontos: TXNumEdit;
    RLLabel28: TRLLabel;
    RLLabel29: TRLLabel;
    RLLabel34: TRLLabel;





































































    Cabecalho: TClientDataSet;
    Cabecalhocontrole: TAutoIncField;
    Cabecalhoidvenda: TIntegerField;
    Cabecalhodatavenda: TDateField;
    Cabecalhodata: TDateField;
    Cabecalhoimovel: TLargeintField;
    Cabecalhovalorvenda: TFloatField;
    Cabecalhoidcomprador: TLargeintField;
    Cabecalhoidparticipante: TLargeintField;
    Cabecalhonome_parte: TStringField;
    Cabecalhohistorico: TMemoField;
    Cabecalhohora: TTimeField;
    cabecalho2: TClientDataSet;
    cabecalho2controle: TAutoIncField;
    cabecalho2idvenda: TIntegerField;
    cabecalho2datavenda: TDateField;
    cabecalho2data: TDateField;
    cabecalho2imovel: TLargeintField;
    cabecalho2valorvenda: TFloatField;
    cabecalho2idcomprador: TLargeintField;
    cabecalho2idparticipante: TLargeintField;
    cabecalho2nome_parte: TStringField;
    cabecalho2historico: TMemoField;
    cabecalho2hora: TTimeField;
    RLBand11: TRLBand;
    RLDBResult6: TRLDBResult;
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLLabel38: TRLLabel;
    RLLabel39: TRLLabel;
    xjurostotal: TXNumEdit;
    xdescontostotal: TXNumEdit;
    xparcelastotal: TXNumEdit;
    xrecebidototaltotal: TFnpNumericEdit;






    RLDBText19: TRLDBText;
    RLLabel44: TRLLabel;
    RLDBText13: TRLDBText;
    RLLabel43: TRLLabel;
    RLDBText14: TRLDBText;
    RLLabel40: TRLLabel;
    RLDBText15: TRLDBText;
    RLLabel42: TRLLabel;
    RLDBText16: TRLDBText;
    RLLabel41: TRLLabel;
    RLDBText18: TRLDBText;
    procedure RLReport1AfterPrint(Sender: TObject);
    procedure RLBand17BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3AfterPrint(Sender: TObject);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand5BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand5AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand10BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand11BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand11AfterPrint(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb03: TFrmRelReceb03;
implementation

uses Tabelas, Funcoes, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrmRelReceb03.RLReport1AfterPrint(Sender: TObject);
begin
  zqcontabil.Close;
  ZQParticip.Close;
  ZQEmpree.Close;
  ZQparcelas.Close;
  Cabecalho2.close;
  cabecalho.close;
  ZQrecebimento_historico_excluidos.close;
end;

procedure TFrmRelReceb03.RLBand11AfterPrint(Sender: TObject);
begin
  totalrecebido.Value:=0;
  xparcelastotal.Value:=0;
  xjurostotal.Value:=0;
  xdescontostotal.Value:=0;
  xrecebidototaltotal.Value:=0;
  RLLabel36.Caption:='                                    ';
  RLLabel37.Caption:='                                    ';
  RLLabel38.Caption:='                                    ';
  RLLabel39.Caption:='                                    ';

end;

procedure TFrmRelReceb03.RLBand11BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel36.Caption:=transform(xrecebidototaltotal.Value,'###,###,##0.00');
  RLLabel37.Caption:=transform(xdescontostotal.Value,'###,###,##0.00');
  RLLabel38.Caption:=transform(xjurostotal.Value,'###,###,##0.00');
  RLLabel39.Caption:=transform(xparcelastotal.Value,'###,###,##0.00');



end;

procedure TFrmRelReceb03.RLBand17BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Cabecalhodata.Value=0 then
  begin
    RLDBText5.Visible:=false;
    RLDBText10.Visible:=false;
  end
  else
  begin
    RLDBText5.Visible:=true;
    RLDBText10.Visible:=true;
  end;
end;

procedure TFrmRelReceb03.RLBand3AfterPrint(Sender: TObject);
begin
  Cabecalho.Filtered:=false;
  Cabecalho.Filter:='idvenda='+quotedstr(ZQcontabil.FieldByName('idvenda').Text);
  Cabecalho.Filtered:=true;
  if Cabecalho.RecordCount>0 then
  begin
     RLSubDetail1.Visible:=true;
     RLBand9.Visible:=true;
//     RLBand7.Visible:=true;
  end
  else
  begin
     RLBand9.Visible:=false;
     RLBand7.Visible:=false;
     RLSubDetail1.Visible:=false;
  end;
end;

procedure TFrmRelReceb03.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  XNsomabaixa.Value:=0;
  XVr_parcela.Value:=0;
  XNJuros.Value:=0;
  XNDescontos.Value:=0;
  RLLabel19.Caption:='                                    ';
  RLLabel28.Caption:='                                    ';
  RLLabel29.Caption:='                                    ';
  RLLabel34.Caption:='                                    ';
end;

procedure TFrmRelReceb03.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin

{  if ZQParcelasvr_rec.Value>ZQParcelasvalor_parcela.Value then
  begin}
     RLDBText7.DataField:='valor_parcela_1';
     XNsomabaixa.Value:=XNsomabaixa.Value+ZQparcelas.FieldByName('Vr_rec').AsFloat;
     XVr_parcela.Value:=XVr_parcela.Value+ZQparcelas.FieldByName('valor_parcela_1').AsFloat;
     XNJuros.Value:=XNJuros.Value+ZQparcelas.FieldByName('Juros_Vr').AsFloat;
     XNDescontos.Value:=XNDescontos.Value+ZQparcelas.FieldByName('Desc_Vr').AsFloat;

     XNPrincipal.Value:=ZQparcelas.FieldByName('valor_parcela_1').AsFloat;//+ZQparcelasJuros_Vr.Value)-ZQparcelasDesc_Vr.Value);
     RLLabel24.Caption:=transform(ZQparcelas.FieldByName('Vr_rec').AsFloat,'###,###,##0.00');
     XNPrincipal.Value:=0;
     totalrecebido.Value:=totalrecebido.Value+ZQparcelas.FieldByName('Vr_rec').AsFloat;
{  end
  else
  begin
     if ZQParcelasvr_rec.Value>0 then
     begin
       RLDBText7.DataField:='vr_rec';
       XNsomabaixa.Value:=XNsomabaixa.Value+ZQParcelasvr_rec.Value;  //((ZQParcelasvr_rec.Value+ZQparcelasJuros_Vr.Value)-ZQparcelasDesc_Vr.Value);
       XNPrincipal.Value:=ZQParcelasvalor_parcela.Value;  //((ZQParcelasvr_rec.Value+ZQparcelasJuros_Vr.Value)-ZQparcelasDesc_Vr.Value);
       RLLabel24.Caption:=transform(XNPrincipal.Value,'###,###,##0.00');
       XNPrincipal.Value:=0;
     end
     else
     begin
       RLDBText7.DataField:='vr_ab';
       XNsomabaixa.Value:=XNsomabaixa.Value+((ZQParcelasvr_ab.Value+ZQparcelasJuros_Vr.Value)-ZQparcelasDesc_Vr.Value);
       XNPrincipal.Value:=((ZQParcelasvr_ab.Value+ZQparcelasJuros_Vr.Value)-ZQparcelasDesc_Vr.Value);
       RLLabel24.Caption:=transform(XNPrincipal.Value,'###,###,##0.00');
       XNPrincipal.Value:=0;

     end;
  end;}
  ZQrecebimento_historico_excluidos.Close;
  ZQrecebimento_historico_excluidos.SQL.Clear;
  ZQrecebimento_historico_excluidos.SQL.Add('SELECT idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,');
  ZQrecebimento_historico_excluidos.SQL.Add('       origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,motivo,Proximo_Reajuste,sequencial ');
  ZQrecebimento_historico_excluidos.SQL.Add(' from recebimento_historico_excluidos where saldo=''0'' and quadralote='+quotedstr(ZQparcelas.FieldByName('quadralote').AsString)+' and ordem='+quotedstr(ZQParcelas.FieldByName('ordem').AsString));
  ZQrecebimento_historico_excluidos.open;
   // nome de quem pagou a parcela
  if ZQrecebimento_historico_excluidos.recordcount>0 then
  begin
     RLLabel26.Caption:=ZQrecebimento_historico_excluidos.FieldByName('nome').AsString;
  end
  else
  begin
     RLLabel26.Caption:=ZQparcelas.FieldByName('nome_parte').AsString;
  end;   
end;

procedure TFrmRelReceb03.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel19.Caption:=transform(XNsomabaixa.Value,'###,###,##0.00');
  RLLabel28.Caption:=transform(XNDescontos.Value,'###,###,##0.00');
  RLLabel29.Caption:=transform(XNJuros.Value,'###,###,##0.00');
  RLLabel34.Caption:=transform(XVr_parcela.Value,'###,###,##0.00');
  xparcelastotal.Value:=xparcelastotal.Value+XVr_parcela.Value;
  xjurostotal.Value:=xjurostotal.Value+XNJuros.Value;
  xdescontostotal.Value:=xdescontostotal.Value+XNDescontos.Value;
  xrecebidototaltotal.Value:=xrecebidototaltotal.Value+XNsomabaixa.Value;

  XVr_parcela.Value:=0;
  XNJuros.Value:=0;
  XNDescontos.Value:=0;
  XNsomabaixa.Value:=0;
end;

procedure TFrmRelReceb03.RLBand5AfterPrint(Sender: TObject);
begin
  XVr_parcela.Value:=0;
  XNJuros.Value:=0;
  XNDescontos.Value:=0;
  XNsomabaixa.Value:=0;
  RLLabel19.Caption:='                                    ';
  RLLabel19.Caption:='                                    ';
  RLLabel28.Caption:='                                    ';
  RLLabel29.Caption:='                                    ';
  RLLabel34.Caption:='                                    ';

end;

procedure TFrmRelReceb03.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;

end;

procedure TFrmRelReceb03.RLBand10BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel27.Caption:=transform(totalrecebido.Value,'###,###,##0.00');
  totalrecebido.Value:=0;
end;


procedure TFrmRelReceb03.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelReceb03, 'ZQcontabil', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilvalorvenal', 'valorvenal', TFloatField, fkLookup, 0, 0, False, '', '###,###,##0.00', '', '', 0, 'imovel', 'DM_Tabelas.ZQImovel', 'idimovel', 'valorvenal', True);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiljuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiladversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiltabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilvenda_idvenda_1', 'venda_idvenda_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilpromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabildoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiltipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabillocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiln_banco', 'n_banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiln_agencia', 'n_agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiln_agencia_v', 'n_agencia_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiln_conta', 'n_conta', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabiln_conta_v', 'n_conta_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQcontabil', 'ZQcontabilnomebanco', 'nomebanco', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelReceb03, 'ZQParticip', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticiptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticiplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticiprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipSexo', 'Sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipcex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipjb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipjudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipplantao', 'plantao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticippaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticippromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticippercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQParticip', 'ZQParticipMarcar', 'Marcar', TSmallintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelReceb03, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreeidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreematriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreetestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreetestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreetestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreetestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreeObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreebairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreepasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreenomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreetipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreeapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQEmpree', 'ZQEmpreedigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelReceb03, 'ZQparcelas', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelassq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelaspercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelassq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelassaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelassq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelassomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelastabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelaspaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvenda_idvenda_1', 'venda_idvenda_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelaspromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelaspercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelastipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelaslocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasnomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvalorvenal', 'valorvenal', TFloatField, fkLookup, 0, 0, False, '', '###,###,##0.00', '', '', 0, 'imovel', 'DM_Tabelas.ZQImovel', 'idimovel', 'valorvenal', True);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelascadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQparcelas', 'ZQparcelasvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosmotivo', 'motivo', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosnome', 'nome', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidosProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelReceb03, 'ZQrecebimento_historico_excluidos', 'ZQrecebimento_historico_excluidossequencial', 'sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.

