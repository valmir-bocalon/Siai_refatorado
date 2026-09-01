unit RelVenda02;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, RLRichFilter, RLFilters, RLDraftFilter,
  RLHTMLFilter, RLPDFFilter, RLXLSFilter, DBClient, FnpNumericEdit;

type
  TFrm_RelVenda02 = class(TForm)
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
    RLSubDetail1: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLLabel9: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLSubDetail2: TRLSubDetail;
    RLBand5: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLDBText10: TRLDBText;
    RLSubDetail3: TRLSubDetail;
    RLBand6: TRLBand;
    RLLabel18: TRLLabel;
    RLBand7: TRLBand;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLDBText14: TRLDBText;
    RLBand9: TRLBand;
    RLLabel19: TRLLabel;
    RLLabel20: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    RLDBResult3: TRLDBResult;
    RLSubDetail4: TRLSubDetail;
    RLBand8: TRLBand;
    RLLabel23: TRLLabel;
    RLBand10: TRLBand;
    RLDBResult1: TRLDBResult;
    RLLabel24: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel26: TRLLabel;
    RLLabel27: TRLLabel;
    RLLabel28: TRLLabel;
    RLDBResult2: TRLDBResult;
    RLLabel29: TRLLabel;
    RLLabel30: TRLLabel;
    RLLabel31: TRLLabel;
    RLDBText22: TRLDBText;
    RLDBText23: TRLDBText;
    RLSubDetail5: TRLSubDetail;
    RLBand11: TRLBand;
    RLBand12: TRLBand;
    RLLabel32: TRLLabel;
    RLLabel34: TRLLabel;
    RLDBText26: TRLDBText;
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLLabel38: TRLLabel;
    RLDBText27: TRLDBText;
    RLDBText28: TRLDBText;
    RLDBText29: TRLDBText;
    RLDBResult4: TRLDBResult;
    RLDBResult5: TRLDBResult;
    RLDBResult6: TRLDBResult;
    RLDBResult7: TRLDBResult;
    RLLabel39: TRLLabel;
    RLLabel40: TRLLabel;
    RLLabel41: TRLLabel;
    RLLabel42: TRLLabel;
    RLDBResult8: TRLDBResult;
    RLDBResult9: TRLDBResult;
    RLDBResult10: TRLDBResult;
    RLSubDetail6: TRLSubDetail;
    RLDBText34: TRLDBText;
    RLBand13: TRLBand;
    RLLabel44: TRLLabel;
    RLLabel45: TRLLabel;
    RLLabel46: TRLLabel;
    RLLabel47: TRLLabel;
    RLLabel50: TRLLabel;
    RLDraw1: TRLDraw;
    RLSubDetail7: TRLSubDetail;
    RLDraw2: TRLDraw;
    RLDBText48: TRLDBText;
    RLDBText38: TRLDBText;
    RLDBText36: TRLDBText;
    RLDBText37: TRLDBText;
    RLDBText41: TRLDBText;
    RLDraw5: TRLDraw;
    RLBand15: TRLBand;
    RLDBText44: TRLDBText;
    RLDBText45: TRLDBText;
    RLDBText46: TRLDBText;
    RLDBText52: TRLDBText;
    RLDBText53: TRLDBText;
    RLDBText54: TRLDBText;
    RLDBText55: TRLDBText;
    RLDBText56: TRLDBText;
    RLDBText49: TRLDBText;
    DSincorp_loteame: TDataSource;
    ZQincorp_loteame: TZQuery;




    DataPQincorp_loteame: TDataSource;
    PQincorp_loteame: TZQuery;




    RLLabel43: TRLLabel;
    RLLabel48: TRLLabel;
    RLLabel55: TRLLabel;
    RLLabel56: TRLLabel;
    RLLabel57: TRLLabel;
    RLBand16: TRLBand;
    RLDBText16: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText18: TRLDBText;
    RLDBText19: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText30: TRLDBText;
    RLDBText31: TRLDBText;
    RLDBText32: TRLDBText;
    RLDBText33: TRLDBText;
    RLDBText47: TRLDBText;
    RLDBText50: TRLDBText;
    RLBand14: TRLBand;
    RLLabel49: TRLLabel;
    RLLabel51: TRLLabel;
    RLLabel52: TRLLabel;
    RLLabel53: TRLLabel;
    RLLabel54: TRLLabel;
    RLDraw3: TRLDraw;
    RLDraw4: TRLDraw;
    RLDBText35: TRLDBText;
    RLDBText39: TRLDBText;
    RLDBText40: TRLDBText;
    RLDBText42: TRLDBText;
    RLDBText43: TRLDBText;
    RLDraw6: TRLDraw;
    RLLabel58: TRLLabel;
    RLLabel59: TRLLabel;
    RLLabel60: TRLLabel;
    RLLabel61: TRLLabel;
    RLLabel62: TRLLabel;
    RLLabel63: TRLLabel;
    RLDBImage2: TRLDBImage;
    RLDBText2: TRLDBText;
    RLBand2: TRLBand;
    RLXLSFilter1: TRLXLSFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLHTMLFilter1: TRLHTMLFilter;
    RLRichFilter1: TRLRichFilter;
    RLLabel76: TRLLabel;
    nparc: TRLLabel;
    RLDBText24: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText25: TRLDBText;
    RLSubDetail8: TRLSubDetail;
    RLBand18: TRLBand;
    RLLabel72: TRLLabel;
    RLLabel73: TRLLabel;
    RLLabel74: TRLLabel;
    RLLabel64: TRLLabel;
    RLBand17: TRLBand;
    RLDBText62: TRLDBText;
    RLDBText21: TRLDBText;
    RLDBMemo1: TRLDBMemo;
    DataCabecalho: TDataSource;
    Cabecalho: TClientDataSet;
    Cabecalhocontrole: TAutoIncField;
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
    Datacabecalho2: TDataSource;
    cabecalho2controle: TAutoIncField;
    cabecalho2datavenda: TDateField;
    cabecalho2data: TDateField;
    cabecalho2imovel: TLargeintField;
    cabecalho2valorvenda: TFloatField;
    cabecalho2idcomprador: TLargeintField;
    cabecalho2idparticipante: TLargeintField;
    cabecalho2nome_parte: TStringField;
    cabecalho2historico: TMemoField;
    cabecalho2hora: TTimeField;
    Cabecalhoidvenda: TIntegerField;
    cabecalho2idvenda: TIntegerField;
    RLLabel11: TRLLabel;
    RLDBText51: TRLDBText;
    RLLabel33: TRLLabel;
    RLDBText57: TRLDBText;
    RLLabel65: TRLLabel;
    RLDBText58: TRLDBText;
    RLLabel66: TRLLabel;
    RLDBText59: TRLDBText;
    RLDBText60: TRLDBText;
    RLLabel67: TRLLabel;
    RLLabel69: TRLLabel;
    RLDraw7: TRLDraw;
    RLDBText15: TRLDBText;
    RLLabel68: TRLLabel;
    RLDBText61: TRLDBText;
    RLLabel70: TRLLabel;
    RLLabel71: TRLLabel;
    RLDBText63: TRLDBText;
    RLDBText64: TRLDBText;
    ZQAux: TZQuery;
    valortotal: TFnpNumericEdit;





















    RLLabel75: TRLLabel;
    RLLabel77: TRLLabel;
    rldesconto: TRLLabel;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand7BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1AfterPrint(Sender: TObject);
    procedure RLBand9BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand16BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand10BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3AfterPrint(Sender: TObject);
    procedure RLSubDetail4BeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure ZQincorp_loteameAfterPost(DataSet: TDataSet);
    procedure PQincorp_loteameAfterPost(DataSet: TDataSet);
    procedure RLBand17BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelVenda02: TFrm_RelVenda02;
  vixi,bxd,desc,jur:double;
  abp,bxp,dscp,jrp,salp:double;
implementation

uses RelVenda,funcoes, tabelas, principal, uRuntimeFields;
{$R *.dfm}
procedure TFrm_RelVenda02.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
//  RLReport1.NewPage;
end;

procedure TFrm_RelVenda02.RLBand7BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Frm_RelVenda.ZQEntrada.FieldByName('Substituicao').AsString='S' then
  begin
    if Frm_RelVenda.ZQcheque_rec_e.RecordCount>0 then
    begin
      RLSubDetail6.Visible:=true;
    end
    else
    begin
      RLSubDetail6.Visible:=false;
    end;
  end
  else
  begin
    RLSubDetail6.Visible:=false;
  end;
  if Frm_RelVenda.baixa.RecordCount>0 then
  begin
//     RLBand15.Visible:=true;
    RLBand15.Visible:=false;
  end
  else
  begin
    RLBand15.Visible:=false;
  end;
  if ZQincorp_loteame.RecordCount>1 then
  begin
    RLDBText49.Visible:=true;
  end
  else
  begin
    RLDBText49.Visible:=false;
  end;

{  if RLBand15.Visible=false then
     vixi:=vixi+Frm_RelVenda.ZQEntradaValor.value
  else
    vixi:=vixi+Frm_RelVenda.ZQEntradaValor.value+Frm_RelVenda.baixaVr_rec.value;}


    // para nao somar a substitui��o, basta Habilitar a linha de baixo
//  if Frm_RelVenda.ZQEntradasubstituicao.Value<>'S' then
     vixi:=vixi+ExRound(Frm_RelVenda.ZQEntrada.FieldByName('Valor').AsFloat,2);

  // nao estava somando todos os valores na entrada ai tirei somar no dia 15/08/2011 nao encontrei onde dou valor para o campo somar
//  if Frm_RelVenda.ZQEntradasomar.Value<>'N' then
//  begin
    if RLBand15.Visible=false then
       bxd:=bxd+ExRound(Frm_RelVenda.ZQBxEntrada.FieldByName('vr').AsFloat,2)
    else
       bxd:=bxd+ExRound(Frm_RelVenda.ZQBxEntrada.FieldByName('vr').AsFloat+Frm_RelVenda.baixa.FieldByName('Vr_rec').AsFloat,2);
  {end
  else  if Frm_RelVenda.ZQEntradasomar.Value='N' then
  begin
    if RLBand15.Visible=false then
       bxd:=bxd+0
    else
       bxd:=bxd+ExRound(Frm_RelVenda.baixaVr_rec.value,2);
  end;}
  jur:=jur+ExRound(Frm_RelVenda.ZQBxEntrada.FieldByName('jr').AsFloat,2);
  desc:=desc+ExRound(Frm_RelVenda.ZQBxEntrada.FieldByName('de').AsFloat,2);

  // cor do grid
  RLBand7.Color  := clWhite;
  if (Frm_RelVenda.ZQBxEntrada.FieldByName('data').AsDateTime=0) and (Frm_RelVenda.ZQEntrada.FieldByName('Dt_Vencimento').AsDateTime>date) then
       RLBand7.Color  := clWhite;
  if (Frm_RelVenda.ZQBxEntrada.FieldByName('data').AsDateTime=0) and (Frm_RelVenda.ZQEntrada.FieldByName('Dt_Vencimento').AsDateTime<date) then
     RLBand7.Color  := $00D7D7FF; //$00ECECFF; // vermelho
  if (Frm_RelVenda.ZQEntrada.FieldByName('Dt_Vencimento').AsDateTime>=Frm_RelVenda.ZQBxEntrada.FieldByName('data').AsDateTime) and (Frm_RelVenda.ZQBxEntrada.FieldByName('data').AsDateTime>0) then
     RLBand7.Color  := $00C4FFC4; //$00ECFFEC;  //verde
  if Frm_RelVenda.ZQEntrada.FieldByName('Dt_Vencimento').AsDateTime<Frm_RelVenda.ZQBxEntrada.FieldByName('data').AsDateTime then
     RLBand7.Color  := $007DFFFF; //$00DFFFFF; // amarelo


end;

procedure TFrm_RelVenda02.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  ZQincorp_loteame.open;
  PQincorp_loteame.open;
  vixi:=0;
  bxd:=0;
  desc:=0;
  jur:=0;

  abp:=0;
  bxp:=0;
  dscp:=0;
  jrp:=0;
  salp:=0;

  if Frm_RelVenda.CBPercent.Checked=true then
  begin
    RLSubDetail8.Visible:=false;
    RLSubDetail5.Visible:=false;
    RLSubDetail4.Visible:=false;
  end;


end;

procedure TFrm_RelVenda02.RLReport1AfterPrint(Sender: TObject);
begin
  ZQincorp_loteame.close;
  PQincorp_loteame.close;
end;

procedure TFrm_RelVenda02.RLBand9BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
  var
  per:double;
begin
//  RLLabel43.Caption:=transform(((vixi+jur)-desc),'###,###,##0.00');
  per:=0;
  if Frm_RelVenda.ZQVenda.FieldByName('valorvenal').AsFloat>0 then
     per:= (vixi*100)/Frm_RelVenda.ZQVenda.FieldByName('valorvenal').AsFloat;
  RLLabel43.Caption:=transform(vixi,'###,###,##0.00');
  RLLabel69.caption:='Entrada Corresponde a '+transform(per,'##0.###')+' %   do Vr. Tabela';
  RLLabel55.Caption:=transform(bxd,'###,###,##0.00');
  RLLabel56.Caption:=transform(jur,'###,###,##0.00');
  RLLabel57.Caption:=transform(desc,'###,###,##0.00');
  vixi:=0;
  per:=0;
  bxd:=0;
  desc:=0;
  jur:=0;
end;

procedure TFrm_RelVenda02.RLBand16BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Frm_RelVenda.ZQParcela.FieldByName('Substituicao').AsString='S' then
  begin
    if Frm_RelVenda.ZQcheque_rec_p.RecordCount>0 then
    begin
      RLSubDetail7.Visible:=true;
    end
    else
    begin
      RLSubDetail7.Visible:=false;
    end;
  end
  else
  begin
    RLSubDetail7.Visible:=false;
  end;
  if PQincorp_loteame.RecordCount>1 then
  begin
    RLDBText50.Visible:=true;
  end
  else
  begin
    RLDBText50.Visible:=false;
  end;



{  abp:=abp+truncar(Frm_RelVenda.ZQParcelaValor.value,4);
  salp:=salp+truncar(Frm_RelVenda.ZQParcelaSaldo.Value,4);
  bxp:=bxp+truncar(Frm_RelVenda.ZQBxParcelavr.Value,4);
  dscp:=dscp+truncar(Frm_RelVenda.ZQBxParcelade.Value,4);
  jrp:=jrp+truncar(Frm_RelVenda.ZQBxParcelajr.Value,4);}

  abp:=abp+ExRound(Frm_RelVenda.ZQParcela.FieldByName('Valor').AsFloat,2);
  salp:=salp+ExRound(Frm_RelVenda.ZQParcela.FieldByName('saldo').AsFloat,2);
  bxp:=bxp+ExRound(Frm_RelVenda.ZQBxParcela.FieldByName('vr').AsFloat,2);
//  dscp:=dscp+ExRound(Frm_RelVenda.ZQBxParcelade.Value,2);
  if (ExRound(Frm_RelVenda.ZQParcela.FieldByName('Valor').AsFloat-Frm_RelVenda.ZQBxParcela.FieldByName('vr').AsFloat,2)>0) and (Frm_RelVenda.ZQBxParcela.FieldByName('vr').AsFloat>0) then
  begin
     rldesconto.Caption:=trim(transform(ExRound((Frm_RelVenda.ZQParcela.FieldByName('Valor').AsFloat-Frm_RelVenda.ZQBxParcela.FieldByName('vr').AsFloat),2),'###,###,###,##0.00'));
     dscp:=dscp+ExRound((Frm_RelVenda.ZQParcela.FieldByName('Valor').AsFloat-Frm_RelVenda.ZQBxParcela.FieldByName('vr').AsFloat),2);
  end
  else
    rldesconto.Caption:='0.00';
  jrp:=jrp+ExRound(Frm_RelVenda.ZQBxParcela.FieldByName('jr').AsFloat,2);

  // cor do grid
  RLBand16.Color  := clWhite;
  if (Frm_RelVenda.ZQBxParcela.FieldByName('data').AsDateTime=0) and (Frm_RelVenda.ZQParcela.FieldByName('Dt_Vencimento').AsDateTime>date) then
     RLBand16.Color  := clWhite;
 if (Frm_RelVenda.ZQBxParcela.FieldByName('data').AsDateTime=0) and (Frm_RelVenda.ZQParcela.FieldByName('Dt_Vencimento').AsDateTime<date) then     
     RLBand16.Color  := $00D7D7FF; //$00ECECFF; // vermelho
  if (Frm_RelVenda.ZQParcela.FieldByName('Dt_Vencimento').AsDateTime>=Frm_RelVenda.ZQBxParcela.FieldByName('data').AsDateTime) and  ( Frm_RelVenda.ZQBxParcela.FieldByName('data').AsDateTime>0) then
     RLBand16.Color  := $00C4FFC4; //$00ECFFEC;  //verde
  if Frm_RelVenda.ZQParcela.FieldByName('Dt_Vencimento').AsDateTime<Frm_RelVenda.ZQBxParcela.FieldByName('data').AsDateTime then
     RLBand16.Color  := $007DFFFF; //$00DFFFFF; // amarelo

end;

procedure TFrm_RelVenda02.RLBand10BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel59.Caption:=transform(abp,'###,###,##0.00');
  RLLabel63.Caption:=transform(salp,'###,###,##0.00');
  RLLabel60.Caption:=transform(bxp,'###,###,##0.00');
  RLLabel61.Caption:=transform(jrp,'###,###,##0.00');
  RLLabel62.Caption:=transform(dscp,'###,###,##0.00');
  abp:=0;
  bxp:=0;
  dscp:=0;
  jrp:=0;
  salp:=0;
end;

procedure TFrm_RelVenda02.RLBand3AfterPrint(Sender: TObject);
begin
  if Frm_RelVenda.CBPercent.Checked=false then
  begin
    Cabecalho.Filtered:=false;
    Cabecalho.Filter:='idvenda='+quotedstr(Frm_RelVenda.ZQVenda.FieldByName('idvenda').Text);
    Cabecalho.Filtered:=true;
    if Cabecalho.RecordCount>0 then
    begin
       RLSubDetail8.Visible:=true;
       RLBand18.Visible:=true;
       RLBand17.Visible:=true;
    end
    else
    begin
       RLBand18.Visible:=false;
       RLBand17.Visible:=false;
       RLSubDetail8.Visible:=false;
    end;
  end;

{  if Frm_RelVenda.ZQaditamento.RecordCount>0 then
  begin
     RLSubDetail9.Visible:=true;
     RLBand4.Visible:=true;
     RLBand19.Visible:=true;
  end
  else
  begin
     RLSubDetail9.Visible:=false;
     RLBand4.Visible:=false;
     RLBand19.Visible:=false;
  end;

  if Frm_RelVenda.zqquitcao.RecordCount>0 then
  begin
     RLSubDetail10.Visible:=true;
     RLBand20.Visible:=true;
     RLBand21.Visible:=true;
  end
  else
  begin
     RLSubDetail10.Visible:=false;
     RLBand20.Visible:=false;
     RLBand21.Visible:=false;
  end;}
end;

procedure TFrm_RelVenda02.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Frm_RelVenda.ZQVenda.FieldByName('codigo_contrato_ref').AsLargeInt>0 then
  begin
    RLLabel68.Visible:=true;
    RLDBText61.Visible:=true;
  end
  else
  begin
    RLLabel68.Visible:=false;
    RLDBText61.Visible:=false;
  end;
  valortotal.Value:=0;
  ZQAux.Close;
  ZQAux.SQL.Clear;
  ZQAux.SQL.Add('Select h.idrecib,h.valor,sum(h.valor) as entradas,idvenda,valorvenda,somar,numordem, cliente,RefBaixa,idloteamento,quadralote,Substituicao,idrecebimento,venda_idvenda,documento,r.Valor,ordem,Dt_Vencimento,TipDoc,saldo,numboleto from recebimento as r  '+
                ' join venda as v on r.venda_idvenda=v.idvenda join recbxhist as h on h.idrecib=r.idrecebimento '+
                ' where documento like  ''%-E-%''  group by idvenda order by Dt_Vencimento,numordem,ordem ');
  ZQAux.Open;
  valortotal.Value:= ZQAux.FieldByName('entradas').AsFloat;

  ZQAux.Close;
  ZQAux.SQL.Clear;
  ZQAux.SQL.Add('Select h.idrecib,h.valor,sum(h.valor) as entradas,idvenda,valorvenda,somar,numordem, cliente,RefBaixa,idloteamento,quadralote,Substituicao,idrecebimento,venda_idvenda,documento,r.Valor,ordem,Dt_Vencimento,TipDoc,saldo,numboleto from recebimento as r  '+
                ' join venda as v on r.venda_idvenda=v.idvenda join recbxhist as h on h.idrecib=r.idrecebimento '+
                ' where documento like  ''%-P-%''  group by idvenda order by Dt_Vencimento,numordem,ordem ');
  ZQAux.Open;
  valortotal.Value:= valortotal.Value+ZQAux.FieldByName('entradas').AsFloat;
  ZQAux.Close;
  RLLabel75.Caption:=trim(transform(valortotal.Value,'###,###,###,##0.00'));
end;

procedure TFrm_RelVenda02.RLSubDetail4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  nparc.Caption:=inttostr(Frm_RelVenda.ZQParcela.recordcount);
end;

procedure TFrm_RelVenda02.ZQincorp_loteameAfterPost(DataSet: TDataSet);
begin
  ZQincorp_loteame.ApplyUpdates;
end;

procedure TFrm_RelVenda02.PQincorp_loteameAfterPost(DataSet: TDataSet);
begin
  PQincorp_loteame.ApplyUpdates;
end;

procedure TFrm_RelVenda02.RLBand17BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Cabecalhodata.Value=0 then
  begin
    RLDBText62.Visible:=false;
    RLDBText21.Visible:=false;
  end
  else
  begin
    RLDBText62.Visible:=true;
    RLDBText21.Visible:=true;
  end;
end;

procedure TFrm_RelVenda02.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelVenda02.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_RelVenda02, 'ZQincorp_loteame', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteamenconta', 'nconta', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteameconta_v', 'conta_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta_v', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteameagencia', 'agencia', TWideStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteameagencia_v', 'agencia_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia_v', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQincorp_loteame', 'ZQincorp_loteamen_banco', 'n_banco', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_banco', True);
  RegisterRuntimeDataSet(TFrm_RelVenda02, 'PQincorp_loteame', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteamenconta', 'nconta', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteameconta_v', 'conta_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta_v', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteameagencia', 'agencia', TWideStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteameagencia_v', 'agencia_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia_v', True);
  RegisterRuntimeField(TFrm_RelVenda02, 'PQincorp_loteame', 'PQincorp_loteamen_banco', 'n_banco', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_banco', True);
  RegisterRuntimeDataSet(TFrm_RelVenda02, 'ZQAux', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxentradas', 'entradas', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_RelVenda02, 'ZQAux', 'ZQAuxnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
