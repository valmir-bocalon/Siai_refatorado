unit RelVenda02_totais_analitico;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, FnpNumericEdit, RLRichFilter, RLXLSFilter,
  RLFilters, RLPDFFilter, RLParser;

type
  TFrm_RelVenda02_totais_analitico = class(TForm)
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
    tabela: TFnpNumericEdit;
    geral: TFnpNumericEdit;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    tlinha: TFnpNumericEdit;
    RLDBImage2: TRLDBImage;
    RLDBText2: TRLDBText;
    RLDraw8: TRLDraw;
    rpar: TFnpNumericEdit;
    rentra: TFnpNumericEdit;
    rvista: TFnpNumericEdit;
    RLLabel23: TRLLabel;
    inicial: TRLLabel;
    RLLabel24: TRLLabel;
    Final: TRLLabel;
    RLGroup1: TRLGroup;
    RLBand4: TRLBand;
    RLLabel25: TRLLabel;
    RLDBText11: TRLDBText;
    RLDraw2: TRLDraw;
    RLGroup2: TRLGroup;
    RLBand8: TRLBand;
    RLLabel30: TRLLabel;
    RLDBText21: TRLDBText;
    RLDraw4: TRLDraw;
    RLBand9: TRLBand;
    RLDraw5: TRLDraw;
    RLBand12: TRLBand;
    RLDBText23: TRLDBText;
    RLDBText25: TRLDBText;
    RLDBText26: TRLDBText;
    RLDBText27: TRLDBText;
    RLDBText28: TRLDBText;
    RLDBText29: TRLDBText;
    RLDBText30: TRLDBText;
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLLabel38: TRLLabel;
    RLLabel39: TRLLabel;
    RLLabel40: TRLLabel;
    RLLabel41: TRLLabel;
    RLLabel42: TRLLabel;
    RLDBText4: TRLDBText;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLSubDetail3: TRLSubDetail;
    RLBand6: TRLBand;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel26: TRLLabel;
    RLLabel31: TRLLabel;
    RLLabel27: TRLLabel;
    RLLabel28: TRLLabel;
    RLLabel29: TRLLabel;
    RLLabel32: TRLLabel;
    RLBand7: TRLBand;
    RLDBText5: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText22: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText18: TRLDBText;
    RLDBText19: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText24: TRLDBText;
    RLDBText34: TRLDBText;
    RLDBText49: TRLDBText;
    RLBand2: TRLBand;
    RLDraw7: TRLDraw;
    RLDBResult3: TRLDBResult;
    RLDBResult4: TRLDBResult;
    RLDBResult5: TRLDBResult;
    RLDBResult6: TRLDBResult;
    RLDBResult7: TRLDBResult;
    RLLabel43: TRLLabel;
    RLLabel48: TRLLabel;
    RLLabel55: TRLLabel;
    RLLabel56: TRLLabel;
    RLLabel57: TRLLabel;
    RLLabel69: TRLLabel;
    RLBand15: TRLBand;
    RLDBText44: TRLDBText;
    RLDBText45: TRLDBText;
    RLDBText46: TRLDBText;
    RLDBText52: TRLDBText;
    RLDBText53: TRLDBText;
    RLDBText54: TRLDBText;
    RLDBText55: TRLDBText;
    RLDBText56: TRLDBText;
    RLSubDetail4: TRLSubDetail;
    RLBand5: TRLBand;
    RLLabel33: TRLLabel;
    RLLabel76: TRLLabel;
    nparc: TRLLabel;
    RLBand11: TRLBand;
    RLBand16: TRLBand;
    RLDraw3: TRLDraw;
    RLBand14: TRLBand;
    RLLabel58: TRLLabel;
    RLLabel62: TRLLabel;
    RLLabel22: TRLLabel;
    RLLabel21: TRLLabel;
    RLDBResult1: TRLDBResult;
    RLDBResult2: TRLDBResult;
    RLLabel20: TRLLabel;
    RLLabel59: TRLLabel;
    RLDraw1: TRLDraw;
    RLLabel17: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel19: TRLLabel;
    RLLabel34: TRLLabel;
    RLLabel44: TRLLabel;
    RLLabel45: TRLLabel;
    RLLabel46: TRLLabel;
    RLLabel47: TRLLabel;
    RLLabel49: TRLLabel;
    RLLabel50: TRLLabel;
    RLDBText50: TRLDBText;
    RLDBText47: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText14: TRLDBText;
    RLDBText15: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText35: TRLDBText;
    RLDBText36: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText31: TRLDBText;
    RLDBText32: TRLDBText;
    RLDBText33: TRLDBText;
    RLDBResult8: TRLDBResult;
    RLLabel51: TRLLabel;
    RLLabel52: TRLLabel;
    RLLabel63: TRLLabel;
    RLLabel60: TRLLabel;
    RLDBResult10: TRLDBResult;
    RLLabel53: TRLLabel;
    RLLabel61: TRLLabel;
    RLSystemInfo5: TRLSystemInfo;
    RLBand3: TRLBand;
    RLLabel54: TRLLabel;
    RLLabel64: TRLLabel;
    RLLabel65: TRLLabel;
    RLLabel66: TRLLabel;
    RLLabel67: TRLLabel;
    RLLabel68: TRLLabel;
    RLDBResult9: TRLDBResult;
    RLDBResult11: TRLDBResult;
    RLSystemInfo6: TRLSystemInfo;
    tabela2: TFnpNumericEdit;
    rvista2: TFnpNumericEdit;
    rentra2: TFnpNumericEdit;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand10BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand7BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand16BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand11BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand12BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand12AfterPrint(Sender: TObject);
    procedure RLBand14BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3AfterPrint(Sender: TObject);
    procedure RLSubDetail4BeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelVenda02_totais_analitico: TFrm_RelVenda02_totais_analitico;
  vixi,bxd,desc,jur:double;
  abp,bxp,dscp,jrp,salp:double;
  
implementation

uses RelVenda,funcoes, tabelas, principal, uRuntimeFields, uSiaiReportPerformance;
{$R *.dfm}
procedure TFrm_RelVenda02_totais_analitico.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
  var
  per:double;
begin
//  RLLabel43.Caption:=transform(((vixi+jur)-desc),'###,###,##0.00');
  per:=0;
  per:= (vixi*100)/Frm_RelVenda.ZQTemp4.FieldByName('valorvenal').AsFloat;
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

procedure TFrm_RelVenda02_totais_analitico.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  vixi:=0;
  bxd:=0;
  desc:=0;
  jur:=0;

  abp:=0;
  bxp:=0;
  dscp:=0;
  jrp:=0;
  salp:=0;

  tabela2.Value:=0;
  rvista2.Value:=0;
  rentra2.Value:=0;

  geral.Value:=0;
  tlinha.Value:=0;
  tabela.Value:=0;
  rvista.Value:=0;
  rentra.Value:=0;
  rpar.Value:=0;
end;

procedure TFrm_RelVenda02_totais_analitico.RLBand10BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel59.Caption:=transform(geral.Value,'###,###,##0.00');
  RLLabel62.Caption:=transform(tabela.Value,'###,###,##0.00');
  RLLabel22.Caption:=transform(rvista.Value,'###,###,##0.00');
  RLLabel21.Caption:=transform(rentra.Value,'###,###,##0.00');
  RLLabel20.Caption:=transform(rpar.Value,'###,###,##0.00');
  tabela.Value:=0;
  geral.Value:=0;
  tlinha.Value:=0;
  rvista.Value:=0;
  rentra.Value:=0;
  rpar.Value:=0;
end;

procedure TFrm_RelVenda02_totais_analitico.RLBand7BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  if Frm_RelVenda.baixa.RecordCount>0 then
  begin
//     RLBand15.Visible:=true;
    RLBand15.Visible:=false;
  end
  else
  begin
    RLBand15.Visible:=false;
  end;
  vixi:=vixi+ExRound(Frm_RelVenda.ZQEntrada.FieldByName('Valor').AsFloat,2);
  if RLBand15.Visible=false then
       bxd:=bxd+ExRound(Frm_RelVenda.ZQBxEntrada.FieldByName('vr').AsFloat,2)
  else
       bxd:=bxd+ExRound(Frm_RelVenda.ZQBxEntrada.FieldByName('vr').AsFloat+Frm_RelVenda.baixa.FieldByName('Vr_rec').AsFloat,2);

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

procedure TFrm_RelVenda02_totais_analitico.RLBand16BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  abp:=abp+ExRound(Frm_RelVenda.ZQParcela.FieldByName('Valor').AsFloat,2);
  salp:=salp+ExRound(Frm_RelVenda.ZQParcela.FieldByName('saldo').AsFloat,2);
  bxp:=bxp+ExRound(Frm_RelVenda.ZQBxParcela.FieldByName('vr').AsFloat,2);
  dscp:=dscp+ExRound(Frm_RelVenda.ZQBxParcela.FieldByName('de').AsFloat,2);
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

procedure TFrm_RelVenda02_totais_analitico.RLBand11BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  RLLabel52.Caption:=transform(abp,'###,###,##0.00');
  RLLabel63.Caption:=transform(salp,'###,###,##0.00');
  RLLabel60.Caption:=transform(bxp,'###,###,##0.00');
  RLLabel61.Caption:=transform(jrp,'###,###,##0.00');
  RLLabel53.Caption:=transform(dscp,'###,###,##0.00');
  abp:=0;
  bxp:=0;
  dscp:=0;
  jrp:=0;
  salp:=0;

end;

procedure TFrm_RelVenda02_totais_analitico.RLBand12BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  if (Frm_RelVenda.ZQTemp4.FieldByName('vr_parcela').AsFloat>0) and (Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat>0) then
  begin
    RLDBText29.Visible:=false;
    RLLabel9.Visible:=true;
    RLDBText27.Visible:=true;
    RLLabel10.Visible:=false;
    rentra.Value:=rentra.Value+Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat;
    rentra2.Value:=rentra2.Value+Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat;
//    geral.Value:=geral.Value+Frm_RelVenda.cvendavr_entrada.Value;
  end;
  if (Frm_RelVenda.ZQTemp4.FieldByName('vr_parcela').AsFloat=0) and (Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat>0) then
  begin
    rvista.Value:=rvista.Value+Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat;
    rvista2.Value:=rvista2.Value+Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat;    
//    geral.Value:=geral.Value+Frm_RelVenda.cvendavr_entrada.Value;
    RLDBText29.Visible:=true;
    RLLabel9.Visible:=false;
    RLDBText27.Visible:=false;
    RLLabel10.Visible:=true;
  end;
  tabela.Value:=tabela.Value+Frm_RelVenda.ZQTemp4.FieldByName('valorvenal').AsFloat;
  tabela2.Value:=tabela2.Value+Frm_RelVenda.ZQTemp4.FieldByName('valorvenal').AsFloat;
//  tabela.Value:=tabela.Value+Frm_RelVenda.ZQTemp4valorvenda.Value;
//  tlinha.Value:=tlinha.Value+(Frm_RelVenda.ZQTemp4valorvenda.Value+((Frm_RelVenda.ZQTemp4juros.value+Frm_RelVenda.ZQTemp4juros_p.value)-(Frm_RelVenda.ZQTemp4descontos.value-Frm_RelVenda.ZQTemp4descontos_p.value)));
  tlinha.Value:=tlinha.Value+(Frm_RelVenda.ZQTemp4.FieldByName('vr_parcela').AsFloat+Frm_RelVenda.ZQTemp4.FieldByName('vr_entrada').AsFloat);
  geral.Value:=geral.Value+tlinha.Value;
  RLLabel11.Caption:=transform(tlinha.Value,'###,###,##0.00');
  rpar.Value:=rpar.Value+Frm_RelVenda.ZQTemp4.FieldByName('vr_parcela').AsFloat;
//  geral.Value:=geral.Value+Frm_RelVenda.ZQTemp4vr_parcela.Value;
//  geral.Value:=geral.Value+Frm_RelVenda.ZQTemp4valorvenda.Value;

end;

procedure TFrm_RelVenda02_totais_analitico.RLBand12AfterPrint(
  Sender: TObject);
begin
  tlinha.Value:=0;
end;

procedure TFrm_RelVenda02_totais_analitico.RLBand14BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  RLLabel59.Caption:=transform(geral.Value,'###,###,##0.00');
  RLLabel62.Caption:=transform(tabela.Value,'###,###,##0.00');
  RLLabel22.Caption:=transform(rvista.Value,'###,###,##0.00');
  RLLabel21.Caption:=transform(rentra.Value,'###,###,##0.00');
  RLLabel20.Caption:=transform(rpar.Value,'###,###,##0.00');
  tabela.Value:=0;
  geral.Value:=0;
  tlinha.Value:=0;
  rvista.Value:=0;
  rentra.Value:=0;
  rpar.Value:=0;

end;

procedure TFrm_RelVenda02_totais_analitico.RLBand3BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  RLLabel66.Caption:=transform(tabela2.Value,'###,###,##0.00');
  RLLabel67.Caption:=transform(rvista2.Value,'###,###,##0.00');
  RLLabel68.Caption:=transform(rentra2.Value,'###,###,##0.00');
  tabela2.Value:=0;
  rvista2.Value:=0;
  rentra2.Value:=0;

end;

procedure TFrm_RelVenda02_totais_analitico.RLBand3AfterPrint(
  Sender: TObject);
begin
  tabela2.Value:=0;
  rvista2.Value:=0;
  rentra2.Value:=0;

end;

procedure TFrm_RelVenda02_totais_analitico.RLSubDetail4BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  nparc.Caption:=inttostr(Frm_RelVenda.ZQParcela.recordcount);
end;

procedure TFrm_RelVenda02_totais_analitico.RLBand1BeforePrint(
  Sender: TObject; var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelVenda02_totais_analitico.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
  MeasureReport(RLReport1, 'Relatorio administrativo - totais analiticos');
end;

end.
