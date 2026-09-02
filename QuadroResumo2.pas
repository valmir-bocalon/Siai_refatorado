unit QuadroResumo2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, StdCtrls, FnpNumericEdit, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter, Vcl.Imaging.jpeg;

type
  TFrm_QuadroResumo2 = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLLabel1: TRLLabel;
    RLLabel3: TRLLabel;
    RLLabel4: TRLLabel;
    RLLabel7: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLLabel8: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLSubDetail2: TRLSubDetail;
    RLBand5: TRLBand;
    RLLabel9: TRLLabel;
    RLSubDetail1: TRLSubDetail;
    RLSubDetail3: TRLSubDetail;
    RLBand6: TRLBand;
    RLDBText27: TRLDBText;
    RLLabel38: TRLLabel;
    RLLabel39: TRLLabel;
    RLLabel40: TRLLabel;
    RLLabel41: TRLLabel;
    RLLabel42: TRLLabel;
    RLLabel43: TRLLabel;
    RLLabel44: TRLLabel;
    RLLabel45: TRLLabel;
    RLLabel46: TRLLabel;
    RLLabel47: TRLLabel;
    RLLabel48: TRLLabel;
    RLLabel49: TRLLabel;
    RLDBText28: TRLDBText;
    RLDBText29: TRLDBText;
    RLDBText30: TRLDBText;
    RLDBText31: TRLDBText;
    RLDBText32: TRLDBText;
    RLDBText33: TRLDBText;
    RLDBText34: TRLDBText;
    RLDBText35: TRLDBText;
    RLDBText36: TRLDBText;
    RLDBText37: TRLDBText;
    RLDBText38: TRLDBText;
    RLSubDetail4: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText8: TRLDBText;
    RLLabel2: TRLLabel;
    RLLabel23: TRLLabel;
    RLLabel25: TRLLabel;
    RLDBText15: TRLDBText;
    RLDBText18: TRLDBText;
    RLSubDetail7: TRLSubDetail;
    RLBand4: TRLBand;
    RLLabel11: TRLLabel;
    RLMemo1: TRLMemo;
    RLSubDetail8: TRLSubDetail;
    RLBand8: TRLBand;
    RLLabel59: TRLLabel;
    RLBand2: TRLBand;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel19: TRLLabel;
    RLLabel20: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel33: TRLLabel;
    RLLabel34: TRLLabel;
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLDBText14: TRLDBText;
    RLDBText26: TRLDBText;
    RLLabel62: TRLLabel;
    RLDBText53: TRLDBText;
    RLLabel31: TRLLabel;
    RLDBText41: TRLDBText;
    RLLabel63: TRLLabel;
    RLLabel64: TRLLabel;
    RLDBText54: TRLDBText;
    RLDBText55: TRLDBText;
    RLDraw1: TRLDraw;
    recebimento: TZQuery;
    DataSource1: TDataSource;
    cheque: TZQuery;
    DataSource2: TDataSource;

















    RLLabel57: TRLMemo;
    RLLabel28: TRLLabel;
    RLDBText20: TRLDBText;
    npg: TFnpNumericEdit;
    RLLabel56: TRLLabel;
    RLLabel22: TRLLabel;
    RLDBText19: TRLDBText;
    RLLabel51: TRLLabel;
    RLDBText56: TRLDBText;
    RLDBText23: TRLDBText;
    RLLabel30: TRLLabel;
    RLDBText21: TRLDBText;
    RLLabel32: TRLLabel;
    RLDBText22: TRLDBText;
    RLLabel52: TRLLabel;
    RLDBText24: TRLDBText;
    RLLabel55: TRLLabel;
    RLDBText44: TRLDBText;
    RLLabel50: TRLLabel;
    RLDBText57: TRLDBText;
    RLLabel29: TRLLabel;
    RLDBText40: TRLDBText;
    RLLabel27: TRLLabel;
    RLDBText42: TRLDBText;
    RLLabel24: TRLLabel;
    RLDBText16: TRLDBText;
    RLLabel26: TRLLabel;
    RLDBText17: TRLDBText;
    RLLabel53: TRLLabel;
    RLDBText25: TRLDBText;
    RLLabel54: TRLLabel;
    RLDBText39: TRLDBText;
    DataSource3: TDataSource;
    recebimento2: TZQuery;

    RLLabel68: TRLLabel;
    RLDBText60: TRLDBText;














    RLBand7: TRLBand;
    RLDraw3: TRLDraw;
    RLMemo3: TRLMemo;
    RLMemo2: TRLMemo;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLDraw2: TRLDraw;
    RLLabel80: TRLLabel;
    RLDBText43: TRLDBText;
    RLLabel81: TRLLabel;
    RLLabel82: TRLLabel;
    RLLabel83: TRLLabel;
    RLDBText61: TRLDBText;
    RLDBText62: TRLDBText;
    RLDBText63: TRLDBText;
    RLLabel84: TRLLabel;
    RLDBText64: TRLDBText;
    RLLabel85: TRLLabel;
    RLMemo4: TRLMemo;
    RLSubDetail6: TRLSubDetail;
    RLSubDetail9: TRLSubDetail;
    RLSubDetail10: TRLSubDetail;
    RLBand9: TRLBand;
    RLLabel61: TRLLabel;
    RLLabel65: TRLLabel;
    RLLabel58: TRLLabel;
    RLLabel60: TRLLabel;
    RLLabel69: TRLLabel;
    RLLabel70: TRLLabel;
    RLLabel71: TRLLabel;
    RLLabel72: TRLLabel;
    RLLabel73: TRLLabel;
    RLDraw5: TRLDraw;
    RLLabel74: TRLLabel;
    RLLabel75: TRLLabel;
    RLLabel76: TRLLabel;
    RLLabel77: TRLLabel;
    RLLabel78: TRLLabel;
    RLLabel79: TRLLabel;
    RLSubDetail11: TRLSubDetail;
    RLBand11: TRLBand;
    pagina: TRLLabel;
    RLDraw4: TRLDraw;
    RLBand12: TRLBand;
    RLLabel88: TRLLabel;
    RLMemo7: TRLMemo;
    RLMemo8: TRLMemo;
    RLLabel89: TRLLabel;
    RLDetailGrid2: TRLDetailGrid;
    RLDBText50: TRLDBText;
    RLLabel6: TRLLabel;
    RLDBText51: TRLDBText;
    RLDBText52: TRLDBText;
    RLLabel5: TRLLabel;
    RLDBText47: TRLDBText;
    RLDBText45: TRLDBText;
    RLDBText46: TRLDBText;
    RLDetailGrid1: TRLDetailGrid;
    RLDBText49: TRLDBText;
    RLLabel66: TRLLabel;
    RLDBText48: TRLDBText;
    RLDetailGrid3: TRLDetailGrid;
    RLDBText58: TRLDBText;
    RLLabel67: TRLLabel;
    RLDBText59: TRLDBText;
    RLImage1: TRLImage;
    RLLabel90: TRLLabel;
    RLBand10: TRLBand;
    RLLabel86: TRLLabel;
    RLDraw6: TRLDraw;
    RLLabel87: TRLLabel;
    RLDBText65: TRLDBText;
    RLLabel91: TRLLabel;
    RLLabel92: TRLLabel;
    RLLabel93: TRLLabel;
    RLLabel0: TRLLabel;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLSubDetail6BeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1PageEnding(Sender: TObject);
    procedure RLReport1PageStarting(Sender: TObject);
    procedure RLDetailGrid2BeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_QuadroResumo2: TFrm_QuadroResumo2;

implementation

uses tabelas, funcoes, Venda, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrm_QuadroResumo2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Frm_QuadroResumo2:=nil;
  Action:=CaFree;
end;

procedure TFrm_QuadroResumo2.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var nm: integer;
begin
  nm:=strtoint(Frm_Venda.Vias.Text);
  DM_Tabelas.ZQEmpresa.Filtered:=false;
  DM_Tabelas.ZQEmpresa.Filter:='codigo='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').text);
  DM_Tabelas.ZQEmpresa.Filtered:=true;
//  RLLabel61.Caption := DM_Tabelas.ZQEmpresanomecidade.Value+' / '+DM_Tabelas.ZQEmpresaestado.Value+', '+dataporextenso(DM_Tabelas.ZQVendadatavenda.Value);
//  RLLabel66.Caption := numeroporextenso(nm);
  DM_Tabelas.ZQEmpresa.Filtered:=false;
  DM_Tabelas.ZQvnd_quadro.close;
  DM_Tabelas.ZQvnd_quadro.sql.Clear;
  DM_Tabelas.ZQvnd_quadro.SQL.Add('select v.idvenda,v.datavenda,v.imovel,v.valorvenda,v.forma_reajuste,');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('       im.idimovel,im.loteamento_idloteamento,im.quadra,im.lote,im.lado,im.medidafrente,im.medidafundo,im.medidaesquerda,im.medidadedireita,im.area,im.confrontofrente,im.confrontofundo,im.confrontodireita,im.confrontoesquerda,');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('       im.esquinanome,im.esquinalinha,im.esquinacurva,im.valorvenal,im.Obervacaoloteamento,im.cornomapa,im.disponivel,im.matri,im.matricula,');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('       lt.idloteamento,lt.cidade_idcidade,lt.matriculaloteamento,lt.datacadastro,lt.datainalguracao,lt.dataconclusao,lt.logotipo,lt.mapa,lt.testemunha1,lt.testemunha2,lt.testemunha3,');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('       lt.testemunha4,lt.Observincorp,lt.bairro,lt.pasta_mapa,lt.nomeloteamento,lt.tipo,lt.codcontabancaria,lt.apelido,lt.digito_dif,');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('       cd.idcidade,cd.nomecid,cd.ddd,cd.emancipacao,cd.estado,cd.cepgeral ');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('  from venda as v join imovel as im on v.imovel=im.idimovel ');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('join loteamento as lt on im.loteamento_idloteamento=lt.idloteamento ');
  DM_Tabelas.ZQvnd_quadro.SQL.Add('join cidade as cd on lt.cidade_idcidade=cd.idcidade where v.idvenda='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
  DM_Tabelas.ZQvnd_quadro.open;
  RLLabel61.Caption := DM_Tabelas.ZQvnd_quadro.FieldByName('nomecid').AsString+' / '+DM_Tabelas.ZQvnd_quadro.FieldByName('estado').AsString+', '+dataporextenso(DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime);
  RLLabel58.Caption := numeroporextenso(nm);
  DM_Tabelas.ZQvnd_quadro.close;
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
  RLLabel90.Caption := RLLabel61.Caption;
end;

procedure TFrm_QuadroResumo2.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
  empresa,varlinha : string;
  vrrec,entrada:double;
  ant,ant2,ant3,tam,tam1,tam2,tam3,tam4,tam5,tam6,tam7,tam8,tam9,tam10,tam11,posi,posi2,vez,tvez:integer;
begin
{  RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count+5;
  RLDraw2.Repaint;
  RLDraw2.Refresh;}
  tam3:=0;
  tam4:=0;
  tam5:=0;
  tam6:=0;
  tam7:=0;
  tam8:=0;
  tam9:=0;
  tam10:=0;
  tam11:=0;
  ant:=0;
  ant2:=0;
  ant3:=0;
  entrada:=0;
  vrrec:=0;
  vez:=0;
  tvez:=0;
  varlinha := '';
  RLMemo1.Lines.Clear;
  if DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat>0 then
  begin
    RLMemo1.Lines.Add('Valor de Tabela: R$ '+alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat,'###,###,###,##0.00'))+' ('+extenso(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat)+'). '+#13+'Valor de Comissão: R$ '+alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat,'###,###,###,##0.00'))+' ('+extenso(DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat)+') pago pelo Compromissário Comprador a Madrid Imóveis Ltda. EPP, inscrita no CNPJ n. 10.593.342/0001-16.'+#13+'Valor a Vista: R$ '+alltrim(transform((DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat-DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat),'###,###,###,##0.00'))+' ('+extenso((DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat-DM_Tabelas.ZQVenda.FieldByName('vlr_comissao').AsFloat))+'). '+#13+'Valor Total a Prazo: R$ '+alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat,'###,###,###,##0.00'))+' ('+extenso(DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat)+'). ');
  end
  else
    RLMemo1.Lines.Add('Valor a Vista: R$ '+alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat,'###,###,###,##0.00'))+' ('+extenso(DM_Tabelas.ZQVenda.FieldByName('valorvenal').AsFloat)+'). '+#13+'Valor Total a Prazo: R$ '+alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat,'###,###,###,##0.00'))+' ('+extenso(DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat)+'). ');
  RLMemo2.Lines.Clear;
  RLMemo3.Lines.Clear;
  if Frm_Venda.XNEEntrada.Value>0 Then Begin
    varlinha := 'Entrada de R$ '+alltrim(transform(Frm_Venda.XNEEntrada.Value,'###,###,###,##0.00')) +' ('+extenso(Frm_Venda.XNEEntrada.Value)+'), ';
    if Frm_Venda.XNQtdeEnt.Value>1 Then
      varlinha := varlinha + ' em '+Frm_Venda.XNQtdeEnt.Text+' vezes '+charrem(alltrim(Frm_Venda.CBPeriodoEnt.Text));
    varlinha := varlinha +' da seguinte forma:';
    RLMemo2.Lines.Clear;
    RLMemo2.Lines.Add(varlinha);
    DM_TAbelas.CDSPagEntrada.First;
    varlinha := '';

    tam:=length(alltrim(transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00')));
{    tam3:=length(DM_Tabelas.CDSPagEntradaagencia.Value);
    tam6:=length(DM_Tabelas.CDSPagEntradaconta.Value);
    tam9:=length(DM_Tabelas.CDSPagEntradanumero.Value);}

    tam3:=12;
    tam6:=12;
    tam9:=12;

    posi2:=0;
    DM_Tabelas.CDSPagEntrada.recordcount;
    while not DM_Tabelas.CDSPagEntrada.Eof do
    begin

      varlinha:='';


      tam2:=length(alltrim(transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00')));
      posi2:=pos('.',alltrim(transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00')));

      tam4:=length(DM_Tabelas.CDSPagEntradaagencia.Value);
      tam7:=length(DM_Tabelas.CDSPagEntradaconta.Value);
      tam10:=length(DM_Tabelas.CDSPagEntradanumero.Value);
      if tam>tam2 then
      begin
        tam1:=tam-tam2;
        varlinha := varlinha + DM_Tabelas.CDSPagEntradaordem.Value+'              '+DM_Tabelas.CDSPagEntradavencimento.Text+'        '+space(tam1)+transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00');
        varlinha := varlinha + '           Tipo    '+DM_Tabelas.CDSPagEntradatipodoc.Value;
      end
      else  if tam<tam2 then
      begin
        tam1:=tam2-tam;
        tam2:=length(Copy('        ', 0, (length('        ') - tam1)));
        if posi2=0 then
           tam2:=tam2+1
        else
           tam2:=tam2+2;
        varlinha := varlinha + DM_Tabelas.CDSPagEntradaordem.Value+'              '+DM_Tabelas.CDSPagEntradavencimento.Text+space(tam2)+transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00');
        varlinha := varlinha + '           Tipo    '+DM_Tabelas.CDSPagEntradatipodoc.Value;
      end
      else
      begin
        varlinha := varlinha + DM_Tabelas.CDSPagEntradaordem.Value+'              '+DM_Tabelas.CDSPagEntradavencimento.Text+'         '+transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00');
        varlinha := varlinha + '           Tipo    '+DM_Tabelas.CDSPagEntradatipodoc.Value;
      end;

      if not empty(DM_Tabelas.CDSPagEntradabanco.Value) Then
      Begin
        if tam3>tam4 then
        begin
          tam5:=tam3-tam4;
          if ant=0 then
             ant:=tam5;
          if tam5<>ant then
             tam5:=ant;
          varlinha := varlinha + '      Bc. '+DM_Tabelas.CDSPagEntradabanco.Value;
          varlinha := varlinha +space(tam5)+'      Ag. '+DM_Tabelas.CDSPagEntradaagencia.Value;

        end
        else if tam3<tam4 then
        begin
          tam5:=tam4-tam3;
          tam4:=length(Copy('      ', 0, (length('      ') - tam5)));
          varlinha := varlinha + '      Bc. '+DM_Tabelas.CDSPagEntradabanco.Value;
          varlinha := varlinha +space(tam4)+'Ag. '+DM_Tabelas.CDSPagEntradaagencia.Value;

        end
        else if tam3=tam4 then
        begin
          varlinha := varlinha + '      Bc. '+DM_Tabelas.CDSPagEntradabanco.Value;
          varlinha := varlinha + '      Ag. '+DM_Tabelas.CDSPagEntradaagencia.Value;
        end;



        if tam6>tam7 then
        begin
          tam8:=tam6-tam7;
          ant2:=length(DM_Tabelas.CDSPagEntradaconta.Value);
          if ant3=0 then
          begin
             varlinha := varlinha +space(tam8)+ '      Cc. '+DM_Tabelas.CDSPagEntradaconta.Value;
          end
          else
          begin
             tam9:=ant3-ant2;
             if tam9>0 then
                tam9:=tam9-1
             else
                tam9:=0;
             varlinha := varlinha +space(tam8)+ '      Cc. '+DM_Tabelas.CDSPagEntradaconta.Value+space(tam9);
          end;
          ant3:=length(DM_Tabelas.CDSPagEntradaconta.Value);
          varlinha := varlinha +space(tam8)+ '      nº. '+DM_Tabelas.CDSPagEntradanumero.Value;

        end;



      end;
      tam4:=0;
      tam5:=0;
      tam7:=0;
      tam8:=0;
      tam10:=0;
      tam11:=0;

      DM_Tabelas.CDSPagEntrada.Next;
//      IF not DM_Tabelas.CDSPagEntrada.Eof Then
//        varlinha := varlinha+chr(10)+chr(13);
      RLMemo2.Lines.Add(varlinha);
{      RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count;
      RLDraw2.Repaint;
      RLDraw2.Refresh;}
    end;
{    RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count+5;
    RLDraw2.Repaint;
    RLDraw2.Refresh;}
    varlinha:='';
  end
  else
  begin
    if not DM_Tabelas.ZQvenda.Active then
       DM_Tabelas.ZQvenda.Active:=true;
    recebimento.Close;
    recebimento.SQL.Clear;
    recebimento.SQL.Add('select idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc from recebimento where documento like '+quotedstr('%-E-%'));
    recebimento.Open;
    recebimento.First;
    cheque.open;
    if recebimento.RecordCount>0 then
    begin
{      while not recebimento.Eof do
      begin
        entrada:=entrada+recebimentovalor.Value;
        inc(vez);
        recebimento.Next;
      end;}

      //
      vez:=recebimento.RecordCount;
      recebimento2.Close;
      recebimento2.SQL.Clear;
      recebimento2.SQL.Add('select idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc, SUM(valor) as vr  from recebimento where documento like '+quotedstr('%-E-%'));
      recebimento2.Open;
      recebimento2.First;

      if recebimento2.RecordCount>0 then
      begin
        entrada:=recebimento2.FieldByName('vr').AsFloat;
      end;
      recebimento2.Close;
      //


      varlinha := 'Entrada de R$ '+alltrim(transform(entrada,'###,###,###,##0.00')) +' ('+extenso(entrada)+'), ';
      posi:=pos('E',recebimento.FieldByName('documento').AsString);
      Inc(posi,2);
      varlinha := varlinha + ' em '+inttostr(vez)+' veze(s) '+charrem(alltrim(copy(recebimento.FieldByName('documento').AsString,posi,20)));
      varlinha := varlinha +' da seguinte forma:';

      RLMemo2.Lines.Clear;
      RLMemo2.Lines.Add(varlinha);
      
      recebimento.First;
      varlinha := '';

      tam3:=length(cheque.FieldByName('agencia').AsString);
      tam6:=length(cheque.FieldByName('conta').AsString);
      posi2:=0;
      tam3:=12;
      tam6:=12;
      tam9:=12;

      tam:=length(alltrim(transform(recebimento.FieldByName('Valor').AsFloat,'###,###,###,##0.00')));
      while not recebimento.Eof do
      begin

        varlinha:='';

        posi:=pos('-',recebimento.FieldByName('ordem').AsString);
        Inc(posi);
        tam2:=length(alltrim(transform(recebimento.FieldByName('Valor').AsFloat,'###,###,###,##0.00')));
        posi2:=pos('.',alltrim(transform(recebimento.FieldByName('Valor').AsFloat,'###,###,###,##0.00')));

        tam4:=length(cheque.FieldByName('agencia').AsString);
        tam7:=length(cheque.FieldByName('conta').AsString);

        if tam>tam2 then
        begin
          tam1:=tam-tam2;
          varlinha := varlinha + copy(recebimento.FieldByName('ordem').AsString,posi,20)+'            '+recebimento.FieldByName('Dt_Vencimento').Text +'        '+space(tam1)+transform(recebimento.FieldByName('Valor').AsFloat,'###,###,###,##0.00')+'            ';
          varlinha := varlinha + '            Tipo   '+recebimento.FieldByName('TipDoc').AsString;
        end
        else if tam<tam2 then
        begin
          tam1:=tam2-tam;
          tam2:=length(Copy('        ', 0, (length('        ') - tam1)));
          if posi2=0 then
             tam2:=tam2+1
          else
             tam2:=tam2+2;
          varlinha := varlinha + copy(recebimento.FieldByName('ordem').AsString,posi,20)+'            '+recebimento.FieldByName('Dt_Vencimento').Text +space(tam2)+transform(recebimento.FieldByName('Valor').AsFloat,'###,###,###,##0.00')+'            ';
          varlinha := varlinha + '            Tipo   '+recebimento.FieldByName('TipDoc').AsString;
        end
        else
        begin
          varlinha := varlinha + copy(recebimento.FieldByName('ordem').AsString,posi,20)+'            '+recebimento.FieldByName('Dt_Vencimento').Text +'         '+transform(recebimento.FieldByName('Valor').AsFloat,'###,###,###,##0.00')+'            ';
          varlinha := varlinha + '            Tipo   '+recebimento.FieldByName('TipDoc').AsString;
        end;

        if cheque.RecordCount>0 then
        begin
          if tam3>tam4 then
          begin
            tam5:=tam3-tam4;
            varlinha := varlinha + '      Bc. '+cheque.FieldByName('banco').AsString;;
            varlinha := varlinha +space(tam5)+'      Ag. '+cheque.FieldByName('agencia').AsString;
          end
          else if tam3<tam4 then
          begin
            tam5:=tam3-tam4;
            tam4:=length(Copy('      ', 0, (length('      ') - tam5)));
            varlinha := varlinha + '      Bc. '+cheque.FieldByName('banco').AsString;
            varlinha := varlinha +space(tam4)+'Ag. '+cheque.FieldByName('agencia').AsString;
          end
          else
          begin
            varlinha := varlinha + '      Bc. '+cheque.FieldByName('banco').AsString;
            varlinha := varlinha + '      Ag. '+cheque.FieldByName('agencia').AsString;
          end;


          if tam6>tam7 then
          begin
            tam8:=tam6-tam7;
            ant2:=length(cheque.FieldByName('conta').AsString);
            if ant3=0 then
            begin
               varlinha := varlinha +space(tam8)+ '      Cc. '+cheque.FieldByName('conta').AsString;
            end
            else
            begin
              tam9:=ant3-ant2;
              if tam9>0 then
                 tam9:=tam9-1
              else
                 tam9:=0;
              varlinha := varlinha +space(tam8)+ '      Cc. '+cheque.FieldByName('conta').AsString+space(tam9);
            end;
            ant3:=length(cheque.FieldByName('conta').AsString);
            varlinha := varlinha +space(tam8)+ '      nº. '+cheque.FieldByName('numero').AsString;

          end;


        end;
        recebimento.Next;
//        IF not recebimento.Eof Then
  //         varlinha := varlinha+chr(10)+chr(13);
        RLMemo2.Lines.Add(varlinha);
{        RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count;
        RLDraw2.Repaint;
        RLDraw2.Refresh;}
      end;
      RLMemo2.Lines.Add(varlinha);
{      RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count+5;
      RLDraw2.Repaint;
      RLDraw2.Refresh;}
      varlinha:='';
    end;
  end;
  recebimento.close;
  cheque.Close;
  RLMemo3.Lines.Clear;
  varlinha := '';

  if Frm_Venda.XNParcela.Value>0 Then Begin
    DM_Tabelas.CDSPagParcela.First;


    if not DM_Tabelas.ZQvenda.Active then
       DM_Tabelas.ZQvenda.Active:=true;

    recebimento.Close;
    recebimento.SQL.Clear;
    recebimento.SQL.Add('select idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc from recebimento where documento like '+quotedstr('%-P-%'));
    recebimento.Open;
    recebimento.First;
    cheque.open;
    vez:=0;
    entrada:=0;
    varlinha := '';
    if recebimento.RecordCount>0 then
    begin
    //
      vez:=recebimento.RecordCount;
      recebimento2.Close;
      recebimento2.SQL.Clear;
      recebimento2.SQL.Add('select idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc, SUM(valor) as vr  from recebimento where documento like '+quotedstr('%-P-%'));
      recebimento2.Open;
      recebimento2.First;

      if recebimento2.RecordCount>0 then
      begin
        entrada:=recebimento2.FieldByName('vr').AsFloat;
      end;
      recebimento2.Close;
      //
{      while not recebimento.Eof do
      begin
        entrada:=entrada+recebimentovalor.Value;
        inc(vez);
        recebimento.Next;
      end;}
      tvez:=vez;
      recebimento.First;
      vrrec:=recebimento.FieldByName('Valor').AsFloat;

      posi:=pos('/',recebimento.FieldByName('ordem').AsString);
      inc(posi);
      if posi>0 then
      begin
         tvez:=strtoint(alltrim(copy(recebimento.FieldByName('ordem').AsString,posi,5)));
         if vez>tvez then
         begin
            vez:=vez div tvez;
            vrrec:=vrrec*vez;
         end;
      end;
    end;

    varlinha := 'O restante de R$ '+alltrim(transform(Frm_Venda.XNParcela.Value,'###,###,###,##0.00'))+' ('+extenso(Frm_Venda.XNParcela.Value)+')';
    if Frm_Venda.XNQtdeParc.Value>1 Then
      varlinha := varlinha + ' em '+inttostr(tvez)+' vezes '+Frm_Venda.CBPeriodoParc.Text;
      varlinha := varlinha + ' de R$ '+alltrim(transform(vrrec,'###,###,###,##0.00'))+' com o primeiro vencimento em '+DM_Tabelas.CDSPagParcelavencimento.Text;
    RLMemo3.Lines.Add(varlinha);
  end
  else
  begin
    if not DM_Tabelas.ZQvenda.Active then
       DM_Tabelas.ZQvenda.Active:=true;
    recebimento.Close;
    recebimento.SQL.Clear;
    recebimento.SQL.Add('select idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc from recebimento where documento like '+quotedstr('%-P-%'));
    recebimento.Open;
    recebimento.First;
    cheque.open;
    vez:=0;
    entrada:=0;
    varlinha := '';
    if recebimento.RecordCount>0 then
    begin
{      while not recebimento.Eof do
      begin
        entrada:=entrada+recebimentovalor.Value;
        inc(vez);
        recebimento.Next;
      end;}
      //
      vez:=recebimento.RecordCount;
      recebimento2.Close;
      recebimento2.SQL.Clear;
      recebimento2.SQL.Add('select idrecebimento,venda_idvenda,documento,Valor,ordem,Dt_Vencimento,TipDoc, SUM(valor) as vr  from recebimento where documento like '+quotedstr('%-P-%'));
      recebimento2.Open;
      recebimento2.First;

      if recebimento2.RecordCount>0 then
      begin
        entrada:=recebimento2.FieldByName('vr').AsFloat;
      end;
      recebimento2.Close;
      //
      tvez:=vez;
      recebimento.First;
      vrrec:=recebimento.FieldByName('Valor').AsFloat;

      posi:=pos('/',recebimento.FieldByName('ordem').AsString);
      inc(posi);
      if posi>0 then
      begin
         tvez:=strtoint(alltrim(copy(recebimento.FieldByName('ordem').AsString,posi,5)));
         if vez>tvez then
         begin
            vez:=vez div tvez;
            vrrec:=vrrec*vez;
         end;
      end;

      posi:=0;
//      DM_Tabelas.CDSPagParcela.First;
      varlinha := 'O restante de R$ '+alltrim(transform(entrada,'###,###,###,##0.00')) +' ('+extenso(entrada)+'), ';
      posi:=pos('P',recebimento.FieldByName('documento').AsString);
      Inc(posi,2);
      varlinha := varlinha + ' em '+inttostr(tvez)+' veze(s) '+copy(recebimento.FieldByName('documento').AsString,posi,20);
 //     varlinha := varlinha +' da seguinte forma:';
      varlinha := varlinha + ' de R$ '+alltrim(transform(vrrec,'###,###,###,##0.00'))+', com o primeiro vencimento em '+recebimento.FieldByName('Dt_Vencimento').Text;
      RLMemo3.Lines.Add(varlinha);
{      recebimento.First;
      varlinha := '';
      while not recebimento.Eof do begin
        posi:=pos('-',recebimentoordem.Value);
        Inc(posi);
        varlinha := varlinha + alltrim(copy(recebimentoordem.Value,posi,20))+'  '+recebimentodt_vencimento.Text +'  '+recebimentovalor.Text;
        varlinha := varlinha + ' Tipo '+recebimentotipdoc.Value;
        if cheque.RecordCount>0 then
        begin
          varlinha := varlinha + ' Bc.'+chequebanco.Value;
          varlinha := varlinha + ' Ag.'+chequeagencia.Value;
          varlinha := varlinha + ' Cc.'+chequeconta.Value;
          varlinha := varlinha + ' nº '+chequenumero.Value;
        end;
        recebimento.Next;
        IF not recebimento.Eof Then
          varlinha := varlinha + ' | ';
      end;
      RLMemo2.Lines.Add(varlinha);}
    end;
  end;
  recebimento.Close;
  cheque.close;
  if empty(RLMemo1.Caption) then
  begin
    RLDraw1.Visible:=false;
  end
  else
  begin
    RLDraw1.Visible:=true;
  end;

//  RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count+5;
  RLDraw2.Height:=(14*rlmemo2.Lines.Count)+5;
  RLDraw2.Repaint;
  RLDraw2.Refresh;
  
  if empty(RLMemo2.Caption) then
  begin
    RLDraw2.Visible:=false;
  end
  else
  begin
    RLDraw2.Visible:=true;
  end;

  if empty(RLMemo3.Caption) then
  begin
    RLDraw3.Visible:=false;
  end
  else
  begin
    RLDraw3.Visible:=true;
  end;


  RLDraw2.Height:=RLDraw2.Height+rlmemo2.Lines.Count+5;
  RLDraw2.Repaint;
  RLDraw2.Refresh;

end;

procedure TFrm_QuadroResumo2.RLSubDetail6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFrm_QuadroResumo2.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
texto:string;  
begin
  if (not empty(DM_Tabelas.ZQVenda.FieldByName('linha').AsString)) and (DM_Tabelas.ZQVenda.FieldByName('linha').AsString<>'0') then
  begin
    texto:='Imóvel distante '+alltrim(DM_Tabelas.ZQVenda.FieldByName('linha').AsString)+' m em linha reta e '+alltrim(DM_Tabelas.ZQVenda.FieldByName('curva').AsString)+' m em curva da '+alltrim(DM_Tabelas.ZQVenda.FieldByName('esquina').AsString);
    RLLabel57.Lines.Clear;
    RLLabel57.Lines.Add(texto);
    RLLabel57.Visible:=true;
  end
  else
  begin
    RLLabel57.Lines.Clear;
    RLLabel57.Visible:=false;
  end;
  if empty(DM_Tabelas.ZQVenda.FieldByName('matricula').AsString) then
  begin
    RLDBText6.DataField:='imovel_matr';
  end
  else
  begin
    RLDBText6.DataField:='matricula';
  end;
  RLLabel56.Caption:=alltrim(DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').AsString);
  RLLabel75.Caption:= alltrim(DM_Tabelas.ZQVenda.FieldByName('quadra').AsString);
  RLLabel76.Caption:= alltrim(DM_Tabelas.ZQVenda.FieldByName('lote').AsString);
  RLLabel77.Caption:= alltrim(DM_Tabelas.ZQVenda.FieldByName('lado').AsString);
  RLLabel78.Caption:= alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('area').AsFloat,'###,##0.00'));
  if empty(DM_Tabelas.ZQVenda.FieldByName('matricula').AsString) then
  begin
    RLLabel79.Caption:=DM_Tabelas.ZQVenda.FieldByName('imovel_matr').AsString;
  end
  else
  begin
    RLLabel79.Caption:=DM_Tabelas.ZQVenda.FieldByName('matricula').AsString;
  end;

end;

procedure TFrm_QuadroResumo2.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if not empty(DM_Tabelas.ZQCompr_Resp_Dados.FieldByName('nomeconjuge').AsString) then
  begin
    RLLabel81.Visible:=true;
    RLDBText61.Visible:=true;
    RLLabel82.Visible:=true;
    RLDBText63.Visible:=true;
    RLLabel83.Visible:=true;
    RLDBText62.Visible:=true;
    RLLabel84.Visible:=true;
    RLDBText64.Visible:=true;

    RLBand3.Height:=101;
  end
  else
  begin
    RLLabel81.Visible:=False;
    RLDBText61.Visible:=False;
    RLLabel82.Visible:=False;
    RLDBText63.Visible:=False;
    RLLabel83.Visible:=False;
    RLDBText62.Visible:=False;
    RLLabel84.Visible:=False;
    RLDBText64.Visible:=False;

    RLBand3.Height:=60;

  end;


end;

procedure TFrm_QuadroResumo2.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin

 if (DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='F') or (DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='I') then
 begin
   RLLabel46.Visible:=true;
   RLDBText31.Visible:=true;
   RLLabel48.Visible:=true;
   RLDBText32.Visible:=true;
   RLLabel28.Visible:=true;
   RLDBText20.Visible:=true;
   RLLabel30.Visible:=true;
   RLDBText21.Visible:=true;
   RLLabel32.Visible:=true;
   RLDBText22.Visible:=true;
   RLLabel52.Visible:=true;
   RLDBText24.Visible:=true;
   RLLabel55.Visible:=true;
   RLDBText44.Visible:=true;
   RLLabel50.Visible:=true;
   RLDBText57.Visible:=true;
   RLLabel22.Visible:=true;
   RLDBText19.Visible:=true;
   RLLabel51.Visible:=true;
   RLDBText56.Visible:=true;
   RLLabel29.Visible:=true;
   RLDBText40.Visible:=true;
   RLLabel27.Visible:=true;
   RLDBText42.Visible:=true;
   RLLabel53.Visible:=true;
   RLDBText25.Visible:=true;
   RLLabel54.Visible:=true;
   RLDBText39.Visible:=true;
   if (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='CASADO(A)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='casado(a)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='UNIÃO ESTÁVEL') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='união estável') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='REL.ESTAVEL') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='rel.estavel') then
   begin
    RLLabel28.Visible:=true;
    RLDBText20.Visible:=true;
    RLLabel32.Visible:=true;
    RLDBText22.Visible:=true;
    RLLabel30.Visible:=true;
    RLDBText21.Visible:=true;
    RLLabel52.Visible:=true;
    RLDBText24.Visible:=true;
    RLLabel55.Visible:=true;
    RLDBText44.Visible:=true;
    RLLabel51.Visible:=true;
    RLDBText56.Visible:=true;
    RLLabel29.Visible:=true;
    RLDBText40.Visible:=true;
    RLLabel27.Visible:=true;
    RLDBText42.Visible:=true;
    RLLabel22.Visible:=true;
    RLDBText19.Visible:=true;
    RLLabel53.Visible:=true;
    RLDBText25.Visible:=true;
    RLLabel54.Visible:=true;
    RLDBText39.Visible:=true;
   end
   else
   begin
    RLLabel28.Visible:=false;
    RLDBText20.Visible:=false;
    RLLabel32.Visible:=false;
    RLDBText22.Visible:=false;
    RLLabel30.Visible:=false;
    RLDBText21.Visible:=false;
    RLLabel52.Visible:=false;
    RLDBText24.Visible:=false;
    RLLabel55.Visible:=false;
    RLDBText44.Visible:=false;
    RLLabel22.Visible:=false;
    RLDBText19.Visible:=false;
    RLLabel51.Visible:=false;
    RLDBText56.Visible:=false;
    RLLabel29.Visible:=false;
    RLDBText40.Visible:=false;
    RLLabel27.Visible:=false;
    RLDBText42.Visible:=false;
    RLLabel53.Visible:=false;
    RLDBText25.Visible:=false;
    RLLabel54.Visible:=false;
    RLDBText39.Visible:=false;
   end;

   if (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='Comunhão Universal de Bens') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='COMUNHáO UNIVERSAL DE BENS') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='Comunhão de Bens') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='COMUNHáO DE BENS') then
   begin
    if empty(RLDBText56.Caption) then
    begin
      RLDBText56.DataSource:=DM_Tabelas.DS_Compr_Dados;
      RLDBText56.DataField:='cid_uf_cob';
    end
    else
    begin
      RLDBText56.DataSource:=DM_Tabelas.DS_Compr_conjuge;
      RLDBText56.DataField:='cart_cid_est';
    end;

    RLLabel51.Visible:=true;
    RLDBText56.Visible:=true;
    RLLabel50.Visible:=true;
    if empty(RLDBText57.Caption) then
    begin
      RLDBText57.DataSource:=DM_Tabelas.DS_Compr_Dados;
      RLDBText57.DataField:='cid_uf_cob';
    end
    else
    begin
      RLDBText57.DataSource:=DM_Tabelas.DS_Compr_conjuge;
      RLDBText57.DataField:='comar_cid_est';
    end;
    RLDBText57.Visible:=true;
    RLLabel30.Visible:=true;
    RLDBText21.Visible:=true;
    RLLabel32.Visible:=true;
    RLDBText22.Visible:=true;
    RLLabel52.Visible:=true;
    RLDBText24.Visible:=true;
    RLLabel55.Visible:=true;
    RLDBText44.Visible:=true;

   end
   else
   begin
    RLLabel51.Visible:=false;
    RLDBText56.Visible:=false;
    RLLabel50.Visible:=false;
    RLDBText57.Visible:=false;
    RLLabel30.Visible:=false;
    RLDBText21.Visible:=false;
    RLLabel32.Visible:=false;
    RLDBText22.Visible:=false;
    RLLabel52.Visible:=false;
    RLDBText24.Visible:=false;
    RLLabel55.Visible:=false;
    RLDBText44.Visible:=false;
   end;





   if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('complemento_cob').AsString)='' then
   begin
    RLLabel40.Visible:=false;
    RLDBText28.Visible:=false;
    RLLabel42.Top  := 144;
    RLDBText29.Top := 144;
    RLLabel41.Top  := 129;
    RLDBText36.Top := 129;
    RLLabel43.Top  := 160;
    RLDBText35.Top := 160;
    RLLabel47.Top  := 177;
    RLDBText37.Top := 177;
    RLLabel62.Top  :=193;
    RLDBText53.Top :=193;
    RLLabel63.Top  :=193;
    RLDBText54.Top :=193;
    RLLabel31.Top  :=209;
    RLDBText41.Top :=209;
    RLLabel64.Top  :=209;
    RLDBText55.Top :=209;
    
   end
   else
   begin
    RLLabel40.Visible:=true;
    RLDBText28.Visible:=true;
    RLLabel42.Top  := 160;
    RLDBText29.Top := 160;
    RLLabel41.Top  := 144;
    RLDBText36.Top := 144;
    RLLabel43.Top  := 177;
    RLDBText35.Top := 177;
    RLLabel47.Top  := 193;
    RLDBText37.Top := 193;
    RLLabel62.Top  :=209;
    RLDBText53.Top :=209;
    RLLabel63.Top  :=209;
    RLDBText54.Top :=209;
    RLLabel31.Top  :=226;
    RLDBText41.Top :=226;
    RLLabel64.Top  :=226;
    RLDBText55.Top :=226;
   end;

   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('Fone1').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('Fone1').AsString)='(00)0000-0000')  then
   begin
    RLLabel62.Visible:=false;
    RLDBText53.Visible:=false;
    if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)<>'') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)<>'(00)0000-0000')  then
    begin
       RLLabel63.Left:=8;
       RLDBText54.Left:=100;
    end;
   end
   else
   begin
    RLLabel63.Left:=212;
    RLDBText54.Left:=307;
    RLLabel62.Visible:=true;
    RLDBText53.Visible:=true;
   end;

   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)='(00)0000-0000')  then
   begin
    RLLabel63.Visible:=false;
    RLDBText54.Visible:=false;
   end
   else
   begin
    RLLabel63.Visible:=true;
    RLDBText54.Visible:=true;
   end;
   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone2').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone2').AsString)='(00)0000-0000')  then
   begin
    RLLabel31.Visible:=false;
    RLDBText41.Visible:=false;
    if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)<>'') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)<>'(00)0000-0000')  then
    begin
       RLLabel64.Left:=8;
       RLDBText55.Left:=100;
    end;

   end
   else
   begin
    RLLabel64.Left:=212;
    RLDBText55.Left:=307;
    RLLabel31.Visible:=true;
    RLDBText41.Visible:=true;
   end;

   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)='(00)0000-0000')  then
   begin
    RLLabel64.Visible:=false;
    RLDBText55.Visible:=false;
   end
   else
   begin
    RLLabel64.Visible:=true;
    RLDBText55.Visible:=true;
   end;


{ if (alltrim(DM_Tabelas.ZQCompr_DadosFone1.Value)='') and (alltrim(DM_Tabelas.ZQCompr_DadosFone2.Value)='') and (alltrim(DM_Tabelas.ZQCompr_DadosFone3.Value)='') and (alltrim(DM_Tabelas.ZQCompr_DadosFone4.Value)='')  and (alltrim(DM_Tabelas.ZQCompr_Dadosnacionalidade.Value)='')then
 begin
   RLLabel46.Top:=155;
   RLDBText31.Top:=155;
 end
 else
 begin
   RLLabel46.Top:=207;
   RLDBText31.Top:=207;
 end;

 if ((alltrim(DM_Tabelas.ZQCompr_DadosFone1.Value)<>'') or (alltrim(DM_Tabelas.ZQCompr_DadosFone3.Value)<>'')) and (alltrim(DM_Tabelas.ZQCompr_DadosFone2.Value)='') and (alltrim(DM_Tabelas.ZQCompr_DadosFone4.Value)='')  and (alltrim(DM_Tabelas.ZQCompr_Dadosnacionalidade.Value)<>'')then
 begin
   RLLabel46.Top:=190;
   RLDBText31.Top:=190;
 end
 else
 begin
   RLLabel46.Top:=207;
   RLDBText31.Top:=207;
 end;
 }
  if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('profissao').AsString)='' then
  begin
   RLLabel46.Visible:=false;
   RLDBText31.Visible:=false;
  end
  else
  begin
   RLLabel46.Visible:=true;
   RLDBText31.Visible:=true;
  end;

 {if alltrim(DM_Tabelas.ZQCompr_Dadosestadocivil.Value)='' then
 begin
   RLLabel48.Visible:=false;
   RLDBText32.Visible:=false;
   RLLabel28.Top:=8;
   RLDBText20.Top:=8;
   RLLabel30.Top:=24;
   RLDBText21.Top:=24;
   RLLabel32.Top:=40;
   RLDBText22.Top:=40;
   RLLabel52.Top:= 40;
   RLDBText24.Top:=40;
   RLLabel55.Top:= 56;
   RLDBText44.Top:=56;
   RLLabel50.Top:= 73;
   RLDBText57.Top:=73;
   RLLabel22.Top:= 89;
   RLDBText19.Top:=89;
   RLLabel51.Top:= 104;
   RLDBText56.Top:=104;
   RLLabel29.Top:= 122;
   RLDBText40.Top:=122;
   RLLabel27.Top:= 139;
   RLDBText42.Top:=139;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel48.Visible:=true;
   RLDBText32.Visible:=true;
   RLLabel28.Top:=24;
   RLDBText20.Top:=24;
   RLLabel30.Top:=40;
   RLDBText21.Top:=40;
   RLLabel32.Top:=56;
   RLDBText22.Top:=56;
   RLLabel52.Top:= 56;
   RLDBText24.Top:=56;
   RLLabel55.Top:= 73;
   RLDBText44.Top:=73;
   RLLabel50.Top:= 89;
   RLDBText57.Top:=89;
   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;


 if alltrim(DM_Tabelas.ZQCompr_conjugeregime.Value)='' then
 begin
   RLLabel28.visible:=false;
   RLDBText20.visible:=false;
   RLLabel30.Top:=24;
   RLDBText21.Top:=24;
   RLLabel32.Top:=40;
   RLDBText22.Top:=40;
   RLLabel52.Top:= 40;
   RLDBText24.Top:=40;
   RLLabel55.Top:= 56;
   RLDBText44.Top:=56;
   RLLabel50.Top:= 73;
   RLDBText57.Top:=73;
   RLLabel22.Top:= 89;
   RLDBText19.Top:=89;
   RLLabel51.Top:= 104;
   RLDBText56.Top:=104;
   RLLabel29.Top:= 122;
   RLDBText40.Top:=122;
   RLLabel27.Top:= 139;
   RLDBText42.Top:=139;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel28.Visible:=true;
   RLDBText20.Visible:=true;
   RLLabel30.Top:=40;
   RLDBText21.Top:=40;
   RLLabel32.Top:=56;
   RLDBText22.Top:=56;
   RLLabel52.Top:= 56;
   RLDBText24.Top:=56;
   RLLabel55.Top:= 73;
   RLDBText44.Top:=73;
   RLLabel50.Top:= 89;
   RLDBText57.Top:=89;
   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;



 if alltrim(DM_Tabelas.ZQCompr_conjugecartorio.Value)='' then
 begin
   RLLabel30.visible:=false;
   RLDBText21.visible:=false;
   RLLabel32.Top:=24;
   RLDBText22.Top:=24;
   RLLabel52.Top:= 24;
   RLDBText24.Top:=24;
   RLLabel55.Top:= 40;
   RLDBText44.Top:=40;
   RLLabel50.Top:= 56;
   RLDBText57.Top:=56;
   RLLabel22.Top:= 73;
   RLDBText19.Top:=73;
   RLLabel51.Top:= 89;
   RLDBText56.Top:=89;
   RLLabel29.Top:= 104;
   RLDBText40.Top:=104;
   RLLabel27.Top:= 122;
   RLDBText42.Top:=122;
   RLLabel53.Top:= 139;
   RLDBText25.Top:=139;
   RLLabel54.Top:= 154;
   RLDBText39.Top:=154;
 end
 else
 begin
   RLLabel30.Visible:=true;
   RLDBText21.Visible:=true;
   RLLabel28.Top:=24;
   RLDBText20.Top:=24;
   RLLabel30.Top:=40;
   RLDBText21.Top:=40;
   RLLabel32.Top:=56;
   RLDBText22.Top:=56;
   RLLabel52.Top:= 56;
   RLDBText24.Top:=56;
   RLLabel55.Top:= 73;
   RLDBText44.Top:=73;
   RLLabel50.Top:= 89;
   RLDBText57.Top:=89;
   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;


 if alltrim(DM_Tabelas.ZQCompr_conjugelivro.Value)='' then
 begin
   RLLabel32.visible:=false;
   RLDBText22.visible:=false;
   if alltrim(DM_Tabelas.ZQCompr_conjugefolha.Value)='' then
   begin
     RLLabel52.visible:=false;
     RLDBText24.visible:=false;
   end;
   RLLabel55.Top:= 56;
   RLDBText44.Top:=56;
   RLLabel50.Top:= 73;
   RLDBText57.Top:=73;
   RLLabel22.Top:= 89;
   RLDBText19.Top:=89;
   RLLabel51.Top:= 104;
   RLDBText56.Top:=104;
   RLLabel29.Top:= 122;
   RLDBText40.Top:=122;
   RLLabel27.Top:= 139;
   RLDBText42.Top:=139;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel32.visible:=true;
   RLDBText22.visible:=true;
   if alltrim(DM_Tabelas.ZQCompr_conjugefolha.Value)<>'' then
   begin
     RLLabel52.visible:=true;
     RLDBText24.visible:=true;
   end;
   RLLabel55.Top:= 73;
   RLDBText44.Top:=73;
   RLLabel50.Top:= 89;
   RLDBText57.Top:=89;
   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;



 if DM_Tabelas.ZQCompr_conjugedata.Value=0 then
 begin
   RLLabel55.visible:=false;
   RLDBText44.visible:=false;
   RLLabel50.Top:= 73;
   RLDBText57.Top:=73;
   RLLabel22.Top:= 89;
   RLDBText19.Top:=89;
   RLLabel51.Top:= 104;
   RLDBText56.Top:=104;
   RLLabel29.Top:= 122;
   RLDBText40.Top:=122;
   RLLabel27.Top:= 139;
   RLDBText42.Top:=139;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel55.visible:=true;
   RLDBText44.visible:=true;
   RLLabel50.Top:= 89;
   RLDBText57.Top:=89;
   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;


{ if empty(DM_Tabelas.ZQCompr_dadoscidnome_cob.Value)='' then
 begin
   RLLabel50.visible:=false;
   RLDBText57.visible:=false;
   RLLabel22.Top:= 89;
   RLDBText19.Top:=89;
   RLLabel51.Top:= 104;
   RLDBText56.Top:=104;
   RLLabel29.Top:= 122;
   RLDBText40.Top:=122;
   RLLabel27.Top:= 139;
   RLDBText42.Top:=139;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel50.visible:=true;
   RLDBText57.visible:=true;
   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;}


{ if alltrim(DM_Tabelas.ZQCompr_conjugenome.Value)='' then
 begin
   RLLabel22.visible:=false;
   RLDBText19.visible:=false;
   RLLabel51.Visible:=false;
   RLDBText56.Visible:=false;
   RLLabel29.Top:= 104;
   RLDBText40.Top:=104;
   RLLabel27.Top:= 122;
   RLDBText42.Top:=122;
   RLLabel53.Top:= 139;
   RLDBText25.Top:=139;
   RLLabel54.Top:= 154;
   RLDBText39.Top:=154;
 end
 else
 begin
   RLLabel22.visible:=true;
   RLDBText19.visible:=true;
   RLLabel51.Visible:=true;
   RLDBText56.Visible:=true;

   RLLabel22.Top:= 104;
   RLDBText19.Top:=104;
   RLLabel51.Top:= 122;
   RLDBText56.Top:=122;
   RLLabel29.Top:= 139;
   RLDBText40.Top:=139;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;



 if alltrim(DM_Tabelas.ZQCompr_conjugecidestnatural.Value)='' then
 begin
   RLLabel29.visible:=false;
   RLDBText40.visible:=false;
   RLLabel27.Top:= 139;
   RLDBText42.Top:=139;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel29.visible:=true;
   RLDBText40.visible:=true;
   RLLabel27.Top:= 154;
   RLDBText42.Top:=154;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;

 if alltrim(DM_Tabelas.ZQCompr_conjugecargo.Value)='' then
 begin
   RLLabel27.visible:=false;
   RLDBText42.visible:=false;
   RLLabel53.Top:= 154;
   RLDBText25.Top:=154;
   RLLabel54.Top:= 171;
   RLDBText39.Top:=171;
 end
 else
 begin
   RLLabel27.visible:=true;
   RLDBText42.visible:=true;
   RLLabel53.Top:= 171;
   RLDBText25.Top:=171;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;



 if alltrim(DM_Tabelas.ZQCompr_conjugedoc2.Value)='' then
 begin
   RLLabel53.visible:=false;
   RLDBText25.visible:=false;
   RLLabel54.Top:= 154;
   RLDBText39.Top:=154;
 end
 else
 begin
   RLLabel53.visible:=true;
   RLDBText25.visible:=true;
   RLLabel54.Top:= 189;
   RLDBText39.Top:=189;
 end;
}

 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc1').AsString)='' then
 begin
   RLLabel54.visible:=false;
   RLDBText39.visible:=false;
 end
 else
 begin
   RLLabel54.visible:=true;
   RLDBText39.visible:=true;
 end;




 if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString)='' then
 begin
   RLLabel48.Visible:=false;
   RLDBText32.Visible:=false;
 end
 else
 begin
   RLLabel48.Visible:=true;
   RLDBText32.Visible:=true;
 end;


 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString)='' then
 begin
   RLLabel28.visible:=false;
   RLDBText20.visible:=false;
 end
 else
 begin
   RLLabel28.Visible:=true;
   RLDBText20.Visible:=true;
 end;

 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cartorio').AsString)='' then
 begin
   RLLabel30.visible:=false;
   RLDBText21.visible:=false;
 end
 else
 begin
   RLLabel30.Visible:=true;
   RLDBText21.Visible:=true;
 end;


 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('livro').AsString)='' then
 begin
   RLLabel32.visible:=false;
   RLDBText22.visible:=false;
   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)='' then
   begin
     RLLabel52.visible:=false;
     RLDBText24.visible:=false;
   end;
 end
 else
 begin
   RLLabel32.visible:=true;
   RLDBText22.visible:=true;
   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)<>'' then
   begin
     RLLabel52.visible:=true;
     RLDBText24.visible:=true;
   end;
 end;



 if DM_Tabelas.ZQCompr_conjuge.FieldByName('data').AsDateTime=0 then
 begin
   RLLabel55.visible:=false;
   RLDBText44.visible:=false;
 end
 else
 begin
   RLLabel55.visible:=true;
   RLDBText44.visible:=true;
 end;

 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('nome').AsString)='' then
 begin
   RLLabel22.visible:=false;
   RLDBText19.visible:=false;
 end
 else
 begin
   RLLabel22.visible:=true;
   RLDBText19.visible:=true;
 end;




{ if alltrim(DM_Tabelas.ZQCompr_dadoscidnome_cob.Value)='' then
 begin
   RLLabel51.visible:=false;
   RLDBText56.visible:=false;
 end
 else
 begin
   RLLabel51.visible:=true;
   RLDBText56.visible:=true;
 end;
 }

 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cidestnatural').AsString)='' then
 begin
   RLLabel29.visible:=false;
   RLDBText40.visible:=false;
 end
 else
 begin
   RLLabel29.visible:=true;
   RLDBText40.visible:=true;
 end;

 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cargo').AsString)='' then
 begin
   RLLabel27.visible:=false;
   RLDBText42.visible:=false;
 end
 else
 begin
   RLLabel27.visible:=true;
   RLDBText42.visible:=true;
 end;



 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc2').AsString)='' then
 begin
   RLLabel53.visible:=false;
   RLDBText25.visible:=false;
 end
 else
 begin
   RLLabel53.visible:=true;
   RLDBText25.visible:=true;
 end;


 if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc1').AsString)='' then
 begin
   RLLabel54.visible:=false;
   RLDBText39.visible:=false;
 end
 else
 begin
   RLLabel54.visible:=true;
   RLDBText39.visible:=true;
 end;



 end
 else
 begin
   RLLabel46.Visible:=false;
   RLDBText31.Visible:=false;
   RLLabel48.Visible:=false;
   RLDBText32.Visible:=false;
   RLLabel28.Visible:=false;
   RLDBText20.Visible:=false;
   RLLabel30.Visible:=false;
   RLDBText21.Visible:=false;
   RLLabel32.Visible:=false;
   RLDBText22.Visible:=false;
   RLLabel52.Visible:=false;
   RLDBText24.Visible:=false;
   RLLabel55.Visible:=false;
   RLDBText44.Visible:=false;
   RLLabel50.Visible:=false;
   RLDBText57.Visible:=false;
   RLLabel22.Visible:=false;
   RLDBText19.Visible:=false;
   RLLabel51.Visible:=false;
   RLDBText56.Visible:=false;
   RLLabel29.Visible:=false;
   RLDBText40.Visible:=false;
   RLLabel27.Visible:=false;
   RLDBText42.Visible:=false;
   RLLabel53.Visible:=false;
   RLDBText25.Visible:=false;
   RLLabel54.Visible:=false;
   RLDBText39.Visible:=false;
   if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('complemento_cob').AsString)='' then
   begin
    RLLabel40.Visible:=false;
    RLDBText28.Visible:=false;
    RLLabel42.Top  := 144;
    RLDBText29.Top := 144;
    RLLabel41.Top  := 129;
    RLDBText36.Top := 129;
    RLLabel43.Top  := 160;
    RLDBText35.Top := 160;
    RLLabel47.Top  := 177;
    RLDBText37.Top := 177;
    RLLabel62.Top  :=193;
    RLDBText53.Top :=193;
    RLLabel63.Top  :=193;
    RLDBText54.Top :=193;
    RLLabel31.Top  :=209;
    RLDBText41.Top :=209;
    RLLabel64.Top  :=209;
    RLDBText55.Top :=209;
   end
   else
   begin
    RLLabel40.Visible:=true;
    RLDBText28.Visible:=true;
    RLLabel42.Top  := 160;
    RLDBText29.Top := 160;
    RLLabel41.Top  := 144;
    RLDBText36.Top := 144;
    RLLabel43.Top  := 177;
    RLDBText35.Top := 177;
    RLLabel47.Top  := 193;
    RLDBText37.Top := 193;
    RLLabel62.Top  :=209;
    RLDBText53.Top :=209;
    RLLabel63.Top  :=209;
    RLDBText54.Top :=209;
    RLLabel31.Top  :=226;
    RLDBText41.Top :=226;
    RLLabel64.Top  :=226;
    RLDBText55.Top :=226;
   end;

   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('Fone1').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('Fone1').AsString)='(00)0000-0000')  then
   begin
    RLLabel62.Visible:=false;
    RLDBText53.Visible:=false;
    if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)<>'') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)<>'(00)0000-0000')  then
    begin
       RLLabel63.Left:=8;
       RLDBText54.Left:=100;
    end;
   end
   else
   begin
    RLLabel63.Left:=212;
    RLDBText54.Left:=307;
    RLLabel62.Visible:=true;
    RLDBText53.Visible:=true;
   end;

   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').AsString)='(00)0000-0000')  then
   begin
    RLLabel63.Visible:=false;
    RLDBText54.Visible:=false;
   end
   else
   begin
    RLLabel63.Visible:=true;
    RLDBText54.Visible:=true;
   end;
   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone2').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone2').AsString)='(00)0000-0000')  then
   begin
    RLLabel31.Visible:=false;
    RLDBText41.Visible:=false;
    if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)<>'') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)<>'(00)0000-0000')  then
    begin
       RLLabel64.Left:=8;
       RLDBText55.Left:=100;
    end;

   end
   else
   begin
    RLLabel64.Left:=212;
    RLDBText55.Left:=307;
    RLLabel31.Visible:=true;
    RLDBText41.Visible:=true;
   end;

   if (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)='') or (alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').AsString)='(00)0000-0000')  then
   begin
    RLLabel64.Visible:=false;
    RLDBText55.Visible:=false;
   end
   else
   begin
    RLLabel64.Visible:=true;
    RLDBText55.Visible:=true;
   end;
 end;

 if (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='CASADO(A)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='casado(a)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='UNIÃO ESTÁVEL') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='união estável') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='REL.ESTAVEL') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='rel.estavel') then
 begin
    RLLabel28.Visible:=true;
    RLDBText20.Visible:=true;
    RLLabel32.Visible:=true;
    RLDBText22.Visible:=true;
    RLLabel30.Visible:=true;
    RLDBText21.Visible:=true;
    RLLabel52.Visible:=true;
    RLDBText24.Visible:=true;
    RLLabel55.Visible:=true;
    RLDBText44.Visible:=true;
    RLLabel51.Visible:=true;
    RLDBText56.Visible:=true;
    RLLabel29.Visible:=true;
    RLDBText40.Visible:=true;
    RLLabel27.Visible:=true;
    RLDBText42.Visible:=true;
    RLLabel22.Visible:=true;
    RLDBText19.Visible:=true;
    RLLabel53.Visible:=true;
    RLDBText25.Visible:=true;
    RLLabel54.Visible:=true;
    RLDBText39.Visible:=true;

 end
 else
 begin
    RLLabel28.Visible:=false;
    RLDBText20.Visible:=false;
    RLLabel32.Visible:=false;
    RLDBText22.Visible:=false;
    RLLabel30.Visible:=false;
    RLDBText21.Visible:=false;
    RLLabel52.Visible:=false;
    RLDBText24.Visible:=false;
    RLLabel55.Visible:=false;
    RLDBText44.Visible:=false;
    RLLabel22.Visible:=false;
    RLDBText19.Visible:=false;
    RLLabel51.Visible:=false;
    RLDBText56.Visible:=false;
    RLLabel29.Visible:=false;
    RLDBText40.Visible:=false;
    RLLabel27.Visible:=false;
    RLDBText42.Visible:=false;
    RLLabel53.Visible:=false;
    RLDBText25.Visible:=false;
    RLLabel54.Visible:=false;
    RLDBText39.Visible:=false;

 end;

 if ((DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='CASADO(A)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='casado(a)')) and ((DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='Comunhão Universal de Bens') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='COMUNHáO UNIVERSAL DE BENS') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='Comunhão de Bens') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='COMUNHáO DE BENS')) then
 begin
    if empty(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('cart_cid_est').AsString)) then
    begin
      RLDBText56.DataField:='';
      RLDBText56.DataSource:=DM_Tabelas.DS_Compr_Dados;
      RLDBText56.DataField:='cid_uf_cob';
      RLLabel51.Visible:=true;
      RLDBText56.Visible:=true;

    end
    else
    begin
      RLDBText56.DataField:='';
      RLDBText56.DataSource:=DM_Tabelas.DS_Compr_conjuge;
      RLDBText56.DataField:='cart_cid_est';
      RLLabel51.Visible:=true;
      RLDBText56.Visible:=true;

    end;

//    RLLabel51.Visible:=true;
  //  RLDBText56.Visible:=true;

    if empty(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('comar_cid_est').AsString)) then
    begin
//      RLDBText57.DataSource:=DM_Tabelas.DS_Compr_Dados;
//      RLDBText57.DataField:='cid_uf_cob';
      RLLabel50.Visible:=false;
      RLDBText57.Visible:=false;

    end
    else
    begin
//      RLDBText57.DataSource:=DM_Tabelas.DS_Compr_conjuge;
//      RLDBText57.DataField:='comar_cid_est';
      RLLabel50.Visible:=true;
      RLDBText57.Visible:=true;

    end;
{    if empty(charrem(RLDBText57.Caption)) then
    begin
      RLLabel50.Visible:=false;
      RLDBText57.Visible:=false;
    end
    else
    begin
      RLLabel50.Visible:=true;
      RLDBText57.Visible:=true;
    end;}
    RLLabel30.Visible:=true;
    RLDBText21.Visible:=true;
    RLLabel32.Visible:=true;
    RLDBText22.Visible:=true;
    RLLabel52.Visible:=true;
    RLDBText24.Visible:=true;
    RLLabel55.Visible:=true;
    RLDBText44.Visible:=true;

    if (alltrim(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('comar_cid_est').AsString))='') and (alltrim(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('cartorio').AsString))='') and (DM_Tabelas.ZQCompr_conjuge.FieldByName('data').AsDateTime=0) and (alltrim(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('livro').AsString))='') and (alltrim(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString))='') then
    begin
      RLLabel22.Top  := 55;
      RLDBText19.Top := 55;
      RLLabel29.Top  := 72;
      RLDBText40.Top := 72;
      RLLabel53.Top  := 88;
      RLDBText25.Top := 88;
      RLLabel54.Top  := 104;
      RLDBText39.Top := 104;
      RLLabel27.Top  := 121;
      RLDBText42.Top := 122;

   end
   else
   begin
      RLLabel22.Top  := 121;
      RLDBText19.Top := 121;
      RLLabel29.Top  := 140;
      RLDBText40.Top := 140;
      RLLabel53.Top  := 157;
      RLDBText25.Top := 157;
      RLLabel54.Top  := 174;
      RLDBText39.Top := 174;
      RLLabel27.Top  := 192;
      RLDBText42.Top := 193;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cidestnatural').AsString)='' then
   begin
     RLLabel29.visible:=false;
     RLDBText40.visible:=false;
   end
   else
   begin
     RLLabel29.visible:=true;
     RLDBText40.visible:=true;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cargo').AsString)='' then
   begin
     RLLabel27.visible:=false;
     RLDBText42.visible:=false;
   end
   else
   begin
     RLLabel27.visible:=true;
     RLDBText42.visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc1').AsString)='' then
   begin
     RLLabel54.visible:=false;
     RLDBText39.visible:=false;
   end
   else
   begin
     RLLabel54.visible:=true;
     RLDBText39.visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc2').AsString)='' then
   begin
     RLLabel53.visible:=false;
     RLDBText25.visible:=false;
   end
   else
   begin
     RLLabel53.visible:=true;
     RLDBText25.visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString)='' then
   begin
     RLLabel48.Visible:=false;
     RLDBText32.Visible:=false;
   end
   else
   begin
     RLLabel48.Visible:=true;
     RLDBText32.Visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString)='' then
   begin
     RLLabel28.visible:=false;
     RLDBText20.visible:=false;
   end
   else
   begin
     RLLabel28.Visible:=true;
     RLDBText20.Visible:=true;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cartorio').AsString)='' then
   begin
     RLLabel30.visible:=false;
     RLDBText21.visible:=false;
   end
   else
   begin
     RLLabel30.Visible:=true;
     RLDBText21.Visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('livro').AsString)='' then
   begin
     RLLabel32.visible:=false;
     RLDBText22.visible:=false;
     if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)='' then
     begin
       RLLabel52.visible:=false;
       RLDBText24.visible:=false;
     end;
   end
   else
   begin
     RLLabel32.visible:=true;
     RLDBText22.visible:=true;
     if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)<>'' then
     begin
       RLLabel52.visible:=true;
       RLDBText24.visible:=true;
     end;
   end;



   if DM_Tabelas.ZQCompr_conjuge.FieldByName('data').AsDateTime=0 then
   begin
     RLLabel55.visible:=false;
     RLDBText44.visible:=false;
   end
   else
   begin
     RLLabel55.visible:=true;
     RLDBText44.visible:=true;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('nome').AsString)='' then
   begin
     RLLabel22.visible:=false;
     RLDBText19.visible:=false;
   end
   else
   begin
     RLLabel22.visible:=true;
     RLDBText19.visible:=true;
   end;

 end

 else if ((DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='CASADO(A)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='casado(a)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='REL.ESTAVEL') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString='rel.estavel')) and ((DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='COMUNHáO PARCIAL DE BENS') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='SEPARAÇÃO DE BENS') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString='UNIÃO ESTÁVEL')) then
 begin
    if empty(RLDBText56.Caption) then
    begin
      RLDBText56.DataField:='';
      RLDBText56.DataSource:=DM_Tabelas.DS_Compr_Dados;
      RLDBText56.DataField:='cid_uf_cob';
    end
    else
    begin
      RLDBText56.DataField:='';    
      RLDBText56.DataSource:=DM_Tabelas.DS_Compr_conjuge;
      RLDBText56.DataField:='cart_cid_est';
    end;

    RLLabel51.Visible:=true;
    RLDBText56.Visible:=true;

    if (alltrim(charrem(DM_Tabelas.ZQCompr_conjuge.FieldByName('comar_cid_est').AsString))='') and (alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cartorio').AsString)='') and (DM_Tabelas.ZQCompr_conjuge.FieldByName('data').AsDateTime=0) and (alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('livro').AsString)='') and (alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)='') then
    begin
      RLLabel22.Top  := 55;
      RLDBText19.Top := 55;
      RLLabel29.Top  := 72;
      RLDBText40.Top := 72;
      RLLabel53.Top  := 88;
      RLDBText25.Top := 88;
      RLLabel54.Top  := 104;
      RLDBText39.Top := 104;
      RLLabel27.Top  := 121;
      RLDBText42.Top := 122;

   end
   else
   begin
      RLLabel22.Top  := 121;
      RLDBText19.Top := 121;
      RLLabel29.Top  := 140;
      RLDBText40.Top := 140;
      RLLabel53.Top  := 157;
      RLDBText25.Top := 157;
      RLLabel54.Top  := 174;
      RLDBText39.Top := 174;
      RLLabel27.Top  := 192;
      RLDBText42.Top := 193;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cidestnatural').AsString)='' then
   begin
     RLLabel29.visible:=false;
     RLDBText40.visible:=false;
   end
   else
   begin
     RLLabel29.visible:=true;
     RLDBText40.visible:=true;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cargo').AsString)='' then
   begin
     RLLabel27.visible:=false;
     RLDBText42.visible:=false;
   end
   else
   begin
     RLLabel27.visible:=true;
     RLDBText42.visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc1').AsString)='' then
   begin
     RLLabel54.visible:=false;
     RLDBText39.visible:=false;
   end
   else
   begin
     RLLabel54.visible:=true;
     RLDBText39.visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('doc2').AsString)='' then
   begin
     RLLabel53.visible:=false;
     RLDBText25.visible:=false;
   end
   else
   begin
     RLLabel53.visible:=true;
     RLDBText25.visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').AsString)='' then
   begin
     RLLabel48.Visible:=false;
     RLDBText32.Visible:=false;
   end
   else
   begin
     RLLabel48.Visible:=true;
     RLDBText32.Visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').AsString)='' then
   begin
     RLLabel28.visible:=false;
     RLDBText20.visible:=false;
   end
   else
   begin
     RLLabel28.Visible:=true;
     RLDBText20.Visible:=true;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('cartorio').AsString)='' then
   begin
     RLLabel30.visible:=false;
     RLDBText21.visible:=false;
   end
   else
   begin
     RLLabel30.Visible:=true;
     RLDBText21.Visible:=true;
   end;


   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('livro').AsString)='' then
   begin
     RLLabel32.visible:=false;
     RLDBText22.visible:=false;
     if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)='' then
     begin
       RLLabel52.visible:=false;
       RLDBText24.visible:=false;
     end;
   end
   else
   begin
     RLLabel32.visible:=true;
     RLDBText22.visible:=true;
     if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('folha').AsString)<>'' then
     begin
       RLLabel52.visible:=true;
       RLDBText24.visible:=true;
     end;
   end;



   if DM_Tabelas.ZQCompr_conjuge.FieldByName('data').AsDateTime=0 then
   begin
     RLLabel55.visible:=false;
     RLDBText44.visible:=false;
   end
   else
   begin
     RLLabel55.visible:=true;
     RLDBText44.visible:=true;
   end;

   if alltrim(DM_Tabelas.ZQCompr_conjuge.FieldByName('nome').AsString)='' then
   begin
     RLLabel22.visible:=false;
     RLDBText19.visible:=false;
   end
   else
   begin
     RLLabel22.visible:=true;
     RLDBText19.visible:=true;
   end;

 end




 else
 begin
    RLLabel51.Visible:=false;
    RLDBText56.Visible:=false;
    RLLabel50.Visible:=false;
    RLDBText57.Visible:=false;
    RLLabel30.Visible:=false;
    RLDBText21.Visible:=false;
    RLLabel32.Visible:=false;
    RLDBText22.Visible:=false;
    RLLabel52.Visible:=false;
    RLDBText24.Visible:=false;
    RLLabel55.Visible:=false;
    RLDBText44.Visible:=false;

    RLLabel29.visible:=false;
    RLDBText40.visible:=false;
    RLLabel27.visible:=false;
    RLDBText42.visible:=false;
    RLLabel54.visible:=false;
    RLDBText39.visible:=false;
    RLLabel53.visible:=false;
    RLDBText25.visible:=false;
    RLLabel28.visible:=false;
    RLDBText20.visible:=false;
    RLLabel30.visible:=false;
    RLDBText21.visible:=false;
    RLLabel32.visible:=false;
    RLDBText22.visible:=false;
    RLLabel52.visible:=false;
    RLDBText24.visible:=false;
    RLLabel55.visible:=false;
    RLDBText44.visible:=false;
    RLLabel22.visible:=false;
    RLDBText19.visible:=false;
 end;
end;

procedure TFrm_QuadroResumo2.RLReport1PageEnding(Sender: TObject);
begin
  Frm_QuadroResumo2.npg.Value:=Frm_QuadroResumo2.npg.Value+1;
end;

procedure TFrm_QuadroResumo2.RLReport1PageStarting(Sender: TObject);
begin
  pagina.Caption:=floattostr(Frm_QuadroResumo2.npg.value);
end;

procedure TFrm_QuadroResumo2.RLDetailGrid2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if DM_Tabelas.CDSAss1assina.Value <> 'S' Then
  begin
    RLDetailGrid2.ColCount:=1;
    RLLabel6.Caption:='___________________________________________________________________________';
    RLLabel6.Left:=107;
    RLLabel6.Top:=30;
    RLLabel6.Width:=529;
    RLDBText50.Left:=350;
    RLDBText50.Top:=48;
    RLDBText50.Width:=48;
    RLDBText51.Left:=359;
    RLDBText51.Top:=60;
    RLDBText51.Width:=30;
    RLDBText52.Left:=343;
    RLDBText52.Top:=75;
    RLDBText52.Width:=62;
    RLLabel5.Visible:=false;
    RLDBText47.Visible:=false;
    RLDBText45.Visible:=false;
    RLDBText46.Visible:=false;
  end;

  if DM_Tabelas.CDSAss1assina.Value = 'S' Then
  begin
    RLDetailGrid2.ColCount:=1;
    RLLabel6.Caption:='___________________________________________________';
    RLLabel6.Left:=5;
    RLLabel6.Top:=30;
    RLLabel6.Width:=361;
    RLDBText50.Left:=160;
    RLDBText50.Top:=48;
    RLDBText50.Width:=48;
    RLDBText51.Left:=169;
    RLDBText51.Top:=60;
    RLDBText51.Width:=30;
    RLDBText52.Left:=153;
    RLDBText52.Top:=75;
    RLDBText52.Width:=62;
    RLLabel5.Visible:=true;
    RLDBText47.Visible:=true;
    RLDBText45.Visible:=true;
    RLDBText46.Visible:=true;
  end;
  
end;


procedure TFrm_QuadroResumo2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_QuadroResumo2, 'recebimento', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento', 'recebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_QuadroResumo2, 'cheque', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequebanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequedono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequenumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequeSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'cheque', 'chequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_QuadroResumo2, 'recebimento2', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo2, 'recebimento2', 'recebimento2vr', 'vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
