unit QuadroResumo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, StdCtrls, FnpNumericEdit;

type
  TFrm_QuadroResumo = class(TForm)
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
    RLLabel24: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel26: TRLLabel;
    RLDBText15: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText18: TRLDBText;
    RLSubDetail5: TRLSubDetail;
    RLBand7: TRLBand;
    RLLabel22: TRLLabel;
    RLDBText19: TRLDBText;
    RLLabel30: TRLLabel;
    RLLabel32: TRLLabel;
    RLLabel50: TRLLabel;
    RLLabel51: TRLLabel;
    RLDBText21: TRLDBText;
    RLDBText22: TRLDBText;
    RLLabel52: TRLLabel;
    RLDBText24: TRLDBText;
    RLLabel53: TRLLabel;
    RLDBText25: TRLDBText;
    RLLabel54: TRLLabel;
    RLDBText39: TRLDBText;
    RLLabel29: TRLLabel;
    RLDBText40: TRLDBText;
    RLLabel27: TRLLabel;
    RLDBText42: TRLDBText;
    RLDBText23: TRLDBText;
    RLDBText43: TRLDBText;
    RLLabel55: TRLLabel;
    RLDBText44: TRLDBText;
    RLBand10: TRLBand;
    RLSubDetail6: TRLSubDetail;
    RLSubDetail7: TRLSubDetail;
    RLBand4: TRLBand;
    RLLabel11: TRLLabel;
    RLMemo1: TRLMemo;
    RLSubDetail8: TRLSubDetail;
    RLBand8: TRLBand;
    RLLabel59: TRLLabel;
    RLDetailGrid1: TRLDetailGrid;
    RLDBText49: TRLDBText;
    RLLabel5: TRLLabel;
    RLDBText48: TRLDBText;
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
    RLSubDetail9: TRLSubDetail;
    RLDetailGrid2: TRLDetailGrid;
    RLDBText50: TRLDBText;
    RLLabel6: TRLLabel;
    RLDBText51: TRLDBText;
    RLDBText52: TRLDBText;
    RLSubDetail10: TRLSubDetail;
    RLBand61: TRLBand;
    RLLabel61: TRLLabel;
    RLLabel62: TRLLabel;
    RLDBText53: TRLDBText;
    RLLabel31: TRLLabel;
    RLDBText41: TRLDBText;
    RLLabel63: TRLLabel;
    RLLabel64: TRLLabel;
    RLDBText54: TRLDBText;
    RLDBText55: TRLDBText;
    RLLabel65: TRLLabel;
    RLLabel66: TRLLabel;
    RLLabel67: TRLLabel;
    RLBand9: TRLBand;
    RLDraw1: TRLDraw;
    RLMemo2: TRLMemo;
    recebimento: TZQuery;
    DataSource1: TDataSource;





























    cheque: TZQuery;
    DataSource2: TDataSource;

















    RLDBText56: TRLDBText;
    RLDBText57: TRLDBText;
    RLDraw2: TRLDraw;
    RLMemo3: TRLMemo;
    RLDraw3: TRLDraw;
    RLLabel57: TRLMemo;
    RLLabel28: TRLLabel;
    RLDBText20: TRLDBText;
    pagina: TRLLabel;
    RLDraw4: TRLDraw;
    npg: TFnpNumericEdit;
    RLLabel56: TRLLabel;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLSubDetail10BeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand7BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1PageEnding(Sender: TObject);
    procedure RLReport1PageStarting(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_QuadroResumo: TFrm_QuadroResumo;

implementation

uses tabelas, funcoes, Venda, uRuntimeFields;

{$R *.dfm}

procedure TFrm_QuadroResumo.RLBand1BeforePrint(Sender: TObject;
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
  RLLabel61.Caption := DM_Tabelas.ZQVenda.FieldByName('Cidade').Value+' / '+DM_Tabelas.ZQVenda.FieldByName('uf').Value+', '+dataporextenso(DM_Tabelas.ZQVenda.FieldByName('datavenda').Value);
  RLLabel66.Caption := numeroporextenso(nm);

end;

procedure TFrm_QuadroResumo.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
  empresa,varlinha : string;
  vrrec,entrada:double;
  posi,vez,tvez:integer;
begin
  entrada:=0;
  vrrec:=0;
  vez:=0;
  tvez:=0;
  varlinha := '';
  RLMemo1.Lines.Clear;
  RLMemo1.Lines.Add('Valor Total: R$ '+alltrim(transform(DM_Tabelas.ZQVenda.FieldByName('valorvenda').value,'###,###,###,##0.00'))+' ('+extenso(DM_Tabelas.ZQVenda.FieldByName('valorvenda').Value)+'). ');

  if Frm_Venda.XNEEntrada.Value>0 Then Begin
    varlinha := 'Entrada de R$ '+alltrim(transform(Frm_Venda.XNEEntrada.Value,'###,###,###,##0.00')) +' ('+extenso(Frm_Venda.XNEEntrada.Value)+'), ';
    if Frm_Venda.XNQtdeEnt.Value>1 Then
      varlinha := varlinha + ' em '+Frm_Venda.XNQtdeEnt.Text+' vezes '+Frm_Venda.CBPeriodoEnt.Text;
    varlinha := varlinha +' da seguinte forma:';
    RLMemo2.Lines.Clear;
    RLMemo2.Lines.Add(varlinha);
    DM_TAbelas.CDSPagEntrada.First;
    varlinha := '';
    while not DM_Tabelas.CDSPagEntrada.Eof do begin
      varlinha := varlinha + DM_Tabelas.CDSPagEntradaordem.Value+'  '+DM_Tabelas.CDSPagEntradavencimento.Text+'  '+alltrim(transform(DM_Tabelas.CDSPagEntradavalor.Value,'###,###,###,##0.00'));
      varlinha := varlinha + ' Tipo '+DM_Tabelas.CDSPagEntradatipodoc.Value;
      if not empty(DM_Tabelas.CDSPagEntradabanco.Value) Then Begin
        varlinha := varlinha + ' Bc.'+DM_Tabelas.CDSPagEntradabanco.Value;
        varlinha := varlinha + ' Ag.'+DM_Tabelas.CDSPagEntradaagencia.Value;
        varlinha := varlinha + ' Cc.'+DM_Tabelas.CDSPagEntradaconta.Value;
        varlinha := varlinha + ' n� '+DM_Tabelas.CDSPagEntradanumero.Value;
      end;
      DM_Tabelas.CDSPagEntrada.Next;
      IF not DM_Tabelas.CDSPagEntrada.Eof Then
    //    varlinha := varlinha + ' | ';
        varlinha := varlinha+chr(10)+chr(13);    
    end;
    RLMemo2.Lines.Add(varlinha);
  end
  else
  begin
    if not DM_Tabelas.ZQvenda.Active then
       DM_Tabelas.ZQvenda.Active:=true;
    recebimento.Close;
    recebimento.SQL.Clear;
    recebimento.SQL.Add('select * from recebimento where documento like '+quotedstr('%-E-%'));
    recebimento.Open;
    recebimento.First;
    cheque.open;
    if recebimento.RecordCount>0 then
    begin
      while not recebimento.Eof do
      begin
        entrada:=entrada+recebimento.FieldByName('Valor').Value;
        inc(vez);
        recebimento.Next;
      end;
      varlinha := 'Entrada de R$ '+alltrim(transform(entrada,'###,###,###,##0.00')) +' ('+extenso(entrada)+'), ';
      posi:=pos('E',recebimento.FieldByName('documento').Value);
      Inc(posi,2);
      varlinha := varlinha + ' em '+inttostr(vez)+' veze(s) '+copy(recebimento.FieldByName('documento').Value,posi,20);
      varlinha := varlinha +' da seguinte forma:';

      RLMemo2.Lines.Clear;
      RLMemo2.Lines.Add(varlinha);
      recebimento.First;
      varlinha := '|';
      while not recebimento.Eof do begin
        posi:=pos('-',recebimento.FieldByName('ordem').Value);
        Inc(posi);
        varlinha := varlinha + alltrim(copy(recebimento.FieldByName('ordem').Value,posi,20))+'|'+recebimento.FieldByName('Dt_Vencimento').Text +'|'+alltrim(transform(recebimento.FieldByName('Valor').value,'###,###,###,##0.00')+'|');
        varlinha := varlinha + ' Tipo '+recebimento.FieldByName('TipDoc').Value;
        if cheque.RecordCount>0 then
        begin
          varlinha := varlinha + ' Bc.'+cheque.FieldByName('banco').Value;
          varlinha := varlinha + ' Ag.'+cheque.FieldByName('agencia').Value;
          varlinha := varlinha + ' Cc.'+cheque.FieldByName('conta').Value;
          varlinha := varlinha + ' n� '+cheque.FieldByName('numero').Value;
        end;
        recebimento.Next;
        IF not recebimento.Eof Then
         // varlinha := varlinha + ' | ';
           varlinha := varlinha+chr(10)+chr(13);
      end;
      RLMemo2.Lines.Add(varlinha);
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
    recebimento.SQL.Add('select * from recebimento where documento like '+quotedstr('%-P-%'));
    recebimento.Open;
    recebimento.First;
    cheque.open;
    vez:=0;
    entrada:=0;
    varlinha := '';
    if recebimento.RecordCount>0 then
    begin
      while not recebimento.Eof do
      begin
        entrada:=entrada+recebimento.FieldByName('Valor').Value;
        inc(vez);
        recebimento.Next;
      end;
      tvez:=vez;
      recebimento.First;
      vrrec:=recebimento.FieldByName('Valor').Value;

      posi:=pos('/',recebimento.FieldByName('ordem').Value);
      inc(posi);
      if posi>0 then
      begin
         tvez:=strtoint(alltrim(copy(recebimento.FieldByName('ordem').Value,posi,5)));
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
      varlinha := varlinha + ' de R$ '+alltrim(transform(vrrec,'###,###,###,##0.00'))+',  com o primeiro vencimento em '+DM_Tabelas.CDSPagParcelavencimento.Text;
    RLMemo3.Lines.Add(varlinha);
  end
  else
  begin
    if not DM_Tabelas.ZQvenda.Active then
       DM_Tabelas.ZQvenda.Active:=true;
    recebimento.Close;
    recebimento.SQL.Clear;
    recebimento.SQL.Add('select * from recebimento where documento like '+quotedstr('%-P-%'));
    recebimento.Open;
    recebimento.First;
    cheque.open;
    vez:=0;
    entrada:=0;
    varlinha := '';
    if recebimento.RecordCount>0 then
    begin
      while not recebimento.Eof do
      begin
        entrada:=entrada+recebimento.FieldByName('Valor').Value;
        inc(vez);
        recebimento.Next;
      end;
      tvez:=vez;
      recebimento.First;
      vrrec:=recebimento.FieldByName('Valor').Value;

      posi:=pos('/',recebimento.FieldByName('ordem').Value);
      inc(posi);
      if posi>0 then
      begin
         tvez:=strtoint(alltrim(copy(recebimento.FieldByName('ordem').Value,posi,5)));
         if vez>tvez then
         begin
            vez:=vez div tvez;
            vrrec:=vrrec*vez;
         end;
      end;

      posi:=0;
//      DM_Tabelas.CDSPagParcela.First;
      varlinha := 'O restante de R$ '+alltrim(transform(entrada,'###,###,###,##0.00')) +' ('+extenso(entrada)+'), ';
      posi:=pos('P',recebimento.FieldByName('documento').Value);
      Inc(posi,2);
      varlinha := varlinha + ' em '+inttostr(tvez)+' veze(s) '+copy(recebimento.FieldByName('documento').Value,posi,20);
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
          varlinha := varlinha + ' n� '+chequenumero.Value;
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

end;

procedure TFrm_QuadroResumo.RLSubDetail10BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFrm_QuadroResumo.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
texto:string;  
begin
  if (not empty(DM_Tabelas.ZQVenda.FieldByName('linha').Value)) and (DM_Tabelas.ZQVenda.FieldByName('linha').Value<>'0') then
  begin
    texto:='Im�vel distante '+alltrim(DM_Tabelas.ZQVenda.FieldByName('linha').Value)+' m em linha reta e '+alltrim(DM_Tabelas.ZQVenda.FieldByName('curva').Value)+' m em curva da '+alltrim(DM_Tabelas.ZQVenda.FieldByName('esquina').Value);
    RLLabel57.Lines.Clear;
    RLLabel57.Lines.Add(texto);
    RLLabel57.Visible:=true;
  end
  else
  begin
    RLLabel57.Lines.Clear;
    RLLabel57.Visible:=false;
  end;
  if empty(DM_Tabelas.ZQVenda.FieldByName('matricula').Value) then
  begin
    RLDBText6.DataField:='imovel_matr';
  end
  else
  begin
    RLDBText6.DataField:='matricula';
  end;
  RLLabel56.Caption:=alltrim(DM_Tabelas.ZQVenda.FieldByName('forma_reajuste').Value);
end;

procedure TFrm_QuadroResumo.RLBand7BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').Value='Comunh�o Universal de Bens') or (DM_Tabelas.ZQCompr_conjuge.FieldByName('regime').Value='COMUNH�O UNIVERSAL DE BENS') then
  begin
    RLLabel51.Visible:=true;
    RLDBText56.Visible:=true;
    RLLabel50.Visible:=true;
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
end;

procedure TFrm_QuadroResumo.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').Value='CASADO(A)') or (DM_Tabelas.ZQCompr_Dados.FieldByName('estadocivil').Value='casado(a)') then
  begin
    RLLabel28.Visible:=true;
    RLDBText20.Visible:=true;
  end
  else
  begin
    RLLabel28.Visible:=false;
    RLDBText20.Visible:=false;
  end;
  if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('complemento_cob').Value)='' then
  begin
    RLLabel40.Visible:=false;
    RLDBText28.Visible:=false;
  end
  else
  begin
    RLLabel40.Visible:=true;
    RLDBText28.Visible:=true;
  end;

  if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('Fone1').Value)='' then
  begin
    RLLabel62.Visible:=false;
    RLDBText53.Visible:=false;
  end
  else
  begin
    RLLabel62.Visible:=true;
    RLDBText53.Visible:=true;
  end;

  if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone3').Value)='' then
  begin
    RLLabel63.Visible:=false;
    RLDBText54.Visible:=false;
  end
  else
  begin
    RLLabel63.Visible:=true;
    RLDBText54.Visible:=true;
  end;
  if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone2').Value)='' then
  begin
    RLLabel31.Visible:=false;
    RLDBText41.Visible:=false;
  end
  else
  begin
    RLLabel31.Visible:=true;
    RLDBText41.Visible:=true;
  end;

  if alltrim(DM_Tabelas.ZQCompr_Dados.FieldByName('fone4').Value)='' then
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

procedure TFrm_QuadroResumo.RLReport1PageEnding(Sender: TObject);
begin
  Frm_QuadroResumo.npg.Value:=Frm_QuadroResumo.npg.Value+1;
end;

procedure TFrm_QuadroResumo.RLReport1PageStarting(Sender: TObject);
begin
  pagina.Caption:=floattostr(Frm_QuadroResumo.npg.value);
end;


procedure TFrm_QuadroResumo.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_QuadroResumo, 'recebimento', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentonumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'recebimento', 'recebimentosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_QuadroResumo, 'cheque', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequebanco', 'banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeagencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequedono', 'dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeCPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequealias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeconta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequenumero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequeSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_QuadroResumo, 'cheque', 'chequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
