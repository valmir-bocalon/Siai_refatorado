unit acertaproximo_reajuste_em_recebimento;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Gauges, StdCtrls, DBClient, Grids, DBGrids;

type
  TFrm_acertaproximo_reajuste_em_recebimento = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButtonArround;
    dxButton1: TdxButtonArround;
    Label1: TLabel;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;


    recebetemp: TClientDataSet;
    Datarecebetemp: TDataSource;
    recebetempidrecebimento: TLargeintField;
    recebetempdocumento: TWideStringField;
    recebetempcliente: TIntegerField;
    recebetempusuario: TIntegerField;
    recebetempDt_Entrada: TDateField;
    recebetempDt_Vencimento: TDateField;
    recebetempValor: TFloatField;
    recebetempObserv: TMemoField;
    recebetempVrDoc: TFloatField;
    recebetempordem: TWideStringField;
    recebetempTipDoc: TWideStringField;
    recebetempsaldo: TFloatField;
    recebetempmarcar: TWideStringField;
    recebetempRefBaixa: TIntegerField;
    recebetemprefvinda: TIntegerField;
    recebetempcontabil: TIntegerField;
    recebetempempresa: TIntegerField;
    recebetempcustodaparcela: TFloatField;
    recebetemporigem: TWideStringField;
    recebetempadversa: TIntegerField;
    recebetemprecpag: TWideStringField;
    recebetempnumordem: TIntegerField;
    recebetempidloteamento: TIntegerField;
    recebetempvenda_idvenda: TIntegerField;
    recebetempquadralote: TWideStringField;
    recebetempnumboleto: TWideStringField;
    recebetempsq: TLargeintField;
    recebetempReajustado: TWideStringField;
    recebetempData_reajuste: TDateField;
    recebetempsomar: TWideStringField;
    recebetempProximo_Reajuste: TWideStringField;
    DBGrid1: TDBGrid;
    Label2: TLabel;
    dxButton2: TdxButtonArround;
    ZQEntrada: TZQuery;































    DS_Entrada: TDataSource;
    ZQVenda: TZQuery;
    DS_Venda: TDataSource;
    DS_Parcela: TDataSource;
    ZQParcela: TZQuery;





























    ZQEmpree: TZQuery;
    DS_Empree: TDataSource;










































































    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dxButton2Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_acertaproximo_reajuste_em_recebimento: TFrm_acertaproximo_reajuste_em_recebimento;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_acertaproximo_reajuste_em_recebimento.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_acertaproximo_reajuste_em_recebimento.BTFecharClick(Sender: TObject);
var
//data:Tdatetime;
data,identificou,encontrou,seq,ql,passou:string;
ano,posi,nm:integer;
begin
  passou:='F';
  Label2.Visible:=true;
  application.ProcessMessages;
  recebetemp.close;
  recebetemp.CreateDataSet;
  DM_TAbelas.ZQRecebimento.sql.clear;
  DM_TAbelas.ZQRecebimento.sql.add('Select * from recebimento where saldo > 0 and TipDoc=''BO''  group by quadralote order by quadralote');
//  DM_TAbelas.ZQRecebimento.sql.add('Select * from recebimento where  TipDoc=''BO''  group by quadralote order by quadralote');
  DM_TAbelas.ZQRecebimento.Open;
  DM_TAbelas.ZQRecebimento.First;
  Gauge1.MaxValue := DM_Tabelas.ZQRecebimento.RecordCount;
  while not DM_TAbelas.ZQRecebimento.Eof do
  begin
    Gauge1.Progress := DM_Tabelas.ZQRecebimento.RecNo;
    posi:=pos('-',DM_TAbelas.ZQRecebimento.FieldByName('ordem').AsString);
    seq:=copy(DM_TAbelas.ZQRecebimento.FieldByName('ordem').AsString,posi,10);

    ZQRecebimento.Close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add('SELECT * FROM recebimento where quadralote='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('quadralote').AsString)+' and  ordem between '+quotedstr(DM_TAbelas.ZQRecebimento.FieldByName('numordem').Text+seq)+' and '+quotedstr(DM_TAbelas.ZQRecebimento.FieldByName('numordem').Text+'-240/240')+' order by ordem');
    ZQRecebimento.Open;
    ZQRecebimento.First;
    if ZQRecebimento.RecordCount>0 then
    begin
      while not ZQRecebimento.Eof do
      begin
        posi:=pos('-',ZQRecebimento.FieldByName('ordem').AsString);
        inc(posi);
        seq:=copy(ZQRecebimento.FieldByName('ordem').AsString,posi,3);
        recebetemp.Insert;
        recebetempidrecebimento.Value:=ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
        recebetempdocumento.Value:=ZQRecebimento.FieldByName('documento').AsString;
        recebetempcliente.Value:=ZQRecebimento.FieldByName('cliente').AsInteger;
        recebetempDt_Entrada.Value:=ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime;
        recebetempDt_Vencimento.Value:=ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
        recebetempordem.Value:=ZQRecebimento.FieldByName('ordem').AsString;
        recebetempTipDoc.Value:=ZQRecebimento.FieldByName('TipDoc').AsString;
        recebetempsaldo.Value:=ZQRecebimento.FieldByName('saldo').AsFloat;
        recebetempnumordem.Value:=ZQRecebimento.FieldByName('numordem').AsInteger;
        recebetempquadralote.value:=ZQRecebimento.FieldByName('quadralote').AsString;
        recebetempidloteamento.value:=ZQRecebimento.FieldByName('idloteamento').AsInteger;
        recebetempsq.Value:=strtoint(seq);
        recebetemp.Post;
        ZQRecebimento.Edit;
        ZQRecebimento.FieldByName('Proximo_Reajuste').Clear;
        ZQRecebimento.post;
        ZQRecebimento.Next;
        application.ProcessMessages;
      end;
    end;
    DM_Tabelas.ZQRecebimento.Next;
  end;
  Gauge1.Progress :=0;
  recebetemp.IndexName:='recebetempIndex1';
  recebetemp.first;
  ql:=recebetempquadralote.value;
  posi:=recebetempsq.value;
  nm:=recebetempsq.Value;
//  data:=IncMonth(recebetempDt_Vencimento.Value, 12);
  data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
  Gauge1.MaxValue := Recebetemp.RecordCount;
  encontrou:='F';
  Label2.caption:='Aguarde! Verificando as Parcelas Selecionadas';
  application.ProcessMessages;
  identificou:='N';
  while not recebetemp.eof do
  begin
    Gauge1.Progress := Recebetemp.RecNo;
    if ql=recebetempquadralote.value then
    begin
{      if (nm=1) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;}
      if (nm=13) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=25) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=37)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=49)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=61) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=73) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=85) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=97) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=109) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=121)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=133)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;

      if (nm=145) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=157)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=169) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=181) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=193) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=205)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=217) and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm=229)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
      if (nm>=241)  and (identificou='N') then
      begin
        data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
        identificou:='S';
      end;
//      recebetempProximo_Reajuste.value:=copy(datetostr(data),4,7);
      if not empty(data) then
      begin
        recebetemp.edit;
        recebetempProximo_Reajuste.value:=data;
        recebetemp.post;
      end;
      recebetemp.Next;
      nm:=recebetempsq.Value;
    end
    else
    begin
      ql:=recebetempquadralote.value;
      posi:=recebetempsq.value;
      nm:=recebetempsq.value;
//      data:=IncMonth(recebetempDt_Vencimento.Value, 12);
      data:=copy(datetostr(recebetempDt_Vencimento.Value),4,7);
      identificou:='N';
    end;
  end;
  Gauge1.Progress :=0;
  Label2.Visible:=false;
  application.ProcessMessages;
//  dxButton2.Enabled:=true;

  // botao processar
  Label2.Caption:='Aguarde Até o Final! Não Desligue a Máquina!';
  Label2.Visible:=true;
  application.ProcessMessages;
  Recebetemp.recordcount;
  Recebetemp.first;
  Gauge1.Progress :=0;
  Gauge1.MaxValue := Recebetemp.RecordCount;
  while not Recebetemp.eof do
  begin
    Gauge1.Progress := Recebetemp.RecNo;
    DM_TAbelas.ZQRecebimento.sql.clear;
    DM_TAbelas.ZQRecebimento.sql.add('Select * from recebimento where idrecebimento='+quotedstr(recebetempidrecebimento.Text));
    DM_TAbelas.ZQRecebimento.Open;
    DM_TAbelas.ZQRecebimento.First;
    while not DM_TAbelas.ZQRecebimento.Eof do
    begin
      DM_TAbelas.ZQRecebimento.Edit;
      DM_TAbelas.ZQRecebimento.FieldByName('Proximo_Reajuste').AsString:=recebetempProximo_Reajuste.Value;
      DM_Tabelas.ZQRecebimento.Post;
      DM_Tabelas.ZQRecebimento.next;
    end;
    Recebetemp.Next;
  end;
  Label2.Caption:='Fim de Processo...';
  Label2.Visible:=false;
  application.ProcessMessages;
  close;

end;

procedure TFrm_acertaproximo_reajuste_em_recebimento.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQRecebimento.close;
  ZQRecebimento.Close;
  recebetemp.close;
  ZQEntrada.close;
  ZQParcela.Close;
  ZQVenda.Close;
end;

procedure TFrm_acertaproximo_reajuste_em_recebimento.dxButton2Click(
  Sender: TObject);
begin

{  Label2.Caption:='Aguarde Até o Final! Não Desligue a Máquina!';
  Label2.Visible:=true;
  application.ProcessMessages;
  Recebetemp.recordcount;
  Recebetemp.first;
  Gauge1.Progress :=0;
  Gauge1.MaxValue := Recebetemp.RecordCount;
  while not Recebetemp.eof do
  begin
    Gauge1.Progress := Recebetemp.RecNo;
    DM_TAbelas.ZQRecebimento.sql.clear;
    DM_TAbelas.ZQRecebimento.sql.add('Select * from recebimento where idrecebimento='+quotedstr(recebetempidrecebimento.Text));
    DM_TAbelas.ZQRecebimento.Open;
    DM_TAbelas.ZQRecebimento.First;
    while not DM_TAbelas.ZQRecebimento.Eof do
    begin
      DM_TAbelas.ZQRecebimento.Edit;
      DM_TAbelas.ZQRecebimentoProximo_Reajuste.Value:=recebetempProximo_Reajuste.Value;
      DM_Tabelas.ZQRecebimento.Post;
      DM_Tabelas.ZQRecebimento.next;
    end;
    Recebetemp.Next;
  end;}
  Label2.Caption:='Fim de Processo...';
  Label2.Visible:=false;
  application.ProcessMessages;
  close;
end;


procedure TFrm_acertaproximo_reajuste_em_recebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeDataSet(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQVenda', 'ZQVendamarca', 'marca', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreelote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreelado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreematri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaproximo_reajuste_em_recebimento, 'ZQEmpree', 'ZQEmpreeproposta', 'proposta', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.



