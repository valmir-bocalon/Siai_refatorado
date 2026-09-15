unit acertavalor_final_venda;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Gauges, StdCtrls;

type
  TFrm_acertavalor_final_venda = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButtonArround;
    dxButton1: TdxButtonArround;
    Label1: TLabel;
    ZQRecebimento: TZQuery;
    DS_Recebimento: TDataSource;
    ZQVenda: TZQuery;


    DS_Venda: TDataSource;
    ZQParcela: TZQuery;
    DS_Parcela: TDataSource;
    ZQEntrada: TZQuery;
    DS_Entrada: TDataSource;
    ZQEmpree: TZQuery;






















    DS_Empree: TDataSource;
    ZQtira_negativo_do_valor_recebimento: TZQuery;
    DataSource1: TDataSource;





































































    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_acertavalor_final_venda: TFrm_acertavalor_final_venda;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_acertavalor_final_venda.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_acertavalor_final_venda.BTFecharClick(Sender: TObject);
begin
  ZQtira_negativo_do_valor_recebimento.Close;
  ZQtira_negativo_do_valor_recebimento.SQL.Clear;
  ZQtira_negativo_do_valor_recebimento.SQL.Add('update recebimento set valor=abs(valor) where valor <0');
  ZQtira_negativo_do_valor_recebimento.ExecSQL;
  ZQtira_negativo_do_valor_recebimento.close;
  Gauge1.Progress:=0;
  ZQRecebimento.sql.clear;
  ZQRecebimento.sql.add('Select idrecebimento,documento,valor,saldo,quadralote,venda_idvenda,ordem,cliente,dt_vencimento from recebimento r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) group by r.quadralote order by r.quadralote');
  ZQRecebimento.Open;
  ZQRecebimento.First;
  Gauge1.MaxValue := ZQRecebimento.RecordCount;
  ZQVenda.open;
  ZQRecebimento.DisableControls;
  while not ZQRecebimento.Eof do
  begin
    Gauge1.Progress := ZQRecebimento.RecNo;
    ZQEntrada.Close;
    ZQEntrada.SQL.Clear;
    ZQEntrada.SQL.Add('Select idrecebimento,documento,valor,saldo,quadralote,venda_idvenda,ordem,cliente,dt_vencimento, sum(valor) as entrada from recebimento r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.quadralote='+quotedstr(ZQRecebimento.FieldByName('quadralote').AsString)+' and r.documento like  ''%-E-%''  order by r.ordem');
    ZQEntrada.open;

    ZQParcela.Close;
    ZQParcela.SQL.Clear;
    ZQParcela.SQL.Add('Select idrecebimento,documento,valor,saldo,quadralote,venda_idvenda,ordem,cliente,dt_vencimento, sum(valor) as parcela from recebimento r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.quadralote='+quotedstr(ZQRecebimento.FieldByName('quadralote').AsString)+' and r.documento like  ''%-P-%''  order by r.ordem');
    ZQParcela.open;
    if (ZQEntrada.FieldByName('entrada').AsFloat>0) or (ZQParcela.FieldByName('parcela').AsFloat>0) then
    begin
      ZQVenda.SQL.Clear;
      ZQVenda.SQL.Add('Select idvenda,valorvenda from venda where idvenda='+quotedstr(ZQRecebimento.FieldByName('venda_idvenda').Text));
      ZQVenda.open;
      if ZQVenda.RecordCount>0 then
      begin
        ZQVenda.Edit;
        ZQVenda.FieldByName('valorvenda').AsFloat:=ZQEntrada.FieldByName('entrada').AsFloat+ZQParcela.FieldByName('parcela').AsFloat;
        ZQVenda.Post;
      end;
    end;
    ZQRecebimento.Next;
  end;
  ZQRecebimento.EnableControls;
  gauge1.Progress:=0;
  ZQEmpree.close;
  ZQVenda.close;
  ZQRecebimento.close;
  ZQEntrada.close;
  ZQParcela.close;
  close;
end;


procedure TFrm_acertavalor_final_venda.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_acertavalor_final_venda, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQRecebimento', 'ZQRecebimentodt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertavalor_final_venda, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertavalor_final_venda, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParceladt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertavalor_final_venda, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradadt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertavalor_final_venda, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreelote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreelado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreematri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertavalor_final_venda, 'ZQtira_negativo_do_valor_recebimento', 'ZQtira_negativo_do_valor_recebimentoobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
