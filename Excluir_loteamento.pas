unit Excluir_loteamento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,  
  dxButton, Gauges, StdCtrls, Mask, dxCore2;

type
  TFrm_Excluir = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    ZQSelecao: TZQuery;
    DS_Selecao: TDataSource;

































































    Label2: TLabel;
    cdlt: TMaskEdit;
    ZQAditamento: TZQuery;
    Dataaditamento: TDataSource;











    DataZQCheque_rec: TDataSource;
    ZQCheque_rec: TZQuery;

















    ZQcomprador: TZQuery;
    Datazqcomprador: TDataSource;





    ZQcomprador_cessao: TZQuery;
    Datacomprador_cessao: TDataSource;











    ZQimovel: TZQuery;
    Dataimovel: TDataSource;























    Datazqincorporador_loteamento: TDataSource;
    zqincorporador_loteamento: TZQuery;





    DataZQIncorporador: TDataSource;
    ZQIncorporador: TZQuery;






    ZQlogin_historico: TZQuery;
    Datalogin_historico: TDataSource;








    DataZQProcuradores: TDataSource;
    ZQProcuradores: TZQuery;





    ZQquadras: TZQuery;
    Datazqquadras: TDataSource;




    Datazqquitacao: TDataSource;
    ZQquitacao: TZQuery;











    Datazqrecebimento_historico_excluidos: TDataSource;
    zqrecebimento_historico_excluidos: TZQuery;






























    Dataremessa_receb: TDataSource;
    ZQremessa_receb: TZQuery;










    ZQvendedor: TZQuery;
    DataZQvendedor: TDataSource;



    ZQRecebimento: TZQuery;
    Datarecebimento: TDataSource;


































    DataZQReBxHi: TDataSource;
    ZQReBxHi: TZQuery;









    DataReceb_Baixa: TDataSource;
    ZQReceb_Baixa: TZQuery;















    Datazqvenda: TDataSource;
    ZQvenda: TZQuery;







    Dataloteamento: TDataSource;
    ZQLoteamento: TZQuery;




















    Label1: TLabel;
    ccli: TMaskEdit;

    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Excluir: TFrm_Excluir;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Excluir.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Excluir.BTFecharClick(Sender: TObject);
begin
  if empty(cdlt.Text) then
  begin
    showmessage('Informe o Código do Loteamento.');
    cdlt.SetFocus;
    exit;
  end;
  if SIMNAO('Deseja Realmente Excluir esse Registro ?','SIM') then
  begin
    ZQSelecao.SQL.Clear;
    ZQSelecao.SQL.Add('SELECT * FROM venda as v join recebimento as rec on v.idvenda=rec.venda_idvenda join recbxhist as rx on rx.idrecib=rec.idrecebimento join Receb_Baixa as bx on bx.refbaixa=rec.refbaixa where rec.idloteamento='+quotedstr(alltrim(cdlt.text)));
    ZQSelecao.open;
    ZQSelecao.First;
    Gauge1.Progress:=0;
    if ZQSelecao.RecordCount>0 then
    begin
      Gauge1.MaxValue:=ZQselecao.RecordCount;
      while not ZQSelecao.Eof do
      begin
        Gauge1.Progress:=ZQselecao.RecNo;
        ZQRecebimento.close;
        ZQRecebimento.SQL.Clear;
//        ZQRecebimento.SQL.Add('select * from recebimento where documento like '+quotedstr(ZQSelecaoidvenda.text+'-%'));
        ZQRecebimento.SQL.Add('select * from recebimento where venda_idvenda='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQRecebimento.open;
        if ZQRecebimento.RecordCount>0 then
        begin
          while not ZQRecebimento.Eof do
          begin
            ZQReBxHi.close;
            ZQReBxHi.SQL.clear;
            ZQReBxHi.SQL.Add('delete from recbxhist where idrecib='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
            ZQReBxHi.ExecSQL;

            ZQReceb_Baixa.close;
            ZQReceb_Baixa.SQL.clear;
            ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr(ZQRecebimento.FieldByName('RefBaixa').Text));
            ZQReceb_Baixa.ExecSQL;

            zqrecebimento_historico_excluidos.close;
            zqrecebimento_historico_excluidos.SQL.clear;
            zqrecebimento_historico_excluidos.SQL.Add('delete from recebimento_historico_excluidos where idrecebimento='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
            zqrecebimento_historico_excluidos.ExecSQL;

            ZQremessa_receb.close;
            ZQremessa_receb.SQL.clear;
            ZQremessa_receb.SQL.Add('delete from remessa_receb where idrec='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
            ZQremessa_receb.ExecSQL;


            ZQRecebimento.Next;
          end;
        end;
        ZQRecebimento.close;
        ZQRecebimento.SQL.Clear;
//        ZQRecebimento.SQL.Add('delete from recebimento where documento like '+quotedstr(ZQSelecaoidvenda.text+'-%'));
        ZQRecebimento.SQL.Add('delete from recebimento where venda_idvenda='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQRecebimento.ExecSQL;

{        ZQRecebimento.close;
        ZQRecebimento.SQL.Clear;
        ZQRecebimento.SQL.Add('select * from recebimento ');
        ZQRecebimento.Open;

        DM_Tabelas.ZQReBxHi.SQL.Clear;
        DM_Tabelas.ZQReBxHi.SQL.Add('select * from recbxhist ');
        DM_Tabelas.ZQReBxHi.Open;

        DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
        DM_Tabelas.ZQReceb_Baixa.SQL.Add('select * from Receb_Baixa ');
        DM_Tabelas.ZQReceb_Baixa.Open;                                 }

        ZQVenda.close;
        ZQVenda.SQL.Clear;
        ZQVenda.SQL.Add('delete from venda where idvenda ='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQVenda.ExecSQL;

        ZQLoteamento.close;
        ZQLoteamento.SQL.Clear;
        ZQLoteamento.SQL.Add('delete from loteamento where idloteamento='+quotedstr(ZQSelecao.FieldByName('idloteamento').Text));
        ZQLoteamento.ExecSQL;

{        DM_Tabelas.ZQLoteamento.SQL.Clear;
        DM_Tabelas.ZQLoteamento.SQL.Add('select * from loteamento ');
        DM_Tabelas.ZQLoteamento.Open;}

        ZQAditamento.close;
        ZQAditamento.SQL.Clear;
        ZQAditamento.SQL.Add('delete from aditamento where idvenda ='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQAditamento.ExecSQL;

        ZQcheque_rec.close;
        ZQcheque_rec.SQL.clear;
        ZQcheque_rec.SQL.Add('delete from cheque_rec where idrecebimento='+quotedstr(ZQSelecao.FieldByName('idrecebimento').Text));
        ZQcheque_rec.ExecSQL;

        ZQcomprador.close;
        ZQcomprador.SQL.Clear;
        ZQcomprador.SQL.Add('delete from comprador where venda_idvenda ='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQcomprador.ExecSQL;

        ZQcomprador_cessao.close;
        ZQcomprador_cessao.SQL.Clear;
        ZQcomprador_cessao.SQL.Add('delete from comprador_cessao where idvenda ='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQcomprador_cessao.ExecSQL;

        ZQimovel.close;
        ZQimovel.SQL.Clear;
        ZQimovel.SQL.Add('delete from imovel where loteamento_idloteamento='+quotedstr(alltrim(cdlt.text)));
        ZQimovel.ExecSQL;

        zqincorporador_loteamento.close;
        zqincorporador_loteamento.SQL.Clear;
        zqincorporador_loteamento.SQL.Add('select * from incorporador_loteamento where loteamento_idloteamento='+quotedstr(alltrim(cdlt.text)));
        zqincorporador_loteamento.open;
        zqincorporador_loteamento.First;
        if zqincorporador_loteamento.RecordCount>0 then
        begin
          while not zqincorporador_loteamento.Eof do
          begin
            ZQIncorporador.close;
            ZQIncorporador.SQL.Clear;
            ZQIncorporador.SQL.Add('delete from incorporador where idincorporador='+quotedstr(zqincorporador_loteamento.FieldByName('incorporador_idincorporador').Text));
            ZQIncorporador.ExecSQL;
            zqincorporador_loteamento.Next;
          end;
        end;
        zqincorporador_loteamento.close;
        zqincorporador_loteamento.SQL.Clear;
        zqincorporador_loteamento.SQL.Add('delete from incorporador_loteamento where loteamento_idloteamento='+quotedstr(alltrim(cdlt.text)));
        zqincorporador_loteamento.ExecSQL;

        ZQlogin_historico.close;
        ZQlogin_historico.SQL.Clear;
        ZQlogin_historico.SQL.Add('delete from login_historico');
        ZQlogin_historico.ExecSQL;

        zqprocuradores.close;
        zqprocuradores.SQL.Clear;
        zqprocuradores.SQL.Add('delete from procuradores where idloteamento='+quotedstr(alltrim(cdlt.text)));
        zqprocuradores.ExecSQL;

        zqquadras.Close;
        zqquadras.SQL.Clear;
        zqquadras.SQL.Add('delete from quadras where loteamento_idloteamento='+quotedstr(alltrim(cdlt.text)));
        zqquadras.ExecSQL;


        ZQquitacao.SQL.Clear;
        ZQquitacao.SQL.Add('delete from quitacao where idvenda ='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQquitacao.ExecSQL;

        ZQvendedor.close;
        ZQvendedor.SQL.Clear;
        ZQvendedor.SQL.Add('delete from vendedor where venda_idvenda ='+quotedstr(ZQSelecao.FieldByName('idvenda').text));
        ZQvendedor.ExecSQL;

        ZQSelecao.Next;
      end;
    end;

    ZQRecebimento.close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add('select * from recebimento where idloteamento='+quotedstr(alltrim(cdlt.text)));
    ZQRecebimento.open;
    if ZQRecebimento.RecordCount>0 then
    begin
      while not ZQRecebimento.Eof do
      begin
        ZQReBxHi.close;
        ZQReBxHi.SQL.clear;
        ZQReBxHi.SQL.Add('delete from recbxhist where idrecib='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
        ZQReBxHi.ExecSQL;

        ZQReceb_Baixa.close;
        ZQReceb_Baixa.SQL.clear;
        ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr(ZQRecebimento.FieldByName('RefBaixa').Text));
        ZQReceb_Baixa.ExecSQL;

        zqrecebimento_historico_excluidos.close;
        zqrecebimento_historico_excluidos.SQL.clear;
        zqrecebimento_historico_excluidos.SQL.Add('delete from recebimento_historico_excluidos where idrecebimento='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
        zqrecebimento_historico_excluidos.ExecSQL;

        ZQremessa_receb.close;
        ZQremessa_receb.SQL.clear;
        ZQremessa_receb.SQL.Add('delete from remessa_receb where idrec='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
        ZQremessa_receb.ExecSQL;


        ZQRecebimento.delete;
      end;
    end;


    ZQRecebimento.close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add('select * from recebimento where cliente='+quotedstr(alltrim(ccli.text)));
    ZQRecebimento.open;
    if ZQRecebimento.RecordCount>0 then
    begin
      while not ZQRecebimento.Eof do
      begin
        ZQReBxHi.close;
        ZQReBxHi.SQL.clear;
        ZQReBxHi.SQL.Add('delete from recbxhist where idrecib='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
        ZQReBxHi.ExecSQL;

        ZQReceb_Baixa.close;
        ZQReceb_Baixa.SQL.clear;
        ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr(ZQRecebimento.FieldByName('RefBaixa').Text));
        ZQReceb_Baixa.ExecSQL;

        zqrecebimento_historico_excluidos.close;
        zqrecebimento_historico_excluidos.SQL.clear;
        zqrecebimento_historico_excluidos.SQL.Add('delete from recebimento_historico_excluidos where idrecebimento='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
        zqrecebimento_historico_excluidos.ExecSQL;

        ZQremessa_receb.close;
        ZQremessa_receb.SQL.clear;
        ZQremessa_receb.SQL.Add('delete from remessa_receb where idrec='+quotedstr(ZQRecebimento.FieldByName('idrecebimento').Text));
        ZQremessa_receb.ExecSQL;


        ZQRecebimento.delete;
      end;
    end;


    Gauge1.Progress:=0;
    ZQVenda.Close;
    ZQSelecao.Close;
    ZQReceb_Baixa.Close;
    ZQReBxHi.Close;
    ZQRecebimento.Close;
    ZQLoteamento.Close;
    zqrecebimento_historico_excluidos.close;
    ZQremessa_receb.close;
    ZQvendedor.close;
    zqquadras.close;
    ZQquitacao.Close;
    ZQlogin_historico.close;
    zqprocuradores.Close;
    ZQimovel.close;
    zqincorporador_loteamento.Close;
    ZQcomprador.close;
    ZQcomprador_cessao.Close;
    ZQAditamento.close;
    ZQcheque_rec.close;
    ccli.clear;
    cdlt.clear;
    cdlt.SetFocus;

        // manualmente
{        conta_bancaria
         usuario
         empresa}





//    close;
  end;
end;


procedure TFrm_Excluir.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQSelecao', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaotabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaonumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaorefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaovalor_1', 'valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaosq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaodataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaovencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaosq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQSelecao', 'ZQSelecaoVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQAditamento', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQAditamento', 'ZQAditamentocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQCheque_rec', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recbanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recdono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recalias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recnumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQCheque_rec', 'ZQCheque_recsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQcomprador', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador', 'ZQcompradoridcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador', 'ZQcompradorpaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador', 'ZQcompradorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador', 'ZQcompradorpromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador', 'ZQcompradorpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQcomprador_cessao', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQcomprador_cessao', 'ZQcomprador_cessaoautocodigo', 'autocodigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQimovel', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovellote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovellado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelcornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimoveldescricao', 'descricao', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimoveldisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQimovel', 'ZQimovelmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'zqincorporador_loteamento', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqincorporador_loteamento', 'zqincorporador_loteamentoincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqincorporador_loteamento', 'zqincorporador_loteamentoloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqincorporador_loteamento', 'zqincorporador_loteamentopercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqincorporador_loteamento', 'zqincorporador_loteamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqincorporador_loteamento', 'zqincorporador_loteamentodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQIncorporador', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQIncorporador', 'ZQIncorporadoridincorporador', 'idincorporador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQIncorporador', 'ZQIncorporadornomeincorp', 'nomeincorp', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQIncorporador', 'ZQIncorporadordoc', 'doc', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQIncorporador', 'ZQIncorporadorobserva', 'observa', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQIncorporador', 'ZQIncorporadorendereco_inco', 'endereco_inco', TWideStringField, fkData, 70, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQIncorporador', 'ZQIncorporadorcidade_inco', 'cidade_inco', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQlogin_historico', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicoidlogin_historico', 'idlogin_historico', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicolocal_senha', 'local_senha', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicodata_2', 'data_2', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicomaquina', 'maquina', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicousuario', 'usuario', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicopassou', 'passou', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQlogin_historico', 'ZQlogin_historicodetalhe', 'detalhe', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQProcuradores', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQProcuradores', 'ZQProcuradoresidProcuradores', 'idProcuradores', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQProcuradores', 'ZQProcuradoresidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQProcuradores', 'ZQProcuradoresidparti', 'idparti', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQProcuradores', 'ZQProcuradoresnome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQProcuradores', 'ZQProcuradoresassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQquadras', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquadras', 'ZQquadrasidquadras', 'idquadras', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquadras', 'ZQquadrasloteamento_idLoteamento', 'loteamento_idLoteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquadras', 'ZQquadrasdescricao', 'descricao', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquadras', 'ZQquadrasapedlido', 'apedlido', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQquitacao', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQquitacao', 'ZQquitacaocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'zqrecebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'zqrecebimento_historico_excluidos', 'zqrecebimento_historico_excluidosmotivo', 'motivo', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQremessa_receb', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebretorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebacao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebNossonumero', 'Nossonumero', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebdesc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebcredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQremessa_receb', 'ZQremessa_recebdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQvendedor', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvendedor', 'ZQvendedoridvendedor', 'idvendedor', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvendedor', 'ZQvendedorcorretor_idcorretor', 'corretor_idcorretor', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvendedor', 'ZQvendedorvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQReBxHi', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHiidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHirefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHiidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHivalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHidescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHipercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHidata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReBxHi', 'ZQReBxHisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQReceb_Baixa', 'ZQReceb_BaixaVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQvenda', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendatabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQvenda', 'ZQvendaEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Excluir, 'ZQLoteamento', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentoidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentomatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentotestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentotestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentotestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentotestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentoObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentobairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentopasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentonomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentotipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentoapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Excluir, 'ZQLoteamento', 'ZQLoteamentoPerc_comissao', 'Perc_comissao', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
end.
