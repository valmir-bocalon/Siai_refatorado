unit zerar_recebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ZConnection, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  StdCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    DS_Price: TDataSource;
    ZQrecebimento: TZQuery;
    zconeccao: TZConnection;
    zqreceb_baixa: TZQuery;
    DS_zm1: TDataSource;
    zqrecbxhist: TZQuery;
    Data_zm1_5: TDataSource;
    Button1: TButton;
    Button2: TButton;
    br: TProgressBar;































































    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses
  uRuntimeFields;

{$R *.dfm}

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ZQrecebimento.close;
  zqreceb_baixa.close;
  zqrecbxhist.close;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  ZQrecebimento.close;
  zqreceb_baixa.close;
  zqrecbxhist.close;

end;

procedure TForm1.Button1Click(Sender: TObject);
begin
    br.Position:=0;
    ZQRecebimento.close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add('SELECT * FROM recebimento where quadralote ');
    ZQRecebimento.SQL.Add(' not in (''ne.12.09.17-14'',''ne.12.09.17-15'',''ne.12.09.17-16'',''ne.12.09.17-17'',''ne.12.09.17-18'',''ne.12.09.17-19'')');
    ZQRecebimento.open;
    if ZQRecebimento.RecordCount>0 then
    begin
      br.Max:= ZQRecebimento.RecordCount;
      br.Position:=0;
      br.Visible:=true;
      while not  ZQRecebimento.Eof do
      begin
        br.Position:= ZQRecebimento.RecNo;
        zqrecbxhist.SQL.clear;
        zqrecbxhist.SQL.Add('delete from recbxhist where idrecib='+quotedstr( ZQRecebimento.FieldByName('idrecebimento').Text));
        zqrecbxhist.ExecSQL;

        ZQReceb_Baixa.SQL.clear;
        ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where refbaixa='+quotedstr( ZQRecebimento.FieldByName('RefBaixa').Text));
        ZQReceb_Baixa.ExecSQL;

        {ZQAditamento.SQL.clear;
        ZQAditamento.SQL.Add('delete from aditamento where idvenda='+quotedstr( ZQRecebimentovenda_idvenda.Text));
        ZQAditamento.ExecSQL;

        ZQQuitacao.SQL.clear;
        ZQQuitacao.SQL.Add('delete from quitacao where idvenda='+quotedstr( ZQRecebimentovenda_idvenda.Text));
        ZQQuitacao.ExecSQL;

        ZQcomprador_cessao.SQL.clear;
        ZQcomprador_cessao.SQL.Add('delete from comprador_cessao where idvenda='+quotedstr( ZQRecebimentovenda_idvenda.Text));
        ZQcomprador_cessao.ExecSQL;


        ZQrecebimento_historico_excluidos.SQL.clear;
        ZQrecebimento_historico_excluidos.SQL.Add('delete from recebimento_historico_excluidos where idrecebimento='+quotedstr( ZQRecebimentoidrecebimento.Text));
        ZQrecebimento_historico_excluidos.ExecSQL;

        ZQrecebimento_historico_cessao.SQL.clear;
        ZQrecebimento_historico_cessao.SQL.Add('delete from recebimento_historico_cessao where idrecebimento='+quotedstr( ZQRecebimentoidrecebimento.Text));
        ZQrecebimento_historico_cessao.ExecSQL;}

        ZQRecebimento.Next;
      end;
    end;
    br.Position:=0;
    br.Visible:=false;
    ZQRecebimento.close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add('delete FROM recebimento where quadralote ');
    ZQRecebimento.SQL.Add(' not in (''ne.12.09.17-14'',''ne.12.09.17-15'',''ne.12.09.17-16'',''ne.12.09.17-17'',''ne.12.09.17-18'',''ne.12.09.17-19'')');
    ZQRecebimento.ExecSQL;

    showmessage('Exclusão Efetuada Com Sucesso!');
    close;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  close;
end;


procedure TForm1.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TForm1, 'ZQrecebimento', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentonumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'ZQrecebimento', 'ZQrecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'zqreceb_baixa', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixaVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqreceb_baixa', 'zqreceb_baixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TForm1, 'zqrecbxhist', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TForm1, 'zqrecbxhist', 'zqrecbxhistvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
