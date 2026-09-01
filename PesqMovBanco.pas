unit PesqMovBanco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, StdCtrls, XBanner;

type
  TFrmPesqMovBanco = class(TForm)
    XBanner1: TXBanner;
    Label5: TLabel;
    Epesq: TEdit;
    DS_PesqMovBanco: TDataSource;
    ZQPesqMovBanco: TZQuery;
    DBGPesqMovBanco: TDBGrid;


























































    Label1: TLabel;
    Lreg: TLabel;
    procedure FormShow(Sender: TObject);
    procedure EpesqChange(Sender: TObject);
    procedure ZQPesqMovBancoCalcFields(DataSet: TDataSet);
    procedure EpesqKeyPress(Sender: TObject; var Key: Char);
    procedure DBGPesqMovBancoKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPesqMovBanco: TFrmPesqMovBanco;

implementation

uses tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmPesqMovBanco.FormShow(Sender: TObject);
begin
  XBanner1.ColorOf := DM_Tabelas.ZQconfiguracaocorup.Value;
  XBanner1.ColorFor := DM_TAbelas.ZQconfiguracaocordw.Value;
  Label5.Font.Color := DM_Tabelas.ZQconfiguracaocortopo.Value;
  Epesq.SetFocus;
end;

procedure TFrmPesqMovBanco.EpesqChange(Sender: TObject);
begin
  ZQPesqMovBanco.SQL.Clear;
  ZQPesqMovBanco.SQL.Add('select * from mov_banco as MB join cliente as CL ON contracodigo=CL.codigo');
  ZQPesqMovBanco.SQL.Add(' join plano_contas as PC ON pc.codigo=plano_contas_codigo');
  ZQPesqMovBanco.SQL.Add(' where Conta_Bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
  if not empty(Epesq.Text) Then Begin
    ZQPesqMovBanco.SQL.Add(' and (descricao like '+quotedstr('%'+Epesq.Text+'%'));
    ZQPesqMovBanco.SQL.Add(' or nome like '+quotedstr('%'+Epesq.Text+'%'));
    ZQPesqMovBanco.SQL.Add(' or documento like '+quotedstr('%'+Epesq.Text+'%')+')');
  end;
  ZQPesqMovBanco.Open;
  Lreg.caption := inttostr(ZQPesqMovBanco.RecordCount);
end;

procedure TFrmPesqMovBanco.ZQPesqMovBancoCalcFields(DataSet: TDataSet);
begin
  ZQPesqMovBanco.FieldByName('plano_descri').Value := ZQPesqMovBanco.FieldByName('Conta_Bancaria_cod_banco').Text+' - '+ZQPesqMovBanco.FieldByName('descricao').Value;
end;

procedure TFrmPesqMovBanco.EpesqKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    Close;
  if key = #13 Then
    if ZQPesqMovBanco.Active Then Begin
      if ZQPesqMovBanco.RecordCount>1 Then
        DBGPesqMovBanco.SetFocus
      else if ZQPesqMovBanco.RecordCount=0 Then
        Close
      else Begin
        DM_Tabelas.ZQMovBancaria.Locate('lancamento',ZQPesqMovBanco.FieldByName('lancamento').Text,[]);
        Close;
      End;
    end
    else
      Close;
end;

procedure TFrmPesqMovBanco.DBGPesqMovBancoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 Then
    DM_Tabelas.ZQMovBancaria.Locate('lancamento',ZQPesqMovBanco.FieldByName('lancamento').Text,[]);
  Close;
end;


procedure TFrmPesqMovBanco.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmPesqMovBanco, 'ZQPesqMovBanco', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancolancamento', 'lancamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoConta_Bancaria_cod_banco', 'Conta_Bancaria_cod_banco', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoPlano_Contas_codigo', 'Plano_Contas_codigo', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodocumento', 'documento', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodt_lanc', 'dt_lanc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancohora_lanc', 'hora_lanc', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodt_conciliado', 'dt_conciliado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancovr_lanc', 'vr_lanc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoCreDeb', 'CreDeb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoSaldo_lanc', 'Saldo_lanc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoSaldo_conci', 'Saldo_conci', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoobs', 'obs', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocontraquem', 'contraquem', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocontracodigo', 'contracodigo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancovinculo', 'vinculo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocodigo', 'codigo', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBanconome', 'nome', TStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancofantasia', 'fantasia', TStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancopessoa_fis_jur', 'pessoa_fis_jur', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoendereco', 'endereco', TStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancobairro', 'bairro', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocidade', 'cidade', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoestado', 'estado', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoponto_referencia', 'ponto_referencia', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancopostal', 'postal', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancotelefone', 'telefone', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocelular', 'celular', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoemail', 'email', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoresponsavel', 'responsavel', TStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocargo', 'cargo', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoendereco_cob', 'endereco_cob', TStringField, fkData, 60, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancobairro_cob', 'bairro_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocidade_cob', 'cidade_cob', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoestado_cob', 'estado_cob', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoativo', 'ativo', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodoc1', 'doc1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodoc2', 'doc2', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBanconascido', 'nascido', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocredito', 'credito', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancotipo', 'tipo', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocodigo_1', 'codigo_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancomascara', 'mascara', TStringField, fkData, 16, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodescricao', 'descricao', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoclassificacao', 'classificacao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocred_debe', 'cred_debe', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancocod_reduzido', 'cod_reduzido', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancodoccomum', 'doccomum', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancousoativo', 'usoativo', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoCDPlano', 'CDPlano', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoDespFixVar', 'DespFixVar', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqMovBanco, 'ZQPesqMovBanco', 'ZQPesqMovBancoplano_descri', 'plano_descri', TStringField, fkCalculated, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
