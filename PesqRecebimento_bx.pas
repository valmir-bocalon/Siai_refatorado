unit PesqRecebimento_bx;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XLabel3D, XBanner, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Grids, DBGrids, Mask, DBCtrls;

type
  TFrmPesqRecebimento_bx = class(TForm)
    XLTit: TXLabel3D;
    DBGrid1: TDBGrid;
    Label1: TLabel;
    DS_TempCliReceb: TDataSource;
    ZQTempCliReceb: TZQuery;
    Edit1: TEdit;
    XBanner6: TXBanner;




















































    DBECod: TDBEdit;
    Label11: TLabel;
    LReg: TLabel;









































    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmPesqRecebimento_bx: TFrmPesqRecebimento_bx;

implementation

uses Tabelas, Cad_Recebimento, Funcoes, RecebBaixa, uRuntimeFields;

{$R *.dfm}

procedure TFrmPesqRecebimento_bx.FormShow(Sender: TObject);
begin
  Edit1.SetFocus;
end;

procedure TFrmPesqRecebimento_bx.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then
    close;

end;

procedure TFrmPesqRecebimento_bx.Edit1KeyPress(Sender: TObject;
  var Key: Char);
var
  ValorPesquisa: Double;
  CodigoValor: Integer;
  TextoValor: string;
begin
  if Key = #13 then begin
  ZQTempCliReceb.SQL.Clear;
  ZQTempCliReceb.SQL.Add('select * from recebimento left join participante as p1 on cliente=p1.idpaticipante');
  ZQTempCliReceb.SQL.Add(' left join participante as p2 on adversa=p2.idpaticipante');
  ZQTempCliReceb.SQL.Add(' where p1.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or p2.nome_parte like '+quotedstr('%'+Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or documento like '+quotedstr('%'+Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or idrecebimento like '+quotedstr(Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or refbaixa like '+quotedstr(Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or refvinda like '+quotedstr(Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or quadralote like '+quotedstr(Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or numboleto like '+quotedstr(Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or ordem like '+quotedstr(Edit1.Text+'%'));
  ZQTempCliReceb.SQL.Add(' or numboleto like '+quotedstr('%'+Edit1.Text+'%'));
  TextoValor := StringReplace(Trim(Edit1.Text), ',', '.', [rfReplaceAll]);
  Val(TextoValor, ValorPesquisa, CodigoValor);
  if (TextoValor <> '') and (CodigoValor = 0) then
  begin
    ZQTempCliReceb.SQL.Add(' or valor=:valor_pesquisa');
    ZQTempCliReceb.ParamByName('valor_pesquisa').AsFloat := ValorPesquisa;
  end;
  ZQTempCliReceb.SQL.Add(' group by idrecebimento');
  ZQTempCliReceb.Open;
  LReg.Caption := inttostr(ZQTempCliReceb.RecordCount);
  if (ZQTempCliReceb.Active) and (ZQTempCliReceb.RecordCount>0) Then
      DBGrid1.SetFocus
    else
      close;
  end;
end;

procedure TFrmPesqRecebimento_bx.DBGrid1KeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (ZQTempCliReceb.RecordCount>0) Then Begin
    DM_tabelas.ZQRecebimento.Locate('idrecebimento',ZQTempCliReceb.FieldByName('idrecebimento').AsLargeInt,[]);
    Close
  end;
end;

procedure TFrmPesqRecebimento_bx.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = ZQTempCliReceb.FieldByName('idrecebimento').Text Then Begin
    DBGrid1.Canvas.Brush.Color := $006CFFFF;
    DBGrid1.Canvas.Font.Color := $00A80000;
//    DBGReceb.Canvas.Font.Style := [FsBold];
  end;
  DBGrid1.DefaultDrawDataCell(Rect, DBGrid1.columns[datacol].field, State);

end;


procedure TFrmPesqRecebimento_bx.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebtipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliReceblocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebidpaticipante_1', 'idpaticipante_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnome_parte_1', 'nome_parte_1', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebdoc1_1', 'doc1_1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebdoc2_1', 'doc2_1', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebendereco_1', 'endereco_1', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebbairro_1', 'bairro_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcidade_1', 'cidade_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcep_1', 'cep_1', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebende_cob_1', 'ende_cob_1', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebbairro_cob_1', 'bairro_cob_1', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcidade_cob_1', 'cidade_cob_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcep_cob_1', 'cep_cob_1', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebFone1_1', 'Fone1_1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebfone2_1', 'fone2_1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebfone3_1', 'fone3_1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebtipopessoa_1', 'tipopessoa_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebaniversario_1', 'aniversario_1', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnacionalidade_1', 'nacionalidade_1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebemail_1', 'email_1', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebnaturalidade_1', 'naturalidade_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliReceblocaldetrab_1', 'localdetrab_1', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebprofissao_1', 'profissao_1', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebestadocivil_1', 'estadocivil_1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebrenda_1', 'renda_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebobservacao_1', 'observacao_1', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecinc_1', 'codrecinc_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodrecBx_1', 'codrecBx_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodpaginc_1', 'codpaginc_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcodpagBx_1', 'codpagBx_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebfone4_1', 'fone4_1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcomplemento_1', 'complemento_1', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcomplemento_cob_1', 'complemento_cob_1', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmPesqRecebimento_bx, 'ZQTempCliReceb', 'ZQTempCliRecebcadastrado_1', 'cadastrado_1', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
