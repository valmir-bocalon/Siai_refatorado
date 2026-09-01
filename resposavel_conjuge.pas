unit resposavel_conjuge;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Gauges, StdCtrls, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset;

type
  TFrm_resposavel_conjuge = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Gauge1: TGauge;
    DS_Rep: TDataSource;
    DSconj: TDataSource;
    ZQRep: TZQuery;
    ZQConj: TZQuery;








































    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_resposavel_conjuge: TFrm_resposavel_conjuge;

implementation

uses tabelas, uRuntimeFields;

{$R *.dfm}

procedure TFrm_resposavel_conjuge.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_resposavel_conjuge.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Frm_resposavel_conjuge:=nil;
  Action:=CaFree;
end;

procedure TFrm_resposavel_conjuge.Button1Click(Sender: TObject);
var
  varposicao : integer;
begin
  ZQConj.SQL.Clear;
  ZQConj.SQL.Add('delete from conjuge');
  ZQConj.ExecSQL;
  ZQConj.SQL.Clear;
  ZQConj.SQL.Add('select * from conjuge');
  ZQConj.Open;
  ZQrep.Open;
  Gauge1.MaxValue := ZQRep.RecordCount;
  while not zqrep.Eof do begin
    Gauge1.Progress := zqrep.RecNo;
    ZQconj.Insert;
    ZQConj.FieldByName('nome').AsString        := ZQRep.FieldByName('nomerepres').AsString;
    ZQConj.FieldByName('idparti').AsInteger     := ZQRep.FieldByName('paticipante_idpaticipante').AsLargeInt;
    ZQConj.FieldByName('doc1').AsString        := ZQRep.FieldByName('documen1').AsString;
    ZQConj.FieldByName('doc2').AsString        := ZQRep.FieldByName('documen2').AsString;
    ZQConj.FieldByName('nasc').AsDateTime        := ZQRep.FieldByName('nasicmento').AsDateTime;
    ZQConj.FieldByName('natural').AsInteger     := ZQRep.FieldByName('cid_natural').AsLargeInt;
    ZQConj.FieldByName('email').AsString       := ZQRep.FieldByName('emailrepres').AsString;
    ZQConj.FieldByName('assina').AsString      := ZQRep.FieldByName('assina').AsString;
    ZQConj.FieldByName('empresa').AsString     := '';
    ZQConj.FieldByName('cargo').AsString       := ZQRep.FieldByName('trabalho').AsString;
    ZQConj.FieldByName('admissa').AsDateTime     := ZQRep.FieldByName('admissao').AsDateTime;
    ZQConj.FieldByName('renda').AsFloat       := ZQRep.FieldByName('renda').AsFloat;
    ZQConj.FieldByName('regime').AsString      := ZQRep.FieldByName('regime').AsString;
    ZQConj.FieldByName('comarca').AsFloat     := ZQRep.FieldByName('comarca').AsInteger;
    ZQConj.FieldByName('cidade_cart').AsFloat := ZQRep.FieldByName('cidade_cart').AsInteger;
    ZQConj.FieldByName('folha').AsString       := ZQRep.FieldByName('folha').AsString;
    ZQConj.FieldByName('livro').AsString       := ZQRep.FieldByName('livro').AsString;
    ZQConj.FieldByName('cartorio').AsString    := ZQRep.FieldByName('cartorio').AsString;
    ZQConj.FieldByName('data').AsDateTime        := ZQRep.FieldByName('data').AsDateTime;
    ZQconj.Post;
    zqrep.Next;
  end;
  ZQConj.SQL.Clear;
  ZQConj.SQL.Add('delete from representa_participante');
  ZQConj.ExecSQL;
  ZQConj.SQL.Clear;
  ZQConj.SQL.Add('select * from conjuge');
  ZQConj.Open; 
  DM_tabelas.ZQRecebimento.First;
  Gauge1.MaxValue :=  DM_Tabelas.ZQRecebimento.RecordCount;
  while not DM_Tabelas.ZQRecebimento.Eof do begin
    varposicao := strtoint(copy(DM_Tabelas.ZQRecebimento.FieldByName('documento').AsString,1,pos('-',DM_Tabelas.ZQRecebimento.FieldByName('documento').AsString)-1));
    Gauge1.Progress :=  DM_Tabelas.ZQRecebimento.RecNo;
    DM_TAbelas.ZQVenda.Locate('idvenda',varposicao,[]);
    DM_Tabelas.ZQRecebimento.Edit;
    DM_Tabelas.ZQRecebimento.FieldByName('idloteamento').AsInteger := DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger;
    DM_Tabelas.ZQRecebimento.Post;
    DM_Tabelas.ZQRecebimento.Next;
  end;
end;


procedure TFrm_resposavel_conjuge.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_resposavel_conjuge, 'ZQRep', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepidrepresenta_participante', 'idrepresenta_participante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQReppaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepdocumen1', 'documen1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepdocumen2', 'documen2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepnasicmento', 'nasicmento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepcid_natural', 'cid_natural', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepnomerepres', 'nomerepres', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepprofiss', 'profiss', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepcartorio', 'cartorio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQReplivro', 'livro', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepfolha', 'folha', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepcidade_cart', 'cidade_cart', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepcomarca', 'comarca', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQReptrabalho', 'trabalho', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepadmissao', 'admissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQReprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQRep', 'ZQRepemailrepres', 'emailrepres', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_resposavel_conjuge, 'ZQConj', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjnome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjdoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjdoc2', 'doc2', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjnasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjnatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjemail', 'email', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjempresa', 'empresa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjcargo', 'cargo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjcomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjcidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjfolha', 'folha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjcartorio', 'cartorio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_resposavel_conjuge, 'ZQConj', 'ZQConjlivro', 'livro', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
