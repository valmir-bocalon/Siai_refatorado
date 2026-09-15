unit acerta_adversa_recebimento;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Gauges, StdCtrls;

type
  TFrm_acertaadversa_recebimento = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButtonArround;
    dxButton1: TdxButtonArround;
    Label1: TLabel;
    ZQRecebimento: TZQuery;


    DS_Recebimento: TDataSource;
    DS_participante: TDataSource;
    ZQParticipante: TZQuery;




    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_acertaadversa_recebimento: TFrm_acertaadversa_recebimento;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_acertaadversa_recebimento.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_acertaadversa_recebimento.BTFecharClick(Sender: TObject);
begin
  ZQParticipante.open;
  ZQRecebimento.open;
  ZQRecebimento.First;

  Gauge1.MaxValue := ZQRecebimento.RecordCount;
  while not ZQRecebimento.Eof do begin
    Gauge1.Progress := ZQRecebimento.RecNo;
    ZQParticipante.Locate('idpaticipante',ZQRecebimento.FieldByName('adversa').AsInteger,[]);
    ZQRecebimento.Edit;
    ZQRecebimento.FieldByName('nomeadversa').AsString := ZQParticipante.FieldByName('nome_parte').AsString;
    ZQRecebimento.Post;
    ZQRecebimento.Next;
  end;
  ZQRecebimento.close;
  ZQParticipante.close;
  Close;
end;

procedure TFrm_acertaadversa_recebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQParticipante.close;
  ZQRecebimento.close;
  Frm_acertaadversa_recebimento:=nil;
  Action:=CaFree;
end;


procedure TFrm_acertaadversa_recebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_acertaadversa_recebimento, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_acertaadversa_recebimento, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaadversa_recebimento, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaadversa_recebimento, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaadversa_recebimento, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_acertaadversa_recebimento, 'ZQParticipante', False);
  RegisterRuntimeField(TFrm_acertaadversa_recebimento, 'ZQParticipante', 'ZQParticipanteidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaadversa_recebimento, 'ZQParticipante', 'ZQParticipantenome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
