unit acerta_adiamento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,  
  dxButton, Gauges, StdCtrls, dxCore2;

type
  TFrm_acertaaditamento = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    Label1: TLabel;
    zqaditamento: TZQuery;










    DS_zqaditamento: TDataSource;

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
  Frm_acertaaditamento: TFrm_acertaaditamento;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_acertaaditamento.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_acertaaditamento.BTFecharClick(Sender: TObject);
begin
  zqaditamento.close;
  zqaditamento.Open;
  zqaditamento.first;
  Gauge1.Progress:=0;
  Gauge1.MaxValue:=zqaditamento.RecordCount;
  while not zqaditamento.eof do
  begin
    Gauge1.Progress:=zqaditamento.RecNo;
    zqaditamento.edit;
    zqaditamento.FieldByName('historico').AsString:='Data: '+datetostr(zqaditamento.FieldByName('datacessao').AsDateTime)+' Hora: '+timetostr(zqaditamento.FieldByName('hora').AsDateTime)+' - '+zqaditamento.FieldByName('historico').AsString;
    zqaditamento.post;
    zqaditamento.next;
  end;
  Gauge1.Progress:=0;
  Close;
end;

procedure TFrm_acertaaditamento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zqaditamento.close;
  Frm_acertaaditamento:=nil;
  Action:=CaFree;
end;


procedure TFrm_acertaaditamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_acertaaditamento, 'zqaditamento', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaaditamento, 'zqaditamento', 'zqaditamentohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
