unit acertaorigem_em_recebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,  
  dxButton, Gauges, StdCtrls, dxCore2;

type
  TFrm_acertaorigem_em_recebimento = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    Label1: TLabel;
    ZQRecebimento: TZQuery;
    DS_Recebimento: TDataSource;


    procedure dxButton1Click(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_acertaorigem_em_recebimento: TFrm_acertaorigem_em_recebimento;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_acertaorigem_em_recebimento.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_acertaorigem_em_recebimento.BTFecharClick(Sender: TObject);
begin
  ZQRecebimento.open;
  ZQRecebimento.First;
  Gauge1.MaxValue := ZQRecebimento.RecordCount;
  while not ZQRecebimento.Eof do begin
    Gauge1.Progress := ZQRecebimento.RecNo;
    if ZQRecebimento.FieldByName('venda_idvenda').AsInteger>0 then
    begin
      ZQRecebimento.Edit;
      ZQRecebimento.FieldByName('origem').AsString := 'V-'+ZQRecebimento.FieldByName('venda_idvenda').Text;
      ZQRecebimento.Post;
    end;  
    ZQRecebimento.Next;
  end;
  ZQRecebimento.close;
  Close;
end;


procedure TFrm_acertaorigem_em_recebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_acertaorigem_em_recebimento, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_acertaorigem_em_recebimento, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_acertaorigem_em_recebimento, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
