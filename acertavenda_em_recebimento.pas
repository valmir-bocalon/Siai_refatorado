unit acertavenda_em_recebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,  
  dxButton, Gauges, StdCtrls, dxCore2;

type
  TFrm_acertavenda_em_recebimento = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    Label1: TLabel;
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
  Frm_acertavenda_em_recebimento: TFrm_acertavenda_em_recebimento;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_acertavenda_em_recebimento.dxButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_acertavenda_em_recebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Frm_acertavenda_em_recebimento:=nil;
  Action:=CaFree;
end;

procedure TFrm_acertavenda_em_recebimento.BTFecharClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Open;
  DM_TAbelas.ZQRecebimento.Open;
  DM_TAbelas.ZQRecebimento.First;
  Gauge1.MaxValue := DM_Tabelas.ZQRecebimento.RecordCount;
  while not DM_TAbelas.ZQRecebimento.Eof do begin
    Gauge1.Progress := DM_Tabelas.ZQRecebimento.RecNo;
    if not empty(DM_TAbelas.ZQRecebimento.FieldByName('documento').AsString) Then Begin
      if DM_Tabelas.ZQVenda.Locate('idvenda',copy(DM_TAbelas.ZQRecebimento.FieldByName('documento').AsString,1,pos('-',DM_TAbelas.ZQRecebimento.FieldByName('documento').AsString)-1),[]) Then Begin
        DM_TAbelas.ZQRecebimento.Edit;
        DM_TAbelas.ZQRecebimento.FieldByName('venda_idvenda').AsInteger := DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
        DM_Tabelas.ZQRecebimento.FieldByName('quadralote').AsString    := DM_TAbelas.ZQVenda.FieldByName('quadra').AsString+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').Text;
        DM_Tabelas.ZQRecebimento.Post;
      end;
    end;
    DM_Tabelas.ZQRecebimento.Next;
  end;
  DM_Tabelas.ZQVenda.close;
  DM_TAbelas.ZQRecebimento.close;
  Close;
end;


procedure TFrm_acertavenda_em_recebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
