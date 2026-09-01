unit zerarNossoNumero_em_recebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,  
  dxButton, Gauges, StdCtrls, dxCore2, Vcl.Mask;

type
  TFrm_Zerar_nosso_numero = class(TForm)
    Gauge1: TGauge;
    BTFechar: TdxButton;
    dxButton1: TdxButton;
    ql: TMaskEdit;
    Label1: TLabel;
    ZQRecBai: TZQuery;
    ZQremessa_receb: TZQuery;
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
  Frm_Zerar_nosso_numero: TFrm_Zerar_nosso_numero;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Zerar_nosso_numero.dxButton1Click(Sender: TObject);
begin
    ZQremessa_receb.Close;
    ZQRecBai.close;
    Close;

end;

procedure TFrm_Zerar_nosso_numero.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Frm_Zerar_nosso_numero:=nil;
  Action:=CaFree;
end;

procedure TFrm_Zerar_nosso_numero.BTFecharClick(Sender: TObject);
begin
    ZQRecBai.close;
    ZQRecBai.SQL.Clear;
    ZQRecBai.SQL.Add('update recebimento set numboleto=''''  where saldo>''0'' and quadralote='+quotedstr(ql.Text)+ ' and Marcar=''0''');
    ZQRecBai.ExecSQL;
    ZQRecBai.close;

    ZQRecBai.SQL.Clear;
    ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
    ZQRecBai.SQL.Add(' from recebimento where marcar=''0'' and  saldo>''0'' and quadralote='+quotedstr(ql.Text));
    ZQRecBai.open;
    ZQRecBai.First;

    Gauge1.MaxValue:=ZQRecBai.RecordCount;
    Gauge1.Progress:=0;
    while not ZQrecBai.Eof do
    begin
      Gauge1.Progress:=ZQRecBai.RecNo;
      ZQremessa_receb.Close;
      ZQremessa_receb.SQL.Clear;
      ZQremessa_receb.SQL.Add('delete from remessa_receb where idrec='+quotedstr(ZQRecBai.FieldByName('idrecebimento').AsString));
      ZQremessa_receb.ExecSQL;

      ZQrecBai.Next;
    end;
    Gauge1.Progress:=0;
    ZQremessa_receb.Close;
    ZQRecBai.close;
    Close;
end;


procedure TFrm_Zerar_nosso_numero.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
