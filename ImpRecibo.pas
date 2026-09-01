unit ImpRecibo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxButton, StdCtrls, XBanner, dxCore2;

type
  TFrmImpRecibo = class(TForm)
    Label1: TLabel;
    DXBSim: TdxButton;
    DXBNao: TdxButton;
    Label2: TLabel;
    Label3: TLabel;
    Label6: TLabel;
    N_Baixa: TEdit;
    Nomecli: TEdit;
    Memoobs: TMemo;
    EValor: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Adversas: TMemo;
    XBanner3: TXBanner;
    procedure FormShow(Sender: TObject);
    procedure DXBSimClick(Sender: TObject);
    procedure DXBNaoClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmImpRecibo: TFrmImpRecibo;

implementation

uses Tabelas, Funcoes, RelRecibodeQuita, uRuntimeFields;

{$R *.dfm}

procedure TFrmImpRecibo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FrmImpRecibo:=nil;
  Action:=CaFree;
end;

procedure TFrmImpRecibo.FormShow(Sender: TObject);
var
  varsoma : double;
begin
  if FrmRelRecibodeQuita=nil then
     FrmRelRecibodeQuita:=TFrmRelRecibodeQuita.Create(Application);

  FrmRelRecibodeQuita.ZQForma.SQL.Clear;
  FrmRelRecibodeQuita.ZQForma.SQL.Add('Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento from receb_baixa where refbaixa='+quotedstr(N_Baixa.Text));
  FrmRelRecibodeQuita.ZQForma.Open;
  varsoma := 0;
  FrmRelRecibodeQuita.ZQForma.DisableControls;
  while not FrmRelRecibodeQuita.ZQForma.Eof do begin
    varsoma := varsoma + FrmRelRecibodeQuita.ZQForma.FieldByName('Vr_rec').AsFloat;
    FrmRelRecibodeQuita.ZQForma.Next;
  end;
  FrmRelRecibodeQuita.ZQForma.EnableControls;
  FrmRelRecibodeQuita.ZQTit.SQL.Clear;
  FrmRelRecibodeQuita.ZQTit.SQL.Add('SELECT idrecbxhist,refer,idrecib,recbxhist.valor,descon,recbxhist.juros,percent_usado,data,idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,recebimento.Valor,Observ,VrDoc,');
  FrmRelRecibodeQuita.ZQTit.SQL.Add('       ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,adversa,recpag,numordem ');
  FrmRelRecibodeQuita.ZQTit.SQL.Add(' FROM recbxhist join recebimento on idrecib=idrecebimento');
  FrmRelRecibodeQuita.ZQTit.SQL.Add('where refer='+N_Baixa.Text);
  FrmRelRecibodeQuita.ZQTit.Open;
//  Adversas.Clear;
  FrmRelRecibodeQuita.ZQTit.DisableControls;
  while not FrmRelRecibodeQuita.ZQTit.Eof do begin
    if (not empty(FrmRelRecibodeQuita.ZQTit.FieldByName('nomeadversa').AsString)) and (pos(FrmRelRecibodeQuita.ZQTit.FieldByName('nomeadversa').AsString,Adversas.Text)=0)  Then
      Adversas.Lines.Add(FrmRelRecibodeQuita.ZQTit.FieldByName('nomeadversa').AsString);
    FrmRelRecibodeQuita.ZQTit.Next;
  end;
  FrmRelRecibodeQuita.ZQTit.EnableControls;
  FrmRelRecibodeQuita.ZQRateio.SQL.Clear;
  FrmRelRecibodeQuita.ZQRateio.SQL.Add('SELECT  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  FrmRelRecibodeQuita.ZQRateio.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem ');
  FrmRelRecibodeQuita.ZQRateio.SQL.Add(' FROM recebimento');
  FrmRelRecibodeQuita.ZQRateio.SQL.Add('where refvinda='+N_Baixa.Text);
  FrmRelRecibodeQuita.ZQRateio.Open;
  EValor.Text := trim(Transform(varsoma,'###,###,###,##0.00'));
end;

procedure TFrmImpRecibo.DXBSimClick(Sender: TObject);
begin
  FrmImpRecibo.FormStyle:=fsNormal;
  FrmRelRecibodeQuita.RLLabel39.Caption := N_Baixa.Text;
  FrmRelRecibodeQuita.RLLabel18.Caption := N_Baixa.Text;
  FrmRelRecibodeQuita.RLLabel19.Caption := EValor.Text;
  FrmRelRecibodeQuita.RLLabel54.Caption := EValor.Text;
  FrmRelRecibodeQuita.RLMemo1.Lines := Memoobs.Lines;
  FrmRelRecibodeQuita.RLMemo2.Lines := Adversas.Lines;
  FrmRelRecibodeQuita.RLMemo3.Lines := Adversas.Lines;
  FrmRelRecibodeQuita.RLReport1.PreviewModal;
  FrmRelRecibodeQuita:=nil;
  FrmImpRecibo.FormStyle:=fsStayOnTop;
  Close;
end;

procedure TFrmImpRecibo.DXBNaoClick(Sender: TObject);
begin
  Close;
end;


procedure TFrmImpRecibo.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
