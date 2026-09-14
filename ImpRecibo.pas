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

uses Tabelas, Funcoes, RelRecibodeQuita, uRuntimeFields, ZDataset;

{$R *.dfm}

procedure TFrmImpRecibo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(FrmRelRecibodeQuita);
  FrmImpRecibo:=nil;
  Action:=CaFree;
end;

procedure TFrmImpRecibo.FormShow(Sender: TObject);
var
  varsoma : double;
  LReferencia: Integer;
  LPartes: TZQuery;
  LNomeParte: string;
begin
  if not TryStrToInt(Trim(N_Baixa.Text), LReferencia) then
  begin
    ShowMessage('Referencia de baixa invalida para a reimpressao do recibo.');
    Close;
    Exit;
  end;

  if FrmRelRecibodeQuita=nil then
     FrmRelRecibodeQuita:=TFrmRelRecibodeQuita.Create(Application);

  Adversas.Clear;
  FrmRelRecibodeQuita.ZQForma.Close;
  FrmRelRecibodeQuita.ZQForma.SQL.Clear;
  FrmRelRecibodeQuita.ZQForma.SQL.Add('Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento from receb_baixa where refbaixa=:referencia');
  FrmRelRecibodeQuita.ZQForma.ParamByName('referencia').AsInteger := LReferencia;
  FrmRelRecibodeQuita.ZQForma.Open;
  varsoma := 0;
  FrmRelRecibodeQuita.ZQForma.DisableControls;
  while not FrmRelRecibodeQuita.ZQForma.Eof do begin
    varsoma := varsoma + FrmRelRecibodeQuita.ZQForma.FieldByName('Vr_rec').AsFloat;
    FrmRelRecibodeQuita.ZQForma.Next;
  end;
  FrmRelRecibodeQuita.ZQForma.EnableControls;
  FrmRelRecibodeQuita.ZQForma.First;
  FrmRelRecibodeQuita.ZQTit.Close;
  FrmRelRecibodeQuita.ZQTit.SQL.Clear;
  FrmRelRecibodeQuita.ZQTit.SQL.Add('SELECT H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,R.idrecebimento,R.documento,R.cliente,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.Observ,R.VrDoc,');
  FrmRelRecibodeQuita.ZQTit.SQL.Add('       R.ordem,R.TipDoc,R.saldo,R.marcar,R.RefBaixa,R.refvinda,R.contabil,R.empresa,R.custodaparcela,R.origem,R.adversa,R.recpag,R.numordem,R.nomeadversa ');
  FrmRelRecibodeQuita.ZQTit.SQL.Add('FROM recbxhist H join recebimento R on H.idrecib=R.idrecebimento ');
  FrmRelRecibodeQuita.ZQTit.SQL.Add('where H.refer=:referencia');
  FrmRelRecibodeQuita.ZQTit.ParamByName('referencia').AsInteger := LReferencia;
  FrmRelRecibodeQuita.ZQTit.Open;
  FrmRelRecibodeQuita.ZQTit.First;

  // Use a local query because ZQTit's nomeadversa is a persistent lookup
  // field. A direct result avoids that lookup replacing a valid name by blank.
  LPartes := TZQuery.Create(nil);
  try
    LPartes.Connection := DM_Tabelas.zconeccao;
    LPartes.SQL.Text :=
      'select distinct coalesce(nullif(P.nome_parte, ''''), ' +
      'nullif(R.nomeadversa, '''')) as nomeparte ' +
      'from recbxhist H ' +
      'join recebimento R on R.idrecebimento=H.idrecib ' +
      'left join participante P on P.idpaticipante=R.adversa ' +
      'where H.refer=:referencia order by nomeparte';
    LPartes.ParamByName('referencia').AsInteger := LReferencia;
    LPartes.Open;
    while not LPartes.Eof do
    begin
      LNomeParte := Trim(LPartes.FieldByName('nomeparte').AsString);
      if (LNomeParte <> '') and (Adversas.Lines.IndexOf(LNomeParte) < 0) then
        Adversas.Lines.Add(LNomeParte);
      LPartes.Next;
    end;
  finally
    LPartes.Free;
  end;
  FrmRelRecibodeQuita.ZQRateio.Close;
  FrmRelRecibodeQuita.ZQRateio.SQL.Clear;
  FrmRelRecibodeQuita.ZQRateio.SQL.Add('SELECT  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  FrmRelRecibodeQuita.ZQRateio.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem ');
  FrmRelRecibodeQuita.ZQRateio.SQL.Add(' FROM recebimento');
  FrmRelRecibodeQuita.ZQRateio.SQL.Add('where refvinda=:referencia');
  FrmRelRecibodeQuita.ZQRateio.ParamByName('referencia').AsInteger := LReferencia;
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
