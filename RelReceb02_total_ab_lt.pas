unit RelReceb02_total_ab_lt;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelReceb02_total_ab_lt = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLLabel1: TRLLabel;
    RLLabel2: TRLLabel;
    RLLabel3: TRLLabel;
    RLLabel4: TRLLabel;
    RLSystemInfo1: TRLSystemInfo;
    RLSystemInfo2: TRLSystemInfo;
    RLLabel5: TRLLabel;
    RLLabel6: TRLLabel;
    RLLabel7: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLLabel8: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    XNsomabaixa: TXNumEdit;
    XNNegociado: TXNumEdit;
    XNRenova: TXNumEdit;
    XNPrincipal: TXNumEdit;
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLBand5: TRLBand;
    RLDBResult1: TRLDBResult;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLDBText7: TRLDBText;
    RLDraw1: TRLDraw;
    RLBand6: TRLBand;
    RLDBResult2: TRLDBResult;
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }
    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02_total_ab_lt: TFrmRelReceb02_total_ab_lt;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, RelRecebimento, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelReceb02_total_ab_lt.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if not empty(FrmRelRecebimento.ZQTempReceber2.FieldByName('refvinda').Text) Then Begin
    XNRenova.Value := XNRenova.Value + FrmRelRecebimento.ZQTempReceber2.FieldByName('Valor').AsFloat;
  end
  else Begin
    XNPrincipal.Value := XNPrincipal.Value + FrmRelRecebimento.ZQTempReceber2.FieldByName('Valor').AsFloat;
  end;

  RLBand3.Height := 16;
  if FrmRelRecebimento.CBObsLanc.Checked Then Begin
    RLBand3.Height := 32;
  end;
{  if not empty(FrmRelRecebimento.ZQTempReceber2nomeempreend.Value) then
  begin
    RLDBText2.DataField:='nomeempreend';
  end
  else
  begin
    RLDBText2.DataField:='nome_parte';
  end;}
  RLBand3.Color  := clWhite;
  if FrmRelRecebimento.ZQTempReceber2.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;

end;

procedure TFrmRelReceb02_total_ab_lt.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if FrmRelRecebimento.ZQBaixa.FieldByName('receb_receb').AsString = 'S' Then Begin
    XNNegociado.Value := XNNegociado.Value + FrmRelRecebimento.ZQBaixa.FieldByName('vrusado').AsFloat;
  end
  else Begin
    XNsomabaixa.Value := XNsomabaixa.Value + FrmRelRecebimento.ZQBaixa.FieldByName('vrusado').AsFloat;
    varvar := varvar+'I' + FrmRelRecebimento.ZQBaixa.FieldByName('Vr_rec').Text;
  end;
end;

procedure TFrmRelReceb02_total_ab_lt.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelReceb02_total_ab_lt.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
  { Use os nomes retornados pela consulta, sem recalcular lookups ao imprimir. }
  RLGroup1.DataFields := 'nome_loteamento';
  RLDBText7.DataField := 'nome_loteamento';
  RLDBText6.DataField := 'comprador_relatorio';
end;

end.
