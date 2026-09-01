unit RelReceb02;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLParser, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter, StdCtrls, XNum, Mask, XEdit;

type
  TFrmRelReceb02 = class(TForm)
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
    RLExpressionParser1: TRLExpressionParser;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLReport1: TRLReport;
    XNsomabaixa: TXNumEdit;
    XNNegociado: TXNumEdit;
    XNRenova: TXNumEdit;
    XNPrincipal: TXNumEdit;
    RLGroup1: TRLGroup;
    RLSubDetail1: TRLSubDetail;
    RLBand4: TRLBand;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLIm_recebeu: TRLImage;
    RLIm_Negociou: TRLImage;
    RLDBText19: TRLDBText;
    RLDBText14: TRLDBText;
    RLDBText15: TRLDBText;
    RLBand6: TRLBand;
    RLDBMemo2: TRLDBMemo;
    RLBand3: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBMemo3: TRLDBMemo;
    RLDBText16: TRLDBText;
    RLIreneg: TRLImage;
    RLDBText18: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText21: TRLDBText;
    RLBand2: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel9: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel30: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel31: TRLLabel;
    RLBand5: TRLBand;
    RLLabel22: TRLLabel;
    RLLabel20: TRLLabel;
    RLLabel26: TRLLabel;
    RLLabel27: TRLLabel;
    RLDBResult1: TRLDBResult;
    RLLabel19: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel23: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel28: TRLLabel;
    RLLabel29: TRLLabel;
    RLImage1: TRLImage;
    RLImage2: TRLImage;
    RLImage3: TRLImage;
    RLBand7: TRLBand;
    RLLabel32: TRLLabel;
    RLDBText22: TRLDBText;
    RLDraw1: TRLDraw;
    RLBand8: TRLBand;
    RLDBResult2: TRLDBResult;
    RLDraw2: TRLDraw;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand5BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02: TFrmRelReceb02;
  varvar : string;
implementation

uses Tabelas, Funcoes, RelRecebimento, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrmRelReceb02.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand4.Visible := False;
  RLBand7.Visible := False;
  if FrmRelRecebimento.CBVerbaixa.Text = 'Sim' Then Begin
    RLBand4.Visible := True;
    if FrmRelRecebimento.CBOBSBaix.Checked Then
      RLBand7.Visible := True;
  end;
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;  
end;

procedure TFrmRelReceb02.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  //if not empty(FrmRelRecebimento.ZQTempReceberRefVinda.Text) Then Begin
  if FrmRelRecebimento.ZQTempReceber.FieldByName('refvinda').Value>0 Then Begin
    RLIreneg.Visible := True;
    XNRenova.Value := XNRenova.Value + FrmRelRecebimento.ZQTempReceber.FieldByName('Valor').Value;
  end
  else Begin
    RLIreneg.Visible := False;
    XNPrincipal.Value := XNPrincipal.Value + FrmRelRecebimento.ZQTempReceber.FieldByName('Valor').Value;
  end;

  RLDBMemo3.Visible := False;
  RLBand3.Height := 16;
  if FrmRelRecebimento.CBObsLanc.Checked Then Begin
    RLDBMemo3.Visible := True;
    RLBand3.Height := 32;
  end;
  if not empty(FrmRelRecebimento.ZQTempReceber.FieldByName('nomeempreend').Value) then
  begin
    RLDBText2.DataField:='nomeempreend';
  end
  else
  begin
    RLDBText2.DataField:='nome_parte';  
  end;



  if FrmRelRecebimento.ZQTempReceber.FieldByName('saldo').value>0  then
  begin
    RLDBText8.DataField:='saldo';
  end
  else
  begin
    RLDBText8.DataField:='valor';
  end;


end;

procedure TFrmRelReceb02.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if FrmRelRecebimento.ZQBaixa.FieldByName('receb_receb').Value = 'S' Then Begin
    RLIm_recebeu.Visible := False;
    RLIm_Negociou.Visible := True;
    XNNegociado.Value := XNNegociado.Value + FrmRelRecebimento.ZQBaixa.FieldByName('vrusado').Value;
  end
  else Begin
    RLIm_recebeu.Visible := True;
    RLIm_Negociou.Visible := False;
    XNsomabaixa.Value := XNsomabaixa.Value + FrmRelRecebimento.ZQBaixa.FieldByName('vrusado').Value;
    varvar := varvar+'I' + FrmRelRecebimento.ZQBaixa.FieldByName('Vr_rec').Text;
  end;
end;

procedure TFrmRelReceb02.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel23.Caption := Transform(XNRenova.Value,'###,###,###,##0.00');
  RLLabel25.Caption := Transform(XNPrincipal.Value,'###,###,###,##0.00');
  RLLabel21.Caption := Transform(XNNegociado.Value,'###,###,###,##0.00');
  RLLabel19.Caption := Transform(XNsomabaixa.Value,'###,###,###,##0.00');
end;


procedure TFrmRelReceb02.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

