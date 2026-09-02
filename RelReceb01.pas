unit RelReceb01;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, QRPDFFilt, QRXLSXFilt;

type
  TFrmRelReceb01 = class(TForm)
    QRRelReceb01: TQuickRep;
    QRBand1: TQRBand;
    QRLabel12: TQRLabel;
    QRSysData4: TQRSysData;
    QRSysData5: TQRSysData;
    QRSysData6: TQRSysData;
    QRDBImage1: TQRDBImage;
    QRDBText8: TQRDBText;
    QRLabel13: TQRLabel;
    QRLabel15: TQRLabel;
    DetailBand1: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    QRDBText9: TQRDBText;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRDBText10: TQRDBText;
    QRSubDetail1: TQRSubDetail;
    QRDBText11: TQRDBText;
    QRDBText12: TQRDBText;
    QRDBText13: TQRDBText;
    QRDBText14: TQRDBText;
    QRDBText15: TQRDBText;
    QRMemo1: TQRMemo;
    QRBand3: TQRBand;
    QRExpr1: TQRExpr;
    QRExpr2: TQRExpr;
    QRDBSomaBaix: TQRDBText;
    QRLSomaBaixa: TQRLabel;
    QRLabel9: TQRLabel;
    QRDBText16: TQRDBText;
    QRXMLSSFilter1: TQRXMLSSFilter;
    QRPDFFilter1: TQRPDFFilter;
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRSubDetail1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRRelReceb01BeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb01: TFrmRelReceb01;
  varsomabaixa : Double;

implementation

uses Tabelas, funcoes, RelRecebimento, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelReceb01.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DetailBand1.Color  := clWhite;
  if (FrmRelRecebimento.CBVerbaixa.ItemIndex=1) and (FrmRelRecebimento.CBZebrar.Checked) and (FrmRelrecebimento.ZQTempReceber.recno mod 2 = 0) then
    DetailBand1.Color  :=  $00DADADA;
  QRSubDetail1.Enabled := false;
  if FrmRelRecebimento.CBVerbaixa.ItemIndex=0 then
    DetailBand1.Frame.DrawTop := True
  else
    DetailBand1.Frame.DrawTop := false;
  if (FrmRelRecebimento.ZQBaixa.RecordCount>0) and (FrmRelRecebimento.CBVerbaixa.ItemIndex=0) Then
    QRSubDetail1.Enabled := True;
end;

procedure TFrmRelReceb01.QRSubDetail1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  varsomabaixa := varsomabaixa + FrmRelRecebimento.ZQBaixa.FieldByName('Vr_rec').AsFloat;
  QRLSomaBaixa.Caption := floattostr(varsomabaixa);
  if FrmRelRecebimento.ZQBaixa.RecordCount=FrmRelRecebimento.ZQBaixa.RecNo then Begin
    QRDBSomaBaix.Enabled := True;
  end
  else
    QRDBSomaBaix.Enabled := False;
end;

procedure TFrmRelReceb01.QRRelReceb01BeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
varsomabaixa := 0;
end;

procedure TFrmRelReceb01.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel12.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelReceb01.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
