unit UFrmRelMovBanco02;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, QRCtrls, QuickRpt, ExtCtrls;

type
  TFrmRelMovBanco02 = class(TForm)
    QRmovbancario02: TQuickRep;
    QRBand1: TQRBand;
    QRLabel12: TQRLabel;
    QRSysData4: TQRSysData;
    QRSysData5: TQRSysData;
    QRSysData6: TQRSysData;
    QRDBImage1: TQRDBImage;
    QRDBText8: TQRDBText;
    QRLabel13: TQRLabel;
    QRLabel15: TQRLabel;
    QRDBab2: TQRLabel;
    QRDBab1: TQRLabel;
    DetailBand1: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRDBText6: TQRDBText;
    QRSubDetail1: TQRSubDetail;
    QRDBText7: TQRDBText;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    QRDBText12: TQRDBText;
    QRDBText13: TQRDBText;
    QRDBText14: TQRDBText;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelMovBanco02: TFrmRelMovBanco02;

implementation

uses Tabelas, funcoes, UFrmRelMovBanco, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelMovBanco02.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DM_tabelas.ZQEmpresa.Open;
  DM_tabelas.ZQEmpresa.First;
  QRDBab1.Caption := 'Conta: '+DM_tabelas.ZQConta_Bancariaconta.Value+'   Agencia: '+DM_tabelas.ZQConta_BancariaAgencia.Value+'  Banco: '+DM_tabelas.ZQConta_Bancarianome_banco.Value;
//  if FrmRelMovBanco.RGOrdem.ItemIndex = 0 then begin
    QRDBab2.Caption := 'Per�odo de '+datetostr(FrmRelMovBanco.XDEInicial.Date)+' � '+datetostr(FrmRelMovBanco.XDEFinal.Date); //+' da data de lan�amento';
  QRLabel12.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelMovBanco02.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DetailBand1.Color  := clWhite;
  if (FrmRelMovBanco.CBZebrado.Checked) and (FrmRelMovBanco.ZQtempPlano.recno mod 2 = 0) then
    DetailBand1.Color  := $00EAEAEA;
  if (FrmRelMovBanco.CBMostLanc.Checked) and (FrmRelMovBanco.ZQTempLanc.RecordCount>0) Then
    QRSubDetail1.Enabled := True
  else
    QRSubDetail1.Enabled := False;
end;


procedure TFrmRelMovBanco02.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
