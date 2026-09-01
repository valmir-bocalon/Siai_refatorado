unit UFrmRelMovBanco01;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, DB;

type
  TFrmRelMovBanco01 = class(TForm)
    QRmovbancario01: TQuickRep;
    DetailBand1: TQRBand;
    QRDData: TQRDBText;
    QRDValor: TQRDBText;
    QRDDocumento: TQRDBText;
    QRDCD: TQRDBText;
    ColumnHeaderBand1: TQRBand;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRDConta: TQRDBText;
    QEDMascara: TQRDBText;
    QRDBText1: TQRDBText;
    QRLSaldo: TQRLabel;
    QRLObs: TQRLabel;
    RDRMemoobs: TQRDBRichText;
    QRSLimite: TQRSubDetail;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
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
    procedure QRmovbancario01BeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure DetailBand1AfterPrint(Sender: TQRCustomBand;
      BandPrinted: Boolean);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelMovBanco01: TFrmRelMovBanco01;

implementation

uses Tabelas, funcoes, UFrmRelMovBanco, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelMovBanco01.QRmovbancario01BeforePrint(
  Sender: TCustomQuickRep; var PrintReport: Boolean);
begin
  QRSLimite.Enabled := false;
  QRDBab1.Caption := 'Conta: '+DM_tabelas.ZQConta_Bancariaconta.Value+'   Agencia: '+DM_tabelas.ZQConta_BancariaAgencia.Value+'  Banco: '+DM_tabelas.ZQConta_Bancarianome_banco.Value;
  if FrmRelMovBanco.RGOrdem.ItemIndex = 0 then begin
    QRDBab2.Caption := 'Per�odo de '+datetostr(FrmRelMovBanco.XDEInicial.Date)+' � '+datetostr(FrmRelMovBanco.XDEFinal.Date)+' da data de lan�amento';
    QRDData.DataField := 'Dt_lanc';
  end
  else
  begin
    QRDBab2.Caption := 'Per�odo de '+datetostr(FrmRelMovBanco.XDEInicial.Date)+' � '+datetostr(FrmRelMovBanco.XDEFinal.Date)+' da data de Compensa��o';
    QRDData.DataField := 'Dt_conciliado';
  end;
  if FrmRelMovBanco.CBSomalimite.Checked then Begin
    DM_tabelas.ZQLimiteEspecial.SQL.Clear;
    DM_tabelas.ZQLimiteEspecial.SQL.Add('select * from limite_especial where conta_bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
    DM_tabelas.ZQLimiteEspecial.Open;
  end;
  DetailBand1.Height := 15;
  QRLObs.Enabled := false;
  RDRMemoobs.Enabled := False;
  DetailBand1.Frame.DrawBottom := false;
  if FrmRelMovBanco.CBMostraObs.Checked then Begin
    DetailBand1.Height := 60;
    QRLObs.Enabled := True;
    RDRMemoobs.Enabled := True;
    DetailBand1.Frame.DrawBottom := True;
  end;
end;

procedure TFrmRelMovBanco01.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
Var
  varlimite : Double;
begin
  DetailBand1.Color  := clWhite;
  if (FrmRelMovBanco.CBZebrado.Checked) and (FrmRelMovBanco.ZQTempMovBan.recno mod 2 = 0) then
    DetailBand1.Color  := $00EAEAEA;
  DM_tabelas.ZQPlanoDeContas.Locate('codigo',FrmRelMovBanco.ZQTempMovBan.FieldByName('Plano_Contas_codigo').Value,[]);
  QRLSaldo.Caption := '';
  if FrmRelMovBanco.CBMostrasaldo.Checked then Begin
    QRLSaldo.Caption := floattostr(FrmRelMovBanco.ZQTempMovBan.FieldByName('Saldo_conci').Value);
    varlimite := 0;
    if FrmRelMovBanco.CBSomalimite.Checked then Begin
      DM_tabelas.ZQLimiteEspecial.SQL.Clear;
      DM_tabelas.ZQLimiteEspecial.SQL.Add('select * from limite_especial where (conta_bancaria_cod_banco='+DM_tabelas.ZQConta_Bancariacod_banco.Text);
      DM_tabelas.ZQLimiteEspecial.SQL.Add(') and (data_valida<= :dt1)');
      DM_tabelas.ZQLimiteEspecial.ParamByName('dt1').AsDate:=strtodate(FrmRelMovBanco.ZQTempMovBan.FieldByName('dt_conciliado').text);
      DM_tabelas.ZQLimiteEspecial.Open;
      DM_tabelas.ZQLimiteEspecial.Last;
      varlimite := DM_tabelas.ZQLimiteEspecialvalor.Value;
    end;
    if FrmRelMovBanco.RGOrdem.ItemIndex = 0 then
      QRLSaldo.Caption := transform(FrmRelMovBanco.ZQTempMovBan.FieldByName('Saldo_lanc').Value+varlimite,'###,###,###,##0.00')
    else
      QRLSaldo.Caption := transform(FrmRelMovBanco.ZQTempMovBan.FieldByName('Saldo_conci').Value+varlimite,'###,###,###,##0.00');
  end;

end;

procedure TFrmRelMovBanco01.DetailBand1AfterPrint(Sender: TQRCustomBand;
  BandPrinted: Boolean);
begin
if DM_tabelas.ZQMovBancaria.RecNo=DM_tabelas.ZQMovBancaria.RecordCount Then
    QRSLimite.Enabled := True;
end;

procedure TFrmRelMovBanco01.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DM_tabelas.ZQEmpresa.Open;
  DM_tabelas.ZQEmpresa.First;
  QRLabel12.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelMovBanco01.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
