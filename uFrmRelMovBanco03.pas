unit uFrmRelMovBanco03;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport;

type
  TFrmRelMovBanco03 = class(TForm)
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
    RLBand2: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLBand3: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel13: TRLLabel;
    RLDBObs: TRLDBMemo;
    RLLabel14: TRLLabel;
    RLDBText3: TRLDBText;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelMovBanco03: TFrmRelMovBanco03;

implementation

uses tabelas, UFrmRelMovBanco, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelMovBanco03.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLDBText3.Visible := false;
  if FrmRelMovBanco.CBMostrasaldo.Checked then
    RLDBText3.Visible := True;
  RLDBObs.Visible := False;
  if FrmRelMovBanco.CBMostraObs.Checked then
    RLDBObs.Visible := True;
  RLLabel2.Caption := 'Per�odo de '+Datetostr(FrmRelMovBanco.XDEInicial.Date)+' � '+Datetostr(FrmRelMovBanco.XDEFinal.Date)+' da date de ';
  if FrmrelMovBanco.RGOrdem.ItemIndex = 0 Then Begin
    RLLabel2.Caption := RLLabel2.Caption +'Lan�amento';
    RLDBText6.DataField := 'dt_lanc';
    RLDBText3.DataField := 'Saldo_lanc';
    if FrmRelMovBanco.CBSomalimite.Checked Then
      RLDBText3.DataField := 'Saldo_lanc_plus';
  end
  else
  Begin
    RLLabel2.Caption := RLLabel2.Caption +'Compensa��o';
    RLDBText6.DataField := 'dt_conciliado';
    RLDBText3.DataField := 'Saldo_conci';
    if FrmRelMovBanco.CBSomalimite.Checked Then
      RLDBText3.DataField := 'saldo_Conci_plus';
  end;
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelMovBanco03.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand2.Color := ClWhite;
  if (frmrelmovbanco.CBZebrado.Checked) and (frmrelmovbanco.ZQTempMovBan.RecNo mod 2 = 0) Then
    RLBand2.Color := ClSilver;
end;


procedure TFrmRelMovBanco03.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
