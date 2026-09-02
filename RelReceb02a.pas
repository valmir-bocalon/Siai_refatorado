unit RelReceb02a;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLParser, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter;

type
  TFrmRelReceb02 = class(TForm)
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
    RLSubDetail2: TRLSubDetail;
    RLBand5: TRLBand;
    RLDBText16: TRLDBText;
    RLDBText18: TRLDBText;
    RLBand6: TRLBand;
    RLLabel18: TRLLabel;
    RLLabel19: TRLLabel;
    RLLabel20: TRLLabel;
    RLExpressionParser1: TRLExpressionParser;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel9: TRLLabel;
    RLBand3: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText17: TRLDBText;
    RLBand4: TRLBand;
    RLDBResult1: TRLDBResult;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02: TFrmRelReceb02;

implementation

uses Tabelas, UFuncoes, RelRecebimento, uRuntimeFields;
{$R *.dfm}

procedure TFrmRelReceb02.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand4.Visible := False;
  if FrmRelRecebimento.CBVerbaixa.Text = 'Sim' Then
    RLBand4.Visible := True;
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelReceb02.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLSubDetail2.Visible := False;
  if FrmRelRecebimento.ZQTempReceber.RecNo = FrmRelRecebimento.ZQTempReceber.RecordCount Then
    RLSubDetail2.Visible := True;
end;


procedure TFrmRelReceb02.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
