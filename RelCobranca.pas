unit RelCobranca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLParser, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter, StdCtrls, XNum, Mask, XEdit;

type
  TFrmRelCobranca = class(TForm)
    RLBand1: TRLBand;
    RLLabel1: TRLLabel;
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
    RLBand5: TRLBand;
    XNsomabaixa: TXNumEdit;
    XNNegociado: TXNumEdit;
    XNRenova: TXNumEdit;
    XNPrincipal: TXNumEdit;
    RLLabel10: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel30: TRLLabel;
    RLDBText2: TRLDBText;
    RLDBText18: TRLDBText;
    RLDBText20: TRLDBText;
    RLLabel2: TRLLabel;
    RLBand2: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel31: TRLLabel;
    RLLabel32: TRLLabel;
    RLLabel33: TRLLabel;
    RLBand3: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBResult1: TRLDBResult;
    RLLabel15: TRLLabel;
    RLDBText4: TRLDBText;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelCobranca: TFrmRelCobranca;
  varvar : string;
implementation

uses Tabelas, Funcoes, Cad_Recebimento, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrmRelCobranca.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelCobranca.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

