unit RelCobranca2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLParser, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter, StdCtrls, XNum, Mask, XEdit, DB, DBClient;

type
  TFrmRelCobranca2 = class(TForm)
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
    RLLabel2: TRLLabel;
    RLBand2: TRLBand;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel31: TRLLabel;
    RLLabel32: TRLLabel;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBResult1: TRLDBResult;
    RLDBText2: TRLDBText;
    RLDBText18: TRLDBText;
    RLDBText20: TRLDBText;
    RLLabel10: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel30: TRLLabel;
    RLLabel9: TRLLabel;
    RLDBText1: TRLDBText;
    RLLabel13: TRLLabel;
    RLDBText4: TRLDBText;
    RLDBText6: TRLDBText;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLDBText9: TRLDBText;
    RLDBText11: TRLDBText;
    RLLabel19: TRLLabel;
    RLDBText12: TRLDBText;
    RLDBResult2: TRLDBResult;
    RLDBResult3: TRLDBResult;
    RLDBResult4: TRLDBResult;
    RLLabel20: TRLLabel;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelCobranca2: TFrmRelCobranca2;
  varvar : string;
implementation

uses Tabelas, Funcoes, Cad_Recebimento, principal, uRuntimeFields, uSiaiReportPerformance;
{$R *.dfm}

procedure TFrmRelCobranca2.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelCobranca2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  FrmRelCobranca2:=nil;
  Action:=CaFree;
  
end;


procedure TFrmRelCobranca2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
  MeasureReport(RLReport1, 'Relatorio do grid');
end;

end.
