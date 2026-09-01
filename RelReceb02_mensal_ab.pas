unit RelReceb02_mensal_ab;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelReceb02_mensal_ab = class(TForm)
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
    RLBand2: TRLBand;
    RLDBResult2: TRLDBResult;
    RLGroup1: TRLGroup;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLDBText3: TRLDBText;
    RLDraw1: TRLDraw;
    RLBand6: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel17: TRLLabel;
    RLDraw2: TRLDraw;
    RLBand3: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText8: TRLDBText;
    RLBand5: TRLBand;
    RLDBResult1: TRLDBResult;
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02_mensal_ab: TFrmRelReceb02_mensal_ab;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, RelRecebimento, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelReceb02_mensal_ab.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand3.Color  := clWhite;
  if FrmRelRecebimento.ZQrelatorio2.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;

end;

procedure TFrmRelReceb02_mensal_ab.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelReceb02_mensal_ab.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
