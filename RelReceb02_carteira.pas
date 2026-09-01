unit RelReceb02_carteira;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelReceb02_carteira = class(TForm)
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
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLDBText3: TRLDBText;
    RLBand6: TRLBand;
    RLBand3: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText8: TRLDBText;
    RLBand5: TRLBand;
    RLDBResult1: TRLDBResult;
    RLLabel10: TRLLabel;
    RLLabel17: TRLLabel;
    RLDraw1: TRLDraw;
    RLDraw2: TRLDraw;
    RLBand2: TRLBand;
    RLDBResult2: TRLDBResult;
    RLLabel11: TRLLabel;
    RLDBText4: TRLDBText;
    RLDBResult3: TRLDBResult;
    RLDBResult4: TRLDBResult;
    RLLabel12: TRLLabel;
    RLDBText5: TRLDBText;
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02_carteira: TFrmRelReceb02_carteira;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, RelRecebimento, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelReceb02_carteira.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand3.Color  := clWhite;
  if FrmRelRecebimento.ZQRelCarteira.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;

end;

procedure TFrmRelReceb02_carteira.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelReceb02_carteira.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
