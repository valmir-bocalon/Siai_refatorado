unit RelReceb02_rescisao;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelReceb02_rescisao = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
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
    RLBand5: TRLBand;
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLLabel10: TRLLabel;
    RLDBText2: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLDraw7: TRLDraw;
    RLDraw1: TRLDraw;
    RLBand6: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLBand7: TRLBand;
    RLSystemInfo5: TRLSystemInfo;
    RLDraw2: TRLDraw;
    RLLabel15: TRLLabel;
    RLLabel14: TRLLabel;
    RLSystemInfo6: TRLSystemInfo;
    RLDraw3: TRLDraw;
    RLLabel16: TRLLabel;
    RLDBText6: TRLDBText;
    RLLabel11: TRLLabel;
    RLDBText3: TRLDBText;
    procedure RLBand5BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReceb02_rescisao: TFrmRelReceb02_rescisao;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, RelRecebimento, RelVenda, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelReceb02_rescisao.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  Frm_RelVenda.ZQRescisao.Close;
end;

procedure TFrmRelReceb02_rescisao.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand6.Color  := clWhite;
  if Frm_RelVenda.ZQRescisao.recno mod 2 = 0 then
    RLBand6.Color  := $00EAEAEA;
end;

procedure TFrmRelReceb02_rescisao.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelReceb02_rescisao.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
