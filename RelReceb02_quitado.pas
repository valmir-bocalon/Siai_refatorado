unit RelReceb02_quitado;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelReceb02_quitado = class(TForm)
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
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLDraw7: TRLDraw;
    RLDraw1: TRLDraw;
    RLBand6: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText3: TRLDBText;
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
  FrmRelReceb02_quitado: TFrmRelReceb02_quitado;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, RelRecebimento, RelVenda, principal, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelReceb02_quitado.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  Frm_RelVenda.ZQrelquitados.Close;
end;

procedure TFrmRelReceb02_quitado.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand6.Color  := clWhite;
  if Frm_RelVenda.ZQrelquitados.recno mod 2 = 0 then
     RLBand6.Color  := $00EAEAEA;
  if Frm_RelVenda.ZQrelquitados.FieldByName('dataquitacao').AsDateTime>0 then
     RLDBText4.DataField:='dataquitacao'
  else
     RLDBText4.DataField:='datavenda';

end;

procedure TFrmRelReceb02_quitado.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelReceb02_quitado.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
