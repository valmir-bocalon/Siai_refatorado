unit Relparticipante;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelparticipante = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
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
    RLLabel10: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    XNPrincipal: TXNumEdit;
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLSubDetail1: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText9: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLBand4: TRLBand;
    Observacao: TRLMemo;
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelparticipante: TFrmRelparticipante;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, Participante, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelparticipante.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand3.Color  := clWhite;
  if Frm_Participante.ZQContratos.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;

  if Frm_Participante.ZQContratos.FieldByName('observacao').AsString<>'' then
  begin
     Observacao.Lines.Clear;
     Observacao.Lines.Add(Frm_Participante.ZQContratos.FieldByName('observacao').AsString);
  end
  else
  begin
     Observacao.Lines.Clear;
  end;   
end;


procedure TFrmRelparticipante.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
