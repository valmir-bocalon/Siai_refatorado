unit rel_SimulacaoAditamento;

interface

uses
  Windows, Messages,
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, Db, DBTables,
  RLReport, RLDraftFilter, RLFilters, RLPDFFilter, RLXLSFilter, RLRichFilter;

type
  Tfrm_SimulacaoAditamento = class(TForm)
    RLReport1: TRLReport;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLBand3: TRLBand;
    RLBand4: TRLBand;
    RLBand5: TRLBand;
    RLSystemInfo2: TRLSystemInfo;
    RLBand6: TRLBand;
    RLLabel1: TRLLabel;
    RLLabel2: TRLLabel;
    RLDBImage1: TRLDBImage;
    RLLabel7: TRLLabel;
    RLSystemInfo3: TRLSystemInfo;
    RLLabel8: TRLLabel;
    RLSystemInfo4: TRLSystemInfo;
    RLSystemInfo5: TRLSystemInfo;
    RLLabel14: TRLLabel;
    RLSystemInfo6: TRLSystemInfo;
    RLLabel24: TRLLabel;
    RLLabel30: TRLLabel;
    RLLabel10: TRLLabel;
    RLDraw1: TRLDraw;
    RLDBText18: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText2: TRLDBText;
    RLLabel3: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLSubDetail1: TRLSubDetail;
    RLDraw2: TRLDraw;
    RLDraw3: TRLDraw;
    RLLabel4: TRLLabel;
    RLLabel5: TRLLabel;
    RLLabel6: TRLLabel;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLBand7: TRLBand;
    RLDBText1: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLRichFilter1: TRLRichFilter;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand7BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frm_SimulacaoAditamento: Tfrm_SimulacaoAditamento;

implementation


uses Tabelas, Funcoes, principal, Venda, usimulaAditamento, uRuntimeFields;


{$R *.dfm}

procedure Tfrm_SimulacaoAditamento.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;

procedure Tfrm_SimulacaoAditamento.RLBand7BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand7.Color := clWhite;
  if frmsimulaaditamento.CdsAditamento.recno mod 2 = 0 then
      RLBand7.Color := $00E1E1E1;
end;


procedure Tfrm_SimulacaoAditamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
