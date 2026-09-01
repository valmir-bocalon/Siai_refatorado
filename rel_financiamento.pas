unit rel_financiamento;

interface

uses
  Windows, Messages,
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, Db, DBTables,
  RLReport, RLDraftFilter, RLFilters, RLPDFFilter, RLRichFilter, RLXLSFilter;

type
  Tfrm_financiamento = class(TForm)
    RLReport1: TRLReport;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLBand3: TRLBand;
    RLBand4: TRLBand;
    RLBand1: TRLBand;
    RLSystemInfo1: TRLSystemInfo;
    RLBand5: TRLBand;
    RLSystemInfo2: TRLSystemInfo;
    RLBand6: TRLBand;
    RLLabel1: TRLLabel;
    RLLabel2: TRLLabel;
    RLLabel5: TRLLabel;
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
    RLLabel15: TRLLabel;
    RLLabel10: TRLLabel;
    RLDraw1: TRLDraw;
    RLDBText18: TRLDBText;
    RLDBText20: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBResult1: TRLDBResult;
    RLDBResult3: TRLDBResult;
    RLLabel3: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frm_financiamento: Tfrm_financiamento;

implementation


uses Tabelas, Funcoes, RelRecebimento, principal, uRuntimeFields;


{$R *.dfm}

procedure Tfrm_financiamento.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand4.Color := clWhite;
  if (FrmRelRecebimento.CBZebrar.Checked) and (FrmRelrecebimento.CDS_finan.recno mod 2 = 0) then
      RLBand4.Color := $00E1E1E1;
end;

procedure Tfrm_financiamento.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure Tfrm_financiamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
