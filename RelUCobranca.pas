unit RelUCobranca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLFilters, RLPDFFilter, RLRichFilter, RLXLSFilter;

type
  TFrm_RelUCobranca = class(TForm)
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
    RLBand2: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLPDFFilter1: TRLPDFFilter;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLDraw1: TRLDraw;
    RLLabel9: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel19: TRLLabel;
    RLLabel20: TRLLabel;
    RLDBText14: TRLDBText;
    RLLabel21: TRLLabel;
    RLDBText15: TRLDBText;
    RLDraw2: TRLDraw;
    RLDBText5: TRLDBText;
    RLLabel22: TRLLabel;
    RLLabel23: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel26: TRLLabel;
    RLLabel27: TRLLabel;
    RLLabel28: TRLLabel;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText16: TRLDBText;
    RLBand3: TRLBand;
    RLDraw3: TRLDraw;
    RLDBResult1: TRLDBResult;
    RLDBResult2: TRLDBResult;
    RLDBResult3: TRLDBResult;
    RLDBResult4: TRLDBResult;
    RLDBResult5: TRLDBResult;
    RLLabel11: TRLLabel;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelUCobranca: TFrm_RelUCobranca;

implementation

uses tabelas, Ucobranca,Principal, uRuntimeFields;
{$R *.dfm}

procedure TFrm_RelUCobranca.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
//  if Frm_BaixaAutomatica.CDSArqRetornoacao.Value='R' Then
//    DM_TAbelas.ImageList1.GetBitmap(2,RLIm_Negociou.Picture);
  RLBand2.Color  := clWhite;
  if frmcobranca.CDSCobranca.recno mod 2 = 0 then
    RLBand2.Color  := $00EAEAEA;

end;

procedure TFrm_RelUCobranca.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelUCobranca.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
