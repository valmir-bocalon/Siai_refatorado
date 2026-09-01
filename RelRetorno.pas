unit RelRetorno;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLFilters, RLPDFFilter, RLRichFilter, RLXLSFilter;

type
  TFrm_RelRetorno = class(TForm)
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
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
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
  Frm_RelRetorno: TFrm_RelRetorno;

implementation

uses BaixaAutomatica, tabelas, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrm_RelRetorno.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
//  if Frm_BaixaAutomatica.CDSArqRetornoacao.Value='R' Then
//    DM_TAbelas.ImageList1.GetBitmap(2,RLIm_Negociou.Picture);
  RLBand2.Color  := clWhite;
  if Frm_BaixaAutomatica.CDSArqRetorno.recno mod 2 = 0 then
    RLBand2.Color  := $00EAEAEA;

end;

procedure TFrm_RelRetorno.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelRetorno.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
