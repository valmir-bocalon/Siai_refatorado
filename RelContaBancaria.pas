unit RelContaBancaria;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter;

type
  TFrm_RelBancaria = class(TForm)
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
    RLLabel9: TRLLabel;
    RLDBText2: TRLDBText;
    RLLabel10: TRLLabel;
    RLLabel12: TRLLabel;
    RLDBText3: TRLDBText;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLLabel16: TRLLabel;
    RLDBText7: TRLDBText;
    RLLabel17: TRLLabel;
    RLDBText8: TRLDBText;
    RLLabel18: TRLLabel;
    RLDBText9: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLLabel11: TRLLabel;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLLabel15: TRLLabel;
    RLDBText6: TRLDBText;
    RLDBText14: TRLDBText;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLReport1AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelBancaria: TFrm_RelBancaria;

implementation

uses RelRecebimento, funcoes, Tabelas, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrm_RelBancaria.RLReport1AfterPrint(Sender: TObject);
begin
  DM_Tabelas.ZQContaBancaria.Filtered:=false;
end;

procedure TFrm_RelBancaria.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelBancaria.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

