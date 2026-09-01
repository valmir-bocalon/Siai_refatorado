unit RelImovelREMESSA;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLXLSFilter,
  RLRichFilter;

type
  TRLReportREMESSA = class(TForm)
    RLReportREMESSA: TRLReport;
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
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel16: TRLLabel;
    RLBand2: TRLBand;
    RLDBResult1: TRLDBResult;
    RLDraw1: TRLDraw;
    RLLabel15: TRLLabel;
    RLLabel17: TRLLabel;
    rtiti: TRLLabel;
    RLDBText9: TRLDBText;
    RLLabel18: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLRichFilter1: TRLRichFilter;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReportREMESSABeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure RLBand2AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  RLReportREMESSA: TRLReportREMESSA;

implementation

uses tabelas, RelImovel, funcoes, RelREMESSA, principal, uRuntimeFields;
{$R *.dfm}

procedure TRLReportREMESSA.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReportREMESSA.NewPage;
end;

procedure TRLReportREMESSA.RLReportREMESSABeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
 rtiti.Caption:=inttostr(DM_Tabelas.ZQRemes_Receb2.RecordCount);
 DM_Tabelas.ZQLoteamento.open;
 DM_Tabelas.ZQLoteamento.Filtered:=false;
 DM_Tabelas.ZQLoteamento.Filter:='idloteamento='+quotedstr(DM_Tabelas.ZQRemes_Receb2.FieldByName('idloteamento').Text);
 DM_Tabelas.ZQLoteamento.Filtered:=true;
end;

procedure TRLReportREMESSA.RLBand2AfterPrint(Sender: TObject);
begin
   DM_Tabelas.ZQLoteamento.Filtered:=false;
   DM_Tabelas.ZQLoteamento.close;
end;

procedure TRLReportREMESSA.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TRLReportREMESSA.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
