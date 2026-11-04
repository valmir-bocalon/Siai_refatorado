unit RelImovelREMESSA2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLXLSFilter,
  RLRichFilter, DB;

type
  TRLReportREMESSA2 = class(TForm)
    RLReportREMESSA2: TRLReport;
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
    RLDBText2: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
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
    RLDBText9: TRLDBText;
    RLLabel18: TRLLabel;
    RLDBText10: TRLDBText;
    RLLabel19: TRLLabel;
    RLDBText11: TRLDBText;
    RLLabel11: TRLLabel;
    RLDBText6: TRLDBText;
    RLLabel20: TRLLabel;
    rtiti: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLRichFilter1: TRLRichFilter;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReportREMESSA2BeforePrint(Sender: TObject;
      var PrintIt: Boolean);
    procedure RLBand2AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
    procedure ConfigurarFontes(ADetalhes, ACabecalho: TDataSource);
  end;

var
  RLReportREMESSA2: TRLReportREMESSA2;

implementation

uses tabelas, RelImovel, funcoes, RelREMESSA, principal, uRuntimeFields;
{$R *.dfm}

procedure TRLReportREMESSA2.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReportREMESSA2.NewPage;
end;

procedure TRLReportREMESSA2.RLReportREMESSA2BeforePrint(Sender: TObject; var PrintIt: Boolean);
begin
  if (RLReportREMESSA2.DataSource <> nil) and (RLReportREMESSA2.DataSource.DataSet <> nil) then
    rtiti.Caption := IntToStr(RLReportREMESSA2.DataSource.DataSet.RecordCount)
  else
    rtiti.Caption := '0';
end;

procedure TRLReportREMESSA2.RLBand2AfterPrint(Sender: TObject);
begin
  { As consultas deste modelo pertencem apenas a esta visualizacao. }
end;

procedure TRLReportREMESSA2.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if Frm_principal <> nil then
    RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TRLReportREMESSA2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

procedure TRLReportREMESSA2.ConfigurarFontes(ADetalhes, ACabecalho: TDataSource);
begin
  RLReportREMESSA2.DataSource := ADetalhes;
  RLDBImage1.DataSource := ACabecalho;
  RLDBText1.DataSource := ACabecalho;
  RLDBText2.DataSource := ADetalhes;
  RLDBText3.DataSource := ADetalhes;
  RLDBText4.DataSource := ADetalhes;
  RLDBText5.DataSource := ADetalhes;
  RLDBText6.DataSource := ADetalhes;
  RLDBText7.DataSource := ADetalhes;
  RLDBText8.DataSource := ADetalhes;
  RLDBText9.DataSource := ADetalhes;
  RLDBText10.DataSource := ADetalhes;
  RLDBText11.DataSource := ADetalhes;
  RLDBResult1.DataSource := ADetalhes;
end;

end.
