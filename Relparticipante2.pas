unit Relparticipante2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls, RLPDFFilter,
  RLRichFilter, RLFilters, RLXLSFilter, RLParser, XNum, RLReport;

type
  TFrmRelparticipante2 = class(TForm)
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
    RLBand2: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    XNPrincipal: TXNumEdit;
    RLExpressionParser1: TRLExpressionParser;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLPDFFilter1: TRLPDFFilter;
    RLLabel9: TRLLabel;
    RLLabel12: TRLLabel;
    RLSubDetail1: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText9: TRLDBText;
    RLDBText2: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLLabel11: TRLLabel;
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelparticipante2: TFrmRelparticipante2;
  varsomabaixa : Double;
  varvar : string;
implementation

uses Tabelas, funcoes, Participante, principal,hormain, uRuntimeFields;

{$R *.dfm}


procedure TFrmRelparticipante2.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLBand3.Color  := clWhite;
  if frmMain.CDS_etq.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;
end;

procedure TFrmRelparticipante2.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if frmMain=nil then
     frmMain := TfrmMain.Create(self);
  FrmRelparticipante2.RLReport1.DataSource:=frmMain.DS_etq;
end;

procedure TFrmRelparticipante2.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrmRelparticipante2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
