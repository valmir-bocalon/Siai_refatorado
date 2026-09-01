unit RelImovel01;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLXLSFilter,
  RLRichFilter;

type
  TFrm_RelImovel01 = class(TForm)
    RLReport1: TRLReport;
    RLBand1: TRLBand;
    RLLabel1: TRLLabel;
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
    RLSubDetail1: TRLSubDetail;
    RLDBImage2: TRLDBImage;
    RLDBText4: TRLDBText;
    RLBand3: TRLBand;
    RLDBText3: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText9: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel10: TRLLabel;
    RLDBText8: TRLDBText;
    RLBand5: TRLBand;
    XConta: TXNumEdit;
    RLDBText10: TRLDBText;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel19: TRLLabel;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLBand6: TRLBand;
    RLLabel26: TRLLabel;
    RLLabel27: TRLLabel;
    RLLabel28: TRLLabel;
    RLBand7: TRLBand;
    RLDBText20: TRLDBText;
    RLDBText21: TRLDBText;
    RLDBText22: TRLDBText;
    RLDraw1: TRLDraw;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLDBText14: TRLDBText;
    RLLabel15: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLRichFilter1: TRLRichFilter;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand5BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelImovel01: TFrm_RelImovel01;
  nm:integer;
implementation

uses tabelas, RelImovel, funcoes, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrm_RelImovel01.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFrm_RelImovel01.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
//  XConta.Value := XConta.Value+1;
  if nm<> Frm_RelImovel.ZQImov.RecordCount then
  begin
    XConta.Value := XConta.Value+Frm_RelImovel.ZQImov.RecordCount;
    nm:=Frm_RelImovel.ZQImov.RecordCount;
  end;  
  RLBand3.Color  := clWhite;
  if Frm_RelImovel.ZQImov.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;

end;

procedure TFrm_RelImovel01.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel15.Caption := '';
  if XConta.Value>0 Then
  RLLabel15.Caption := Transform(strtofloat(XConta.Text),'###,###,##0')+'  Im�veis';

end;

procedure TFrm_RelImovel01.RLReport1AfterPrint(Sender: TObject);
begin
  nm:=0;
  xconta.Value:=0;
end;

procedure TFrm_RelImovel01.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelImovel01.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
