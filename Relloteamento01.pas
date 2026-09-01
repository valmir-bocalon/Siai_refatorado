unit Relloteamento01;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLXLSFilter,
  RLRichFilter;

type
  TFrm_Relloteamento01 = class(TForm)
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
    RLSubDetail1: TRLSubDetail;
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
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLDBText10: TRLDBText;
    RLLabel17: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel19: TRLLabel;
    RLLabel20: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    RLDBText11: TRLDBText;
    RLDBText12: TRLDBText;
    RLDBText13: TRLDBText;
    RLDBText14: TRLDBText;
    RLDBText15: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBImage1: TRLDBImage;
    RLDBText1: TRLDBText;
    RLDBText17: TRLDBText;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand5BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLSubDetail1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Relloteamento01: TFrm_Relloteamento01;
  nm:integer;

implementation

uses tabelas, Relloteamento, funcoes, principal, uRuntimeFields;
{$R *.dfm}

procedure TFrm_Relloteamento01.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
//  RLReport1.NewPage;
end;

procedure TFrm_Relloteamento01.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
//  XConta.Value := XConta.Value+1;
  if nm<> Frm_Relloteamento.ZQImov.RecordCount then
  begin
    XConta.Value := XConta.Value+Frm_Relloteamento.ZQImov.RecordCount;
    nm:=Frm_Relloteamento.ZQImov.RecordCount;
  end;
  RLBand3.Color  := clWhite;
  if Frm_Relloteamento.ZQImov.recno mod 2 = 0 then
    RLBand3.Color  := $00EAEAEA;
end;

procedure TFrm_Relloteamento01.RLBand5BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel15.Caption := '';
  if XConta.Value>0 Then
  RLLabel15.Caption := Transform(strtofloat(XConta.Text),'###,###,##0')+'  Im�veis';
end;

procedure TFrm_Relloteamento01.RLReport1AfterPrint(Sender: TObject);
begin
  nm:=0;
  xconta.Value:=0;  
end;

procedure TFrm_Relloteamento01.RLSubDetail1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  Frm_Relloteamento.ZQImov.close;
  Frm_Relloteamento.ZQImov.SQL.Clear;
  Frm_Relloteamento.ZQImov.SQL.Add('select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento, '+
                                   ' cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto '+
                                   '  from imovel where loteamento_idloteamento='+quotedstr(Frm_Relloteamento.CDSEmpreendcodempre.Text));
  Frm_Relloteamento.ZQImov.Open;
  Frm_Relloteamento.ZQImov.Filtered := False;
  if Frm_Relloteamento.FCBS.Checked<> Frm_Relloteamento.FCBN.Checked Then
  begin
    if Frm_Relloteamento.FCBS.Checked Then
      Frm_Relloteamento.ZQImov.Filter := 'disponivel='+quotedstr('SIM')
    else
      Frm_Relloteamento.ZQImov.Filter := 'disponivel='+quotedstr('N�O');
    Frm_Relloteamento.ZQImov.Filtered := True;
  end;
  Frm_Relloteamento.ZQImov.RecordCount;
end;

procedure TFrm_Relloteamento01.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_Relloteamento01.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.


