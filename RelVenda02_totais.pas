unit RelVenda02_totais;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, FnpNumericEdit, RLRichFilter, RLXLSFilter,
  RLFilters, RLPDFFilter, RLParser;

type
  TFrm_RelVenda02_totais = class(TForm)
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
    RLSubDetail1: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLBand10: TRLBand;
    RLLabel58: TRLLabel;
    RLLabel62: TRLLabel;
    RLDBText3: TRLDBText;
    RLDBText9: TRLDBText;
    RLLabel59: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel18: TRLLabel;
    tabela: TFnpNumericEdit;
    geral: TFnpNumericEdit;
    RLDraw1: TRLDraw;
    RLDBText10: TRLDBText;
    RLPDFFilter1: TRLPDFFilter;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
    RLLabel19: TRLLabel;
    tlinha: TFnpNumericEdit;
    RLDBImage2: TRLDBImage;
    RLDBText2: TRLDBText;
    RLDraw7: TRLDraw;
    RLDraw8: TRLDraw;
    RLLabel15: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel9: TRLLabel;
    RLLabel10: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel20: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    rpar: TFnpNumericEdit;
    rentra: TFnpNumericEdit;
    rvista: TFnpNumericEdit;
    RLSystemInfo5: TRLSystemInfo;
    RLDBResult1: TRLDBResult;
    RLDBResult2: TRLDBResult;
    RLLabel23: TRLLabel;
    inicial: TRLLabel;
    RLLabel24: TRLLabel;
    Final: TRLLabel;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand10BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelVenda02_totais: TFrm_RelVenda02_totais;
implementation

uses RelVenda,funcoes, tabelas, principal, uRuntimeFields;
{$R *.dfm}
procedure TFrm_RelVenda02_totais.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFrm_RelVenda02_totais.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  tabela.Value:=0;
  geral.Value:=0;
  tlinha.Value:=0;
  rvista.Value:=0;
  rentra.Value:=0;
  rpar.Value:=0;
end;

procedure TFrm_RelVenda02_totais.RLBand10BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel59.Caption:=transform(geral.Value,'###,###,##0.00');
  RLLabel62.Caption:=transform(tabela.Value,'###,###,##0.00');
  RLLabel22.Caption:=transform(rvista.Value,'###,###,##0.00');
  RLLabel21.Caption:=transform(rentra.Value,'###,###,##0.00');
  RLLabel20.Caption:=transform(rpar.Value,'###,###,##0.00');
  tabela.Value:=0;
  geral.Value:=0;
  tlinha.Value:=0;
  rvista.Value:=0;
  rentra.Value:=0;
  rpar.Value:=0;
end;

procedure TFrm_RelVenda02_totais.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if (Frm_RelVenda.cvendavr_parcela.Value>0) and (Frm_RelVenda.cvendavr_entrada.Value>0) then
  begin
    RLDBText10.Visible:=false;
    RLLabel18.Visible:=true;
    RLDBText3.Visible:=true;
    rentra.Value:=rentra.Value+Frm_RelVenda.cvendavr_entrada.Value;
//    geral.Value:=geral.Value+Frm_RelVenda.cvendavr_entrada.Value;
    RLLabel17.Visible:=false;

  end;
  if (Frm_RelVenda.cvendavr_parcela.Value=0) and (Frm_RelVenda.cvendavr_entrada.Value>0) then
  begin
    RLDBText10.Visible:=true;
    rvista.Value:=rvista.Value+Frm_RelVenda.cvendavr_entrada.Value;
//    geral.Value:=geral.Value+Frm_RelVenda.cvendavr_entrada.Value;
    RLLabel18.Visible:=false;
    RLDBText3.Visible:=false;
    RLLabel17.Visible:=true;
  end;
  tabela.Value:=tabela.Value+Frm_RelVenda.cvendavalorvenal.Value;
//  tabela.Value:=tabela.Value+Frm_RelVenda.cvendavalorvenda.Value;
//  tlinha.Value:=tlinha.Value+(Frm_RelVenda.cvendavalorvenda.Value+((Frm_RelVenda.cvendajuros.value+Frm_RelVenda.cvendajuros_p.value)-(Frm_RelVenda.cvendadescontos.value-Frm_RelVenda.cvendadescontos_p.value)));
  tlinha.Value:=tlinha.Value+(Frm_RelVenda.cvendavr_parcela.Value+Frm_RelVenda.cvendavr_entrada.Value);
  geral.Value:=geral.Value+tlinha.Value;
  RLLabel19.Caption:=transform(tlinha.Value,'###,###,##0.00');
  rpar.Value:=rpar.Value+Frm_RelVenda.cvendavr_parcela.Value;
//  geral.Value:=geral.Value+Frm_RelVenda.cvendavr_parcela.Value;
//  geral.Value:=geral.Value+Frm_RelVenda.cvendavalorvenda.Value;
end;

procedure TFrm_RelVenda02_totais.RLBand3AfterPrint(Sender: TObject);
begin
  tlinha.Value:=0;
end;

procedure TFrm_RelVenda02_totais.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelVenda02_totais.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
