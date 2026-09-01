unit RelBolBanco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLXLSFilter,
  RLRichFilter;

type
  TFrm_RelBolBanco = class(TForm)
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
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLDBText3: TRLDBText;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLLabel16: TRLLabel;
    RLDBText7: TRLDBText;
    RLLabel17: TRLLabel;
    RLDBText8: TRLDBText;
    RLLabel18: TRLLabel;
    RLDBText9: TRLDBText;
    RLBand3: TRLBand;
    RLSystemInfo5: TRLSystemInfo;
    RLMemo1: TRLMemo;
    marca0: TXNumEdit;
    Marca1: TXNumEdit;
    RLLmarca0: TRLLabel;
    RLLmarca1: TRLLabel;
    marcado1: TXNumEdit;
    RLLmarcado1: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    RLRichFilter1: TRLRichFilter;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_RelBolBanco: TFrm_RelBolBanco;

implementation

uses RelRecebimento, funcoes, Tabelas, principal, uRuntimeFields;

{$R *.dfm}

procedure TFrm_RelBolBanco.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  marca0.Value := marca0.Value + DM_tabelas.ZQRecebParticp.FieldByName('saldo').AsFloat;
  RLLabel11.Caption := DM_Tabelas.ZQRecebParticp.FieldByName('ende_cob').AsString+',     C.E.P.: '+DM_Tabelas.ZQRecebParticp.FieldByName('cep_cob').AsString;
  RLMemo1.Lines.Clear;
  if not empty(DM_Tabelas.ZQBol_men.FieldByName('linha1').AsString) Then
    RLMemo1.Lines.Add(DM_Tabelas.ZQBol_men.FieldByName('linha1').AsString);
  if not empty(DM_Tabelas.ZQBol_men.FieldByName('linha2').AsString) Then
    RLMemo1.Lines.Add(DM_Tabelas.ZQBol_men.FieldByName('linha2').AsString);
  if not empty(DM_Tabelas.ZQBol_men.FieldByName('linha3').AsString) Then
    RLMemo1.Lines.Add(DM_Tabelas.ZQBol_men.FieldByName('linha3').AsString);
  if not empty(DM_Tabelas.ZQBol_men.FieldByName('linha4').AsString) Then
    RLMemo1.Lines.Add(DM_Tabelas.ZQBol_men.FieldByName('linha4').AsString);
end;

procedure TFrm_RelBolBanco.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLmarca0.Caption := 'Valor enviado......: '+transforma(marca0.Value,12,2);
  RLLmarca1.Caption := 'Valor bloqueado: '+transforma(marca1.Value,12,2);
  RLLmarcado1.Caption := 'Bloqueados: '+marcado1.Text;
end;

procedure TFrm_RelBolBanco.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFrm_RelBolBanco.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

