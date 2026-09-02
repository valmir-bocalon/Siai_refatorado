unit Relajuste2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, RLParser, RLRichFilter, RLXLSFilter, RLFilters,
  RLPDFFilter, StdCtrls, XNum, Mask, XEdit;

type
  TFrmRelReajuste2 = class(TForm)
    RLBand1: TRLBand;
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
    RLExpressionParser1: TRLExpressionParser;
    RLPDFFilter1: TRLPDFFilter;
    RLRichFilter1: TRLRichFilter;
    RLReport1: TRLReport;
    XNsomabaixa: TXNumEdit;
    XNNegociado: TXNumEdit;
    XNRenova: TXNumEdit;
    XNPrincipal: TXNumEdit;
    RLGroup1: TRLGroup;
    RLBand2: TRLBand;
    RLBand3: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel9: TRLLabel;
    RLLabel18: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel30: TRLLabel;
    RLBand4: TRLBand;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText17: TRLDBText;
    RLDBText16: TRLDBText;
    RLDBText18: TRLDBText;
    RLDBText20: TRLDBText;
    RLBand5: TRLBand;
    RLDBResult1: TRLDBResult;
    RLDBResult2: TRLDBResult;
    RLBand6: TRLBand;
    RLDBResult3: TRLDBResult;
    RLDBResult4: TRLDBResult;
    RLLabel16: TRLLabel;
    RLXLSFilter1: TRLXLSFilter;
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLGroup1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelReajuste2: TFrmRelReajuste2;
  varvar : string;
implementation

uses Tabelas, Funcoes, ReajusteDeParcelas2, principal,RelRecebimento, uRuntimeFields;
{$R *.dfm}

procedure TFrmRelReajuste2.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmRelReajuste2.RLGroup1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
    if FrmRelRecebimento.RGFormato.ItemIndex=6 then
    begin
      case FrmRelRecebimento.RGOrdem.ItemIndex of
        0:FrmRelReajuste2.RLLabel16.Caption:='Vencimento: '+FrmRelRecebimento.ZQRecebimento.FieldByName('Dt_Vencimento').Text;
        1:FrmRelReajuste2.RLLabel16.Caption:='Comprador: '+FrmRelRecebimento.ZQRecebimento.FieldByName('nomeadversa').Text;
        2:FrmRelReajuste2.RLLabel16.Caption:='Quadra-Lote: '+FrmRelRecebimento.ZQRecebimento.FieldByName('quadralote').Text;
        3:FrmRelReajuste2.RLLabel16.Caption:='Tipo Documento: '+FrmRelRecebimento.ZQRecebimento.FieldByName('TipDoc').Text;
        4:FrmRelReajuste2.RLLabel16.Caption:='Loteamento: '+FrmRelRecebimento.ZQRecebimento.FieldByName('idloteamento').Text;
      end;
    end;
end;


procedure TFrmRelReajuste2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

