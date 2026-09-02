unit Recebitotpago;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter;

type
  TFRM_Recebitotpago = class(TForm)
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
    XNumEdit1: TXNumEdit;
    XNumEdit2: TXNumEdit;
    XNumEdit3: TXNumEdit;
    XNumEdit4: TXNumEdit;
    XNumEdit5: TXNumEdit;
    XNumEdit6: TXNumEdit;
    XNumEdit7: TXNumEdit;
    RLPDFFilter1: TRLPDFFilter;
    RLBand2: TRLBand;
    RLLabel9: TRLLabel;
    RLLabel32: TRLLabel;
    RLBand4: TRLBand;
    RLDBMemo1: TRLDBMemo;
    RLBand3: TRLBand;
    RLDBMemo2: TRLDBMemo;
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLDBText4: TRLDBText;
    RLBand5: TRLBand;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLDBText5: TRLDBText;
    RLDraw1: TRLDraw;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand8BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLLabel32AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FRM_Recebitotpago: TFRM_Recebitotpago;

implementation

uses RelRecebimento, funcoes, principal, uRuntimeFields;
{$R *.dfm}

procedure TFRM_Recebitotpago.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFRM_Recebitotpago.RLBand8BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  XNumEdit3.Value:=0;
end;

procedure TFRM_Recebitotpago.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  XNumEdit1.Value:=0;
  XNumEdit2.Value:=0;
  XNumEdit3.Value:=0;
  XNumEdit4.Value:=0;
  XNumEdit5.Value:=0;
  XNumEdit6.Value:=0;
  XNumEdit7.Value:=0;  

  RLLabel32.Caption :='';


end;

procedure TFRM_Recebitotpago.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  FrmRelRecebimento.ZRTotPago.First;
  while not FrmRelRecebimento.ZRTotPago.eof do begin

         XNumEdit5.Value:=XNumEdit5.Value+FrmRelRecebimento.ZRTotPago.FieldByName('somapago').AsFloat;
    FrmRelRecebimento.ZRTotPago.Next;
  end;

  if XNumEdit5.Value>0 then
     RLLabel32.Caption := Transform(strtofloat(XNumEdit5.Text),'###,###,##0.00');


  XNumEdit4.Value:=0;
  XNumEdit5.Value:=0;
  XNumEdit6.Value:=0;

end;

procedure TFRM_Recebitotpago.RLLabel32AfterPrint(Sender: TObject);
begin
  RLLabel32.Caption :='';
end;

procedure TFRM_Recebitotpago.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFRM_Recebitotpago.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
