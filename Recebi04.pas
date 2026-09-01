unit Recebi04;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum;

type
  TFRM_Recebi04 = class(TForm)
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
    RLDBText2: TRLDBText;
    RLDBText3: TRLDBText;
    RLSubDetail1: TRLSubDetail;
    RLBand3: TRLBand;
    RLDBText4: TRLDBText;
    RLDBText5: TRLDBText;
    RLBand4: TRLBand;
    RLLabel9: TRLLabel;
    RLDBText6: TRLDBText;
    RLDBText7: TRLDBText;
    RLDBText8: TRLDBText;
    RLDBText9: TRLDBText;
    RLDBText10: TRLDBText;
    RLDBText11: TRLDBText;
    RLDBText13: TRLDBText;
    RLBand5: TRLBand;
    RLLabel10: TRLLabel;
    RLLabel11: TRLLabel;
    RLLabel12: TRLLabel;
    RLLabel13: TRLLabel;
    RLLabel14: TRLLabel;
    RLLabel15: TRLLabel;
    RLLabel16: TRLLabel;
    RLLabel17: TRLLabel;
    RLLabel19: TRLLabel;
    RLDBResult1: TRLDBResult;
    RLDBResult2: TRLDBResult;
    RLDBResult5: TRLDBResult;
    RLBand7: TRLBand;
    RLBand8: TRLBand;
    RLLabel20: TRLLabel;
    XNumEdit1: TXNumEdit;
    XNumEdit2: TXNumEdit;
    XNumEdit3: TXNumEdit;
    RLLabel18: TRLLabel;
    RLLabel21: TRLLabel;
    RLLabel22: TRLLabel;
    RLSubDetail2: TRLSubDetail;
    RLBand6: TRLBand;
    RLLabel23: TRLLabel;
    RLLabel24: TRLLabel;
    RLLabel25: TRLLabel;
    RLLabel26: TRLLabel;
    RLDBText12: TRLDBText;
    RLDBText14: TRLDBText;
    RLDBText15: TRLDBText;
    RLDBText16: TRLDBText;
    RLLabel27: TRLLabel;
    RLDBText17: TRLDBText;
    RLLabel28: TRLLabel;
    RLDBText18: TRLDBText;
    RLLabel29: TRLLabel;
    RLDBText19: TRLDBText;
    RLIreneg: TRLImage;
    RLImage1: TRLImage;
    RLLabel30: TRLLabel;
    RLLabel31: TRLLabel;
    RLLabel32: TRLLabel;
    XNumEdit4: TXNumEdit;
    XNumEdit5: TXNumEdit;
    RLLabel33: TRLLabel;
    RLIreneg2: TRLImage;
    RLLabel34: TRLLabel;
    procedure RLBand2BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand8BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLDBText8BeforePrint(Sender: TObject; var Text: String;
      var PrintIt: Boolean);
    procedure RLDBResult1BeforePrint(Sender: TObject; var Text: String;
      var PrintIt: Boolean);
    procedure RLReport1BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand3BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand4BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLBand6BeforePrint(Sender: TObject; var PrintIt: Boolean);
    procedure RLLabel3BeforePrint(Sender: TObject; var Text: String;
      var PrintIt: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FRM_Recebi04: TFRM_Recebi04;

implementation

uses RelRecebimento, funcoes, uRuntimeFields;
{$R *.dfm}

procedure TFRM_Recebi04.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFRM_Recebi04.RLBand8BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel18.Caption :='';
  RLLabel21.Caption :='';
  RLLabel22.Caption :='';
  if XNumEdit1.Value>0 then
    RLLabel18.Caption := Transform(strtofloat(XNumEdit1.Text),'###,###,##0.00');
  if XNumEdit2.Value>0 then
  RLLabel21.Caption := Transform(strtofloat(XNumEdit2.Text),'###,###,##0.00');
  if XNumEdit3.Value>0 then
  RLLabel22.Caption := Transform(strtofloat(XNumEdit3.Text),'###,###,##0.00');
end;

procedure TFRM_Recebi04.RLDBText8BeforePrint(Sender: TObject;
  var Text: String; var PrintIt: Boolean);
begin
  if FrmRelRecebimento.ZROQReceb2.FieldByName('saldo').Value>0 then
  begin
     RLDBText8.DataField:='Valor';
  //   XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1Valor.Value;
  end
  else
  begin
     RLDBText8.DataField:='Vr_rec';
//     XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1Vr_rec.Value;
  end;   
end;

procedure TFRM_Recebi04.RLDBResult1BeforePrint(Sender: TObject;
  var Text: String; var PrintIt: Boolean);
begin
//  if FrmRelRecebimento.ZROQReceb1saldo.Value>0 then
//     RLDBResult1.DataField:='Valor'
//  else
  //   RLDBResult1.DataField:='Vr_rec';

end;

procedure TFRM_Recebi04.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  XNumEdit1.Value:=0;
  XNumEdit2.Value:=0;
  XNumEdit4.Value:=0;
  XNumEdit5.Value:=0;
  RLLabel31.Caption :='';
  RLLabel32.Caption :='';

end;

procedure TFRM_Recebi04.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
  var
  tt:double;
begin
  if FrmRelRecebimento.ZROQReceb2.FieldByName('Substituicao').Value='S' then
  begin
    RLIreneg.Visible := True;
//    if FrmRelRecebimento.ZROQReceb1tipdoc.Value ='CH' then
//    begin
      RLLabel29.Visible:=false;
      RLDBText19.Visible:=false;
//      RLLabel34.Visible:=true;
//    end
//    else
//    begin
//      RLLabel29.Visible:=true;
//      RLDBText19.Visible:=true;
//    end;
  end
  else
  begin
    RLIreneg.Visible := false;
    RLLabel29.Visible:=true;
    RLDBText19.Visible:=true;
    RLLabel34.Visible:=false;    
  end;
  tt:=0;
  if FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value>FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value then
  begin
     RLDBText8.DataField:='Valor';
     XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value;
     RLDBText10.DataField:='Valor';
     tt:=FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Juros_Vr').Value;
     RLLabel33.Caption:=Transform(tt,'###,###,##0.00');
     XNumEdit2.Value:=XNumEdit2.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Juros_Vr').Value;
  end
  else
  begin
     RLDBText8.DataField:='vr_rec';
     XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value;
     RLDBText10.DataField:='vr_rec';
     tt:=FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Juros_Vr').Value;
     RLLabel33.Caption:=Transform(tt,'###,###,##0.00');
     XNumEdit2.Value:=XNumEdit2.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Juros_Vr').Value;
  end;

end;

procedure TFRM_Recebi04.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  FrmRelRecebimento.ZROQReceb2.First;
  while not FrmRelRecebimento.ZROQReceb2.eof do begin

    if FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value>FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value then
    begin
       XNumEdit4.Value:=XNumEdit4.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value;
       XNumEdit5.Value:=XNumEdit5.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Valor').Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Juros_Vr').Value;
    end
    else
    begin
       XNumEdit4.Value:=XNumEdit4.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value;
       XNumEdit5.Value:=XNumEdit5.Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Vr_rec').Value+FrmRelRecebimento.ZROQReceb2.FieldByName('Juros_Vr').Value;
    end;
    FrmRelRecebimento.ZROQReceb2.Next;
  end;
  if XNumEdit4.Value>0 then
     RLLabel31.Caption := Transform(strtofloat(XNumEdit4.Text),'###,###,##0.00');
  if XNumEdit5.Value>0 then
     RLLabel32.Caption := Transform(strtofloat(XNumEdit5.Text),'###,###,##0.00');
  XNumEdit4.Value:=0;
  XNumEdit5.Value:=0;

end;

procedure TFRM_Recebi04.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if FrmRelRecebimento.ZROQCh.FieldByName('Substituicao').Value='S' then
  begin
//      RLLabel34.Visible:=true;
//    RLIreneg2.Visible := True;
//    RLLabel29.Visible:=false;
//    RLDBText19.Visible:=false;
  end
  else
  begin
      RLLabel34.Visible:=false;  
//    RLIreneg2.Visible := false;
//    RLLabel29.Visible:=true;
//    RLDBText19.Visible:=true;
  end;

end;

procedure TFRM_Recebi04.RLLabel3BeforePrint(Sender: TObject;
  var Text: String; var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usu�rio:'+Frm_principal.xusuario.Caption;
end;


procedure TFRM_Recebi04.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
