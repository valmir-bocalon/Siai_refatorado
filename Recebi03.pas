unit Recebi03;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RLReport, StdCtrls, XNum, RLFilters, RLPDFFilter, RLRichFilter,
  RLXLSFilter;

type
  TFRM_Recebi03 = class(TForm)
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
    RLLabel35: TRLLabel;
    RLLabel36: TRLLabel;
    RLLabel37: TRLLabel;
    RLLabel38: TRLLabel;
    RLLabel39: TRLLabel;
    RLLabel40: TRLLabel;
    RLLabel41: TRLLabel;
    RLLabel42: TRLLabel;
    RLLabel43: TRLLabel;
    RLLabel44: TRLLabel;
    RLLabel45: TRLLabel;
    RLDraw1: TRLDraw;
    RLLabel46: TRLLabel;
    RLLabel47: TRLLabel;
    RLLabel48: TRLLabel;
    RLLabel49: TRLLabel;
    RLDraw2: TRLDraw;
    RLLabel50: TRLLabel;
    RLLabel51: TRLLabel;
    RLLabel52: TRLLabel;
    RLLabel53: TRLLabel;
    RLBand9: TRLBand;
    RLDBMemo1: TRLDBMemo;
    RLDBMemo2: TRLDBMemo;
    RLLabel54: TRLLabel;
    RLDBText20: TRLDBText;
    XNumEdit6: TXNumEdit;
    XNumEdit7: TXNumEdit;
    RLLabel55: TRLLabel;
    RLPDFFilter1: TRLPDFFilter;
    XNumEdit8: TXNumEdit;
    XNumEdit9: TXNumEdit;
    RLLabel56: TRLLabel;
    RLLabel57: TRLLabel;
    RLXLSFilter1: TRLXLSFilter;
    RLRichFilter1: TRLRichFilter;
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
    procedure RLLabel31AfterPrint(Sender: TObject);
    procedure RLLabel32AfterPrint(Sender: TObject);
    procedure RLLabel55AfterPrint(Sender: TObject);
    procedure RLBand1BeforePrint(Sender: TObject; var PrintIt: Boolean);
  private
    FParcelasTotalGeral: TStringList;
    function RegistrarParcelaUnica(ALista: TStringList;
      const AIdRecib: Int64): Boolean;

    procedure AfterConstruction; override;
  public
    { Public declarations }
    destructor Destroy; override;
  end;

var
  FRM_Recebi03: TFRM_Recebi03;

implementation

uses RelRecebimento, funcoes, principal, uRuntimeFields;
{$R *.dfm}

destructor TFRM_Recebi03.Destroy;
begin
  FreeAndNil(FParcelasTotalGeral);
  inherited;
end;

function TFRM_Recebi03.RegistrarParcelaUnica(ALista: TStringList;
  const AIdRecib: Int64): Boolean;
var
  Chave: string;
begin
  Chave := IntToStr(AIdRecib);
  Result := ALista.IndexOf(Chave) < 0;
  if Result then
    ALista.Add(Chave);
end;

procedure TFRM_Recebi03.RLBand2BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLReport1.NewPage;
end;

procedure TFRM_Recebi03.RLBand8BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel18.Caption :='';
  RLLabel21.Caption :='';
  RLLabel22.Caption :='';
  RLLabel56.Caption :='Total de Descontos: 0.00';
  RLLabel57.Caption :='Total de Juros: 0.00';
  if XNumEdit1.Value>0 then
    RLLabel18.Caption := Transform(strtofloat(XNumEdit1.Text),'###,###,##0.00');
  if XNumEdit2.Value>0 then
     RLLabel21.Caption := Transform(strtofloat(XNumEdit2.Text),'###,###,##0.00');
//  XNumEdit3.Value:=XNumEdit2.Value-XNumEdit1.Value;
  if XNumEdit3.Value<>0 then
  RLLabel22.Caption := Transform(strtofloat(XNumEdit3.Text),'###,###,##0.00');
  XNumEdit3.Value:=0;

  if XNumEdit8.Value>0 then
    RLLabel56.Caption := 'Total de Descontos: '+Transform(strtofloat(XNumEdit8.Text),'###,###,##0.00');
  if XNumEdit9.Value>0 then
     RLLabel57.Caption := 'Total de Juros: '+Transform(strtofloat(XNumEdit9.Text),'###,###,##0.00');
  XNumEdit8.Value:=0;
  XNumEdit9.Value:=0;
end;

procedure TFRM_Recebi03.RLDBText8BeforePrint(Sender: TObject;
  var Text: String; var PrintIt: Boolean);
begin
{  if FrmRelRecebimento.ZROQReceb1saldo.Value>0 then
  begin
     RLDBText8.DataField:='valor_1';
  //   XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1valor_1.Value;
  end
  else
  begin
     RLDBText8.DataField:='Vr_rec';
//     XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1Vr_rec.Value;
  end;}
end;

procedure TFRM_Recebi03.RLDBResult1BeforePrint(Sender: TObject;
  var Text: String; var PrintIt: Boolean);
begin
//  if FrmRelRecebimento.ZROQReceb1saldo.Value>0 then
//     RLDBResult1.DataField:='valor_1'
//  else
  //   RLDBResult1.DataField:='Vr_rec';

end;

procedure TFRM_Recebi03.RLReport1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if FParcelasTotalGeral = nil then
  begin
    FParcelasTotalGeral := TStringList.Create;
    FParcelasTotalGeral.Sorted := True;
    FParcelasTotalGeral.Duplicates := dupIgnore;
  end
  else
    FParcelasTotalGeral.Clear;

  XNumEdit1.Value:=0;
  XNumEdit2.Value:=0;
  XNumEdit3.Value:=0;
  XNumEdit4.Value:=0;
  XNumEdit5.Value:=0;
  XNumEdit6.Value:=0;
  XNumEdit7.Value:=0;
  XNumEdit8.Value:=0;
  XNumEdit9.Value:=0;
  RLLabel31.Caption :='';
  RLLabel32.Caption :='';
  RLLabel55.caption :='';
  RLLabel18.Caption :='';
  RLLabel21.Caption :='';
  RLLabel22.caption :='';

end;

procedure TFRM_Recebi03.RLBand3BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
  tt,vb:double;
  ParcelaUnica: Boolean;
begin

  if FrmRelRecebimento.ZROQReceb1.FieldByName('juros').AsFloat>0 then
  begin
     RLDBText11.Font.Color:=clRed;
     RLDBText11.Visible:=true;   // juros
     RLDBText20.Visible:=false; // descontos
  end
  else if FrmRelRecebimento.ZROQReceb1.FieldByName('descon').AsFloat>0 then
  begin
     RLDBText20.Font.Color:=clBlue;
     RLDBText11.Visible:=false;// juros
     RLDBText20.Visible:=true; // descontos
  end
  else
  begin
     RLDBText20.Font.Color:=clBlack;
     RLDBText11.Font.Color:=clBlack;
     RLDBText11.Visible:=false;// juros
     RLDBText20.Visible:=true; // desconto
     RLDBText20.Caption:='0.00';
  end;
  if FrmRelRecebimento.ZROQReceb1.FieldByName('substituicao').AsString='S' then
  begin
    RLIreneg.Visible := True;
    RLDBText19.Visible:=false;
//    RLLabel34.Visible:=true;
  end
  else
  begin
    RLIreneg.Visible := false;
//    RLDBText19.Visible:=true;
    RLLabel34.Visible:=false;
  end;
  if FrmRelRecebimento.ZROQReceb1.FieldByName('Vr_rec').AsFloat>0 then
  begin
    RLDBText18.DataField:='dataref';
//    RLDBText18.DataField:='Vencimento';
  end
  else
  begin
   RLDBText18.DataField:='';
  end;

  tt:=0;
  vb:=0;
  ParcelaUnica := RegistrarParcelaUnica(FParcelasTotalGeral,
    FrmRelRecebimento.ZROQReceb1.FieldByName('idrecib').AsInteger);
{  if FrmRelRecebimento.ZROQReceb1vr_rec.Value>FrmRelRecebimento.ZROQReceb1valor_1.Value then
  begin
     RLDBText8.DataField:='valor_1';
     XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1valor_1.Value;

     RLLabel35.Caption:=transform(FrmRelRecebimento.ZROQReceb1valor_1.Value,'###,###,##0.00');

     RLDBText10.DataField:='valor_1';
     tt:=FrmRelRecebimento.ZROQReceb1valor_1.Value+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
     vb:=FrmRelRecebimento.ZROQReceb1valor_1.Value;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
     RLLabel33.Caption:=Transform(tt,'###,###,##0.00');
     RLLabel54.caption:=Transform(vb,'###,###,##0.00');
     XNumEdit2.Value:=XNumEdit2.Value+FrmRelRecebimento.ZROQReceb1valor_1.Value+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
  end
  else
  begin}
//     if FrmRelRecebimento.ZROQReceb1vr_rec.Value>0 then
//     begin
       if ParcelaUnica then
       begin
       RLDBText8.DataField:='valor_1';//'valor_parcela_1';
       XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1.FieldByName('Valor_1').AsFloat; //FrmRelRecebimento.ZROQReceb1Valor_parcela_1.Value;
       RLLabel35.Caption:=transform(FrmRelRecebimento.ZROQReceb1.FieldByName('Valor_1').AsFloat,'###,###,##0.00'); //transform(FrmRelRecebimento.ZROQReceb1valor_parcela_1.Value,'###,###,##0.00');
       end
       else
       begin
         RLDBText8.DataField:='';
         RLLabel35.Caption:='';
       end;
  {   end
     else
     begin
       RLDBText8.DataField:='vr_ab';
       XNumEdit1.Value:=XNumEdit1.Value+FrmRelRecebimento.ZROQReceb1Vr_ab.Value;
       RLLabel35.Caption:=transform(FrmRelRecebimento.ZROQReceb1vr_ab.Value,'###,###,##0.00');
     end;
     RLDBText10.DataField:='vr_rec';
     if (FrmRelRecebimento.ZROQReceb1vr_rec.Value=tt) and (FrmRelRecebimento.ZROQReceb1Juros_Vr.Value>0) then
     begin                          }
       tt:=FrmRelRecebimento.ZROQReceb1.FieldByName('valor').AsFloat;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
       if ParcelaUnica then
         vb:=FrmRelRecebimento.ZROQReceb1.FieldByName('Valor_1').AsFloat
       else
         vb:=0;
       XNumEdit2.Value:=XNumEdit2.Value+FrmRelRecebimento.ZROQReceb1.FieldByName('valor').AsFloat;//-FrmRelRecebimento.ZROQReceb1Juros_Vr.Value+FrmRelRecebimento.ZROQReceb1Desc_Vr.Value;
{     end
     else if (FrmRelRecebimento.ZROQReceb1vr_rec.Value<>tt) and (FrmRelRecebimento.ZROQReceb1Juros_Vr.Value>=0) then
     begin
       tt:=FrmRelRecebimento.ZROQReceb1vr_rec.Value;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
       vb:=FrmRelRecebimento.ZROQReceb1vr_rec.Value+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
       XNumEdit2.Value:=XNumEdit2.Value+FrmRelRecebimento.ZROQReceb1Vr_rec.Value+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value+FrmRelRecebimento.ZROQReceb1Desc_Vr.Value;
     end
     else if (FrmRelRecebimento.ZROQReceb1vr_rec.Value=tt) and (FrmRelRecebimento.ZROQReceb1Juros_Vr.Value=0) then
     begin
       tt:=FrmRelRecebimento.ZROQReceb1vr_rec.Value;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
       vb:=FrmRelRecebimento.ZROQReceb1vr_rec.Value;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
       XNumEdit2.Value:=XNumEdit2.Value+FrmRelRecebimento.ZROQReceb1Vr_rec.Value;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
     end;}

     RLLabel33.Caption:=Transform(tt,'###,###,##0.00');
     if ParcelaUnica then
       RLLabel54.caption:=Transform(vb,'###,###,##0.00')
     else
       RLLabel54.caption:='';

//     if FrmRelRecebimento.ZROQReceb1Juros_Vr.Value>0 then
//        showmessage('aqui');

//  end;

  if FrmRelRecebimento.ZROQReceb1.FieldByName('Vr_rec').AsFloat>FrmRelRecebimento.ZROQReceb1.FieldByName('Valor_1').AsFloat then
  begin
     RLDBText17.DataField:='vr_rec';
  end
  else
  begin
     if FrmRelRecebimento.ZROQReceb1.FieldByName('Vr_rec').AsFloat>0 then
     begin
       RLDBText17.DataField:='vr_rec';
     end
     else
     begin
       RLDBText17.DataField:='valor';
     end;
  end;
{  if not empty(FrmRelRecebimento.ZROQReceb1obsebx.Value) then
     RLBand9.Visible:=true
  else
     RLBand9.Visible:=false;}
end;

procedure TFRM_Recebi03.RLBand4BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
var
  ParcelasSubtotal: TStringList;
begin
  ParcelasSubtotal := TStringList.Create;
  try
    ParcelasSubtotal.Sorted := True;
    ParcelasSubtotal.Duplicates := dupIgnore;

    FrmRelRecebimento.ZROQReceb1.First;
    while not FrmRelRecebimento.ZROQReceb1.eof do begin

{    if FrmRelRecebimento.ZROQReceb1vr_rec.Value>FrmRelRecebimento.ZROQReceb1valor_1.Value then
    begin
       XNumEdit4.Value:=XNumEdit4.Value+FrmRelRecebimento.ZROQReceb1valor_1.Value;
       XNumEdit5.Value:=XNumEdit5.Value+FrmRelRecebimento.ZROQReceb1valor_1.Value+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
    end
    else
    begin
       if FrmRelRecebimento.ZROQReceb1Vr_rec.Value>0 then
       begin}
         if RegistrarParcelaUnica(ParcelasSubtotal,
           FrmRelRecebimento.ZROQReceb1.FieldByName('idrecib').AsInteger) then
         XNumEdit4.Value:=XNumEdit4.Value+FrmRelRecebimento.ZROQReceb1.FieldByName('Valor_1').AsFloat;  //FrmRelRecebimento.ZROQReceb1Valor_parcela_1.Value;
         XNumEdit5.Value:=XNumEdit5.Value+FrmRelRecebimento.ZROQReceb1.FieldByName('valor').AsFloat;//+FrmRelRecebimento.ZROQReceb1Juros_Vr.Value;
         XNumEdit6.Value:=XNumEdit6.Value+(FrmRelRecebimento.ZROQReceb1.FieldByName('juros').AsFloat-FrmRelRecebimento.ZROQReceb1.FieldByName('descon').AsFloat);
         XNumEdit3.Value:=XNumEdit3.Value+(FrmRelRecebimento.ZROQReceb1.FieldByName('juros').AsFloat-FrmRelRecebimento.ZROQReceb1.FieldByName('descon').AsFloat);
         XNumEdit8.Value:=XNumEdit8.Value+FrmRelRecebimento.ZROQReceb1.FieldByName('descon').AsFloat;
         XNumEdit9.Value:=XNumEdit9.Value+FrmRelRecebimento.ZROQReceb1.FieldByName('juros').AsFloat;
{       end
       else
       begin
         XNumEdit4.Value:=XNumEdit4.Value+FrmRelRecebimento.ZROQReceb1Vr_ab.Value;
       end;
    end;}
      FrmRelRecebimento.ZROQReceb1.Next;
    end;
  finally
    ParcelasSubtotal.Free;
  end;

  if XNumEdit4.Value>0 then
     RLLabel31.Caption := Transform(strtofloat(XNumEdit4.Text),'###,###,##0.00');
  if XNumEdit5.Value>0 then
     RLLabel32.Caption := Transform(strtofloat(XNumEdit5.Text),'###,###,##0.00');
  if XNumEdit6.Value<>0 then
     RLLabel55.Caption := Transform(strtofloat(XNumEdit6.Text),'###,###,##0.00');
  if XNumEdit5.Value>XNumEdit4.Value then
     RLLabel55.Font.Color:=clRed
  else if XNumEdit5.Value<XNumEdit4.Value then
     RLLabel55.Font.Color:=clBlue
  else
    RLLabel55.Font.Color:=clBlack;
  XNumEdit4.Value:=0;
  XNumEdit5.Value:=0;
  XNumEdit6.Value:=0;

end;

procedure TFRM_Recebi03.RLBand6BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  if FrmRelRecebimento.ZROQCh.FieldByName('Substituicao').AsString='S' then
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
  if FrmRelRecebimento.ZROQReceb1.FieldByName('Vr_rec').AsFloat>FrmRelRecebimento.ZROQReceb1.FieldByName('Valor_1').AsFloat then
  begin
     RLDBText17.DataField:='vr_rec';
  end
  else
  begin
     if FrmRelRecebimento.ZROQReceb1.FieldByName('Vr_rec').AsFloat>0 then
     begin
       RLDBText17.DataField:='vr_rec';
     end
     else
     begin
       RLDBText17.DataField:='valor';
     end;
  end;

end;

procedure TFRM_Recebi03.RLLabel31AfterPrint(Sender: TObject);
begin
  RLLabel31.Caption :='';
end;

procedure TFRM_Recebi03.RLLabel32AfterPrint(Sender: TObject);
begin
  RLLabel32.Caption :='';
end;

procedure TFRM_Recebi03.RLLabel55AfterPrint(Sender: TObject);
begin
    RLLabel55.Caption :='';
end;

procedure TFRM_Recebi03.RLBand1BeforePrint(Sender: TObject;
  var PrintIt: Boolean);
begin
  RLLabel3.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;


procedure TFRM_Recebi03.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.

