unit Resumo_quita;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, StdCtrls,QRPDFFilt, QRXLSXFilt;

type
  TFrmResumo_quita = class(TForm)
    QRResumo_quita: TQuickRep;
    QRBand1: TQRBand;
    QRSysData4: TQRSysData;
    QRSysData5: TQRSysData;
    QRSysData6: TQRSysData;
    QRDBImage1: TQRDBImage;
    QRLabel15: TQRLabel;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRDBText1: TQRDBText;
    QRLabel3: TQRLabel;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRSubDetail1: TQRSubDetail;
    QRBand3: TQRBand;
    QRDBText4: TQRDBText;
    QRBand4: TQRBand;
    QRLabel8: TQRLabel;
    QRLabel10: TQRLabel;
    QRSubDetail2: TQRSubDetail;
    QRDBText7: TQRDBText;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    QRSubDetail3: TQRSubDetail;
    QRLabel13: TQRLabel;
    QRLabel16: TQRLabel;
    QRDBText11: TQRDBText;
    QRDBText13: TQRDBText;
    QRDBText15: TQRDBText;
    QRBand5: TQRBand;
    QRBand6: TQRBand;
    QRSubDetail4: TQRSubDetail;
    QRBand7: TQRBand;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRSubDetail5: TQRSubDetail;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRDBText16: TQRDBText;
    QRLabel30: TQRLabel;
    QRDBText17: TQRDBText;
    QRLabel31: TQRLabel;
    QRSubDetail6: TQRSubDetail;
    QRDBText18: TQRDBText;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRShape2: TQRShape;
    QRLabel40: TQRLabel;
    QRLabel6: TQRLabel;
    QRBand8: TQRBand;
    QRLabel49: TQRLabel;
    QRLabel50: TQRLabel;
    QRLabel51: TQRLabel;
    QRLabel52: TQRLabel;
    QRBand9: TQRBand;
    QRLabel2: TQRLabel;
    QRShape4: TQRShape;
    QRLabel7: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel11: TQRLabel;
    QRDBText5: TQRDBText;
    QRLabel12: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel41: TQRLabel;
    QRShape1: TQRShape;
    QRLabel42: TQRLabel;
    QRLabel43: TQRLabel;
    QRPDFFilter1: TQRPDFFilter;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    QRLabel46: TQRLabel;
    QRDBText6: TQRDBText;
    QRDBText8: TQRDBText;
    QRDBText12: TQRDBText;
    QRXMLSSFilter1: TQRXMLSSFilter;
    procedure QRResumo_quitaBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand4BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRResumo_quitaAfterPrint(Sender: TObject);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand7BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    { Private declarations }

    procedure ConfigurarCamposImovel;
    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmResumo_quita: TFrmResumo_quita;

implementation

uses Tabelas, funcoes,Venda, principal, DB;

{$R *.dfm}

type
  { Recalcula os lookups da venda no buffer atual sem alterar o estado do
    dataset. O relatorio e preparado depois que a venda e filtrada. }
  TDataSetCalcFieldsAccess = class(TDataSet);

procedure RecalcularDadosQuitacao;
begin
  if (DM_Tabelas = nil) or
     (not DM_Tabelas.ZQVenda.Active) or
     DM_Tabelas.ZQVenda.IsEmpty then
    Exit;

  TDataSetCalcFieldsAccess(DM_Tabelas.ZQVenda).GetCalcFields(
    DM_Tabelas.ZQVenda.ActiveBuffer);
end;

procedure TFrmResumo_quita.ConfigurarCamposImovel;
var
  LField: TField;
begin
  if (DM_Tabelas = nil) or
     (not DM_Tabelas.ZQvnd_quadro.Active) or
     DM_Tabelas.ZQvnd_quadro.IsEmpty then
    Exit;

  { A banda e preparada somente depois que a consulta mestre-detalhe foi
    posicionada. Assim os tres TQRDBText leem o registro da venda atual. }
  LField := DM_Tabelas.ZQvnd_quadro.FindField('apelido');
  if (LField <> nil) and (Trim(LField.AsString) <> '') then
  begin
    QRDBText1.DataSet := DM_Tabelas.ZQvnd_quadro;
    QRDBText1.DataField := 'apelido';
  end
  else if DM_Tabelas.ZQvnd_quadro.FindField('nomeloteamento') <> nil then
  begin
    QRDBText1.DataSet := DM_Tabelas.ZQvnd_quadro;
    QRDBText1.DataField := 'nomeloteamento';
  end;

  if DM_Tabelas.ZQvnd_quadro.FindField('quadra') <> nil then
  begin
    QRDBText2.DataSet := DM_Tabelas.ZQvnd_quadro;
    QRDBText2.DataField := 'quadra';
  end;
  if DM_Tabelas.ZQvnd_quadro.FindField('lote') <> nil then
  begin
    QRDBText3.DataSet := DM_Tabelas.ZQvnd_quadro;
    QRDBText3.DataField := 'lote';
  end;
end;

procedure TFrmResumo_quita.QRResumo_quitaBeforePrint(
  Sender: TCustomQuickRep; var PrintReport: Boolean);
begin
  RecalcularDadosQuitacao;
  DM_Tabelas.ZQEmpresa.Open;
  DM_Tabelas.ZQEmpresa.First;
  if DM_Tabelas.ZQvnd_quadro.Active then
    DM_Tabelas.ZQvnd_quadro.Close;
  DM_Tabelas.ZQvnd_quadro.Open;
  ConfigurarCamposImovel;

  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Open;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Open;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  DM_Tabelas.ZQCompr_Dados.First;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger,[]);
  DM_Tabelas.ZQCidade.Open;
  RecalcularDadosQuitacao;

  DM_TAbelas.CDSAss1.Close;
  DM_TAbelas.CDSAss1.CreateDataSet;


  DM_Tabelas.zqprocuradores.Filtered:=false;
  DM_Tabelas.zqprocuradores.Filter:='Assina=''S''';
  DM_Tabelas.zqprocuradores.filtered:=true;



  if (DM_Tabelas.CDSIncorpassina.Value='S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString<>'S') then
  begin
    DM_TAbelas.CDSAss1.Insert;
    DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
    if alltrim(DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString)<>'ALICIO VILAR PONTES' then
       DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.CDSIncorpnomeparti.Value;
    DM_Tabelas.CDSAss1conjuge.Value := DM_Tabelas.CDSIncorpconjuge.Value;
    DM_Tabelas.CDSAss1idconjuge.Value := DM_Tabelas.CDSIncorpidconjuge.Value;
    DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.CDSIncorpassina.Value;
    DM_TAbelas.CDSAss1.Post;
  end;

  if (DM_Tabelas.CDSIncorpassina.Value='S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString='S') then
  begin
    DM_Tabelas.zqprocuradores.First;
    DM_Tabelas.zqprocuradores.DisableControls;
    while not DM_Tabelas.zqprocuradores.Eof do
    begin
      DM_TAbelas.CDSAss1.Insert;
      DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
      DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
      DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;

      DM_Tabelas.CDSAss1conjuge.Value := DM_Tabelas.CDSIncorpconjuge.Value;
      DM_Tabelas.CDSAss1idconjuge.Value := DM_Tabelas.CDSIncorpidconjuge.Value;
      DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.CDSIncorpassina.Value;


      DM_TAbelas.CDSAss1.Post;
      DM_Tabelas.zqprocuradores.Next;
    end;
    DM_Tabelas.zqprocuradores.EnableControls;
  end;

  if (DM_Tabelas.CDSIncorpassina.Value<>'S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString='S') then
  begin
    DM_Tabelas.zqprocuradores.First;
    DM_Tabelas.zqprocuradores.DisableControls;
    while not DM_Tabelas.zqprocuradores.Eof do
    begin
      DM_TAbelas.CDSAss1.Insert;
      DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
      DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
      DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;
  //    DM_Tabelas.CDSAss1assina.Value := DM_Tabelas.zqprocuradoresassina.Value;
      DM_TAbelas.CDSAss1.Post;
      DM_Tabelas.zqprocuradores.Next;
    end;
    DM_Tabelas.zqprocuradores.EnableControls;
  end;


  if (DM_Tabelas.CDSIncorpassina.Value<>'S') and (DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString<>'S') then
  begin
    DM_TAbelas.CDSAss1.Insert;
    DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString;
    DM_Tabelas.CDSAss1cargo.Value :=' Vendedor(a)';
    if alltrim(DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').AsString)<>'ALICIO VILAR PONTES' then
       DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.CDSIncorpnomeparti.Value;
    DM_TAbelas.CDSAss1.Post;
  end;


  DM_Tabelas.zqprocuradores.Filtered:=false;


//  DM_TAbelas.CDSAss1.Post;
  DM_TAbelas.CDSAss2.Close;
  DM_TAbelas.CDSAss2.CreateDataSet;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').AsString) then Begin
    DM_TAbelas.CDSAss2.Insert;
    DM_TAbelas.CDSAss2assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').AsString;
    DM_Tabelas.CDSAss2cargo.Value :='Testemunha';
    DM_TAbelas.CDSAss2.Post;
  end;
  DM_TAbelas.CDSAss.Close;
  DM_TAbelas.CDSAss.CreateDataSet;
  DM_Tabelas.ZQCompr_Dados.DisableControls;
  while not DM_Tabelas.ZQCompr_Dados.Eof do begin
    DM_TAbelas.CDSAss.Insert;
    DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Dados.FieldByName('nome_parte').AsString;
    DM_Tabelas.CDSAsscargo.Value :='Comprador(a)';
    DM_TAbelas.CDSAss.Post;
    DM_TAbelas.ZQCompr_Resp_Dados.DisableControls;
    while not DM_TAbelas.ZQCompr_Resp_Dados.Eof do begin
      DM_TAbelas.CDSAss.Insert;
      DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Resp_Dados.FieldByName('nomerepres').AsString;
      DM_Tabelas.CDSAsscargo.Value :='Responsável';
      DM_TAbelas.CDSAss.Post;
      DM_TAbelas.ZQCompr_Resp_Dados.Next;
    end;
    DM_TAbelas.ZQCompr_Resp_Dados.EnableControls;
    DM_TAbelas.ZQCompr_Conjuge.DisableControls;
    while not DM_TAbelas.ZQCompr_Conjuge.Eof do
    begin
      if DM_Tabelas.ZQCompr_Dados.FieldByName('tipopessoa').AsString='F' then
      begin
        if DM_Tabelas.ZQCompr_conjuge.FieldByName('assina').AsString = 'S' Then
        Begin
          DM_TAbelas.CDSAss.Insert;
          DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_conjuge.FieldByName('nome').AsString;
          DM_Tabelas.CDSAsscargo.Value :='Conjuge';
          DM_TAbelas.CDSAss.Post;
        end;
      end;
      DM_TAbelas.ZQCompr_Conjuge.Next;
    end;
    DM_TAbelas.ZQCompr_Conjuge.EnableControls;
    DM_Tabelas.ZQCompr_Dados.Next;
  end;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  DM_Tabelas.ZQCompr_Dados.EnableControls;
end;

procedure TFrmResumo_quita.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel8.Caption:=transform(Frm_Venda.XDentrada.value,'###,###,###,##0.00');
  QRLabel10.Caption:=transform(Frm_Venda.XDqtdeentrada.Value,'##0');

end;

procedure TFrmResumo_quita.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  ConfigurarCamposImovel;
end;

procedure TFrmResumo_quita.QRBand4BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel13.Caption:=transform((Frm_Venda.sldpago.value-Frm_Venda.XDentrada.value),'###,###,###,##0.00');
  QRLabel16.Caption:=transform(Frm_Venda.xdparpagas.Value,'##0');
  QRLabel18.Caption:=transform(Frm_Venda.sldpago.value,'###,###,###,##0.00');
  QRLabel20.Caption:=transform(Frm_Venda.ultparc.value,'###,###,###,##0.00');
  QRLabel21.Caption:='Vincendas: '+transform(Frm_Venda.xdsldvicenda.value-Frm_Venda.xdsldatrasadas.Value,'##0')+'  Atrasadas: '+transform(Frm_Venda.xdsldatrasadas.Value,'##0')+'  Total: '+transform(Frm_Venda.xdsldvicenda.value,'##0');
  QRLabel22.Caption:=transform(Frm_Venda.xdsldremanescente.value,'###,###,###,##0.00');
  QRLabel25.Caption:=transform(Frm_Venda.xdliquidado.value,'###,###,###,##0.00');
  QRLabel27.Caption:=transform(Frm_Venda.vrfinal.value,'###,###,###,##0.00');
  QRLabel28.Caption := DM_Tabelas.ZQvnd_quadro.FieldByName('nomecid').AsString+' / '+DM_Tabelas.ZQvnd_quadro.FieldByName('estado').AsString+', '+dataporextenso(date);
  QRLabel33.Caption := transform(Frm_Venda.Ultimo_Igpm.value,'###,###,###,##0.00');
  QRLabel35.Caption := transform(Frm_Venda.xpcorrecao_igpm.value,'##0.00')+'%  -  Vlr.:'+transform(Frm_Venda.xvcorrecao_igpm.value,'###,###,###,##0.00');
  QRLabel41.Caption := transform(Frm_Venda.vr_price.value,'###,###,###,##0.00');
  QRLabel37.Caption := transform(Frm_Venda.xvracrescimo.value,'###,###,###,##0.00');
  QRLabel39.Caption := transform(Frm_Venda.xvrdesconto.value,'###,###,###,##0.00');
  QRLabel43.Caption:=transform(Frm_Venda.Xcorrigido.Value,'###,###,###,##0.00');  
end;

procedure TFrmResumo_quita.QRResumo_quitaAfterPrint(Sender: TObject);
begin
  DM_Tabelas.ZQEmpresa.close;
end;

procedure TFrmResumo_quita.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel40.Caption:='Usuário:'+Frm_principal.xusuario.Caption;
end;

procedure TFrmResumo_quita.QRBand7BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
 QRLabel12.Caption:='Tot. Parc. Financ.:'+frm_venda.xdqtdeparcelas.Text;
end;


procedure TFrmResumo_quita.AfterConstruction;
begin
  inherited AfterConstruction;
  QRBand2.BeforePrint := QRBand2BeforePrint;
end;

end.
