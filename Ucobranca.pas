unit Ucobranca;

interface

uses ButtonDxArround, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Vcl.StdCtrls, Vcl.Mask, Vcl.ExtCtrls,
  FnpNumericEdit, Vcl.Grids, Vcl.DBGrids, Datasnap.DBClient, Vcl.ComCtrls, JvComponentBase, JvEnterTab;

type
  Tfrmcobranca = class(TForm)
    ZQRecebimento: TZQuery;
    DS_Recebimento: TDataSource;













































    Panel1: TPanel;
    Label1: TLabel;
    codcli: TMaskEdit;
    nomecli: TMaskEdit;
    Label6: TLabel;
    Label7: TLabel;
    perjur: TFnpNumericEdit;
    permul: TFnpNumericEdit;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    DBGrid1: TDBGrid;
    CDSCobranca: TClientDataSet;
    DS_cobranca: TDataSource;
    CDSCobrancaidrecebimento: TLargeintField;
    CDSCobrancadocumento: TWideStringField;
    CDSCobrancacliente: TIntegerField;
    CDSCobrancausuario: TIntegerField;
    CDSCobrancaDt_Entrada: TDateField;
    CDSCobrancaDt_Vencimento: TDateField;
    CDSCobrancaValor: TFloatField;
    CDSCobrancaObserv: TWideMemoField;
    CDSCobrancaVrDoc: TFloatField;
    CDSCobrancaordem: TWideStringField;
    CDSCobrancaTipDoc: TWideStringField;
    CDSCobrancasaldo: TFloatField;
    CDSCobrancacontabil: TIntegerField;
    CDSCobrancaorigem: TWideStringField;
    CDSCobrancaadversa: TIntegerField;
    CDSCobrancanumordem: TIntegerField;
    CDSCobrancaidloteamento: TIntegerField;
    CDSCobrancavenda_idvenda: TIntegerField;
    CDSCobrancaquadralote: TWideStringField;
    CDSCobrancanomeadversa: TWideStringField;
    CDSCobrancaPerc_multa: TFloatField;
    CDSCobrancaperc_juros: TFloatField;
    CDSCobrancaperc_descontos: TFloatField;
    CDSCobrancavr_multa: TFloatField;
    CDSCobrancavr_juros: TFloatField;
    CDSCobrancaVr_descontos: TFloatField;
    CDSCobrancasaldo_corrigido: TFloatField;
    Label2: TLabel;
    perdesc: TFnpNumericEdit;
    RGOrdem: TRadioGroup;
    Panel2: TPanel;
    database: TDateTimePicker;
    inicial: TDateTimePicker;
    final: TDateTimePicker;
    dxButton1: TdxButtonArround;
    dxButton2: TdxButtonArround;
    dxButton3: TdxButtonArround;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    sld_vencidos: TFnpNumericEdit;
    sld_venc_outros: TFnpNumericEdit;
    sld_venc_final: TFnpNumericEdit;
    sld_avencer: TFnpNumericEdit;
    sld_avenc_outros: TFnpNumericEdit;
    sld_avenc_final: TFnpNumericEdit;
    sld_geral: TFnpNumericEdit;
    sld_geral_outros: TFnpNumericEdit;
    sld_geral_final: TFnpNumericEdit;
    pb: TProgressBar;
    CDSCobrancadias_atraso: TFloatField;
    JvEnterAsTab1: TJvEnterAsTab;
    procedure dxButton3Click(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure databaseExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure dxButton2Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmcobranca: Tfrmcobranca;

implementation

uses tabelas,funcoes, RelUCobranca, uRuntimeFields;

{$R *.dfm}

procedure Tfrmcobranca.databaseExit(Sender: TObject);
begin
  final.Date:=Date;
end;

procedure Tfrmcobranca.DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  If CDSCobrancadias_atraso.Value > 0 then // condição
     Dbgrid1.Canvas.Font.Color:= clRed // coloque aqui a cor desejada
  else
     Dbgrid1.Canvas.Font.Color:= clGreen; // coloque aqui a cor desejada
  Dbgrid1.DefaultDrawDataCell(Rect, dbgrid1.columns[datacol].field, State);
inherited;
  ShowScrollBar(DBGrid1.Handle,SB_VERT,True); //força a mostragem barra Vertical
end;

procedure Tfrmcobranca.dxButton1Click(Sender: TObject);
begin
  sld_vencidos.Value:=0;
  sld_venc_outros.Value:=0;
  sld_venc_final.Value:=0;

  sld_avencer.Value:=0;
  sld_avenc_outros.Value:=0;
  sld_avenc_final.Value:=0;

  sld_geral.Value:=0;
  sld_geral_outros.Value:=0;
  sld_geral_final.Value:=0;

  pb.Visible:=true;
  pb.Position:=0;
  CDSCobranca.Close;
  CDSCobranca.CreateDataSet;
  CDSCobranca.Open;
  CDSCobranca.EmptyDataSet;
  ZQRecebimento.Close;
  ZQRecebimento.SQL.Clear;
  if RGOrdem.ItemIndex=0 then
  begin
     ZQRecebimento.SQL.Add('Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
     ZQRecebimento.SQL.Add('       custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
     ZQRecebimento.SQL.Add('       somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
     ZQRecebimento.SQL.Add('       dt_nao_pagou_no_mes,descricao_juridico ');
     ZQRecebimento.SQL.Add('      from Recebimento where adversa='+quotedstr(codcli.Text)+' and  saldo>0 and DT_Entrada between :dt1 and :dt2 order by DT_Entrada');
  end
  else
  begin
     ZQRecebimento.SQL.Add('Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
     ZQRecebimento.SQL.Add('       custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
     ZQRecebimento.SQL.Add('       somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
     ZQRecebimento.SQL.Add('       dt_nao_pagou_no_mes,descricao_juridico ');
     ZQRecebimento.SQL.Add(' from Recebimento where adversa='+quotedstr(codcli.Text)+'  and  saldo>0 and DT_Vencimento between :dt1 and :dt2 order by DT_Vencimento');
  end;
  ZQRecebimento.ParamByName('dt1').AsDate:=inicial.Date;
  ZQRecebimento.ParamByName('dt2').AsDate:=final.Date;
  ZQRecebimento.Open;
  ZQRecebimento.First;
  pb.Max:=ZQRecebimento.RecordCount;
  ZQRecebimento.DisableControls;
  while not ZQRecebimento.Eof do
  begin
    pb.Position:=ZQRecebimento.RecNo;
    CDSCobranca.Insert;
    CDSCobrancaidrecebimento.Value:=ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
    CDSCobrancadocumento.Value:=ZQRecebimento.FieldByName('documento').AsString;
    CDSCobrancacliente.Value:=ZQRecebimento.FieldByName('cliente').AsInteger;
    CDSCobrancausuario.Value:=ZQRecebimento.FieldByName('usuario').AsInteger;
    CDSCobrancaDt_Entrada.Value:=ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime;
    CDSCobrancaDt_Vencimento.Value:=ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
    CDSCobrancaValor.Value:= ZQRecebimento.FieldByName('Valor').AsFloat;
    CDSCobrancaObserv.Value:= ZQRecebimento.FieldByName('Observ').AsString;
    CDSCobrancaVrDoc.Value:= ZQRecebimento.FieldByName('VrDoc').AsFloat;
    CDSCobrancaordem.Value:=ZQRecebimento.FieldByName('ordem').AsString;
    CDSCobrancaTipDoc.Value:=ZQRecebimento.FieldByName('TipDoc').AsString;
    CDSCobrancasaldo.Value:=ZQRecebimento.FieldByName('saldo').AsFloat;
    CDSCobrancacontabil.Value:=ZQRecebimento.FieldByName('contabil').AsInteger;
    CDSCobrancaorigem.Value:=ZQRecebimento.FieldByName('origem').AsString;
    CDSCobrancaadversa.Value:=ZQRecebimento.FieldByName('adversa').AsInteger;
    CDSCobrancanumordem.Value:=ZQRecebimento.FieldByName('numordem').AsInteger;
    CDSCobrancaidloteamento.Value:=ZQRecebimento.FieldByName('idloteamento').AsInteger;
    CDSCobrancavenda_idvenda.Value:=ZQRecebimento.FieldByName('venda_idvenda').AsInteger;
    CDSCobrancaquadralote.Value:=ZQRecebimento.FieldByName('quadralote').AsString;
    CDSCobrancanomeadversa.Value:=ZQRecebimento.FieldByName('nomeadversa').AsString;
    CDSCobrancaPerc_multa.Value:=permul.Value/30;
    CDSCobrancaperc_juros.Value:=perjur.Value;
    CDSCobrancaperc_descontos.Value:=perdesc.Value;

    if ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime<=database.Date then
    begin
      CDSCobrancadias_atraso.Value:= strtofloat(Calc_Dias(ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime,database.Date));
    end
    else
      CDSCobrancadias_atraso.Value:=0;

    CDSCobrancavr_multa.Value:=(ZQRecebimento.FieldByName('saldo').AsFloat*(CDSCobrancadias_atraso.Value*CDSCobrancaPerc_multa.Value))/100;
    if CDSCobrancadias_atraso.Value>0 then
       CDSCobrancavr_juros.Value:=(ZQRecebimento.FieldByName('saldo').AsFloat*perjur.Value)/100
     else
       CDSCobrancavr_juros.Value:=0;
    if perdesc.Value>0 then
      CDSCobrancaVr_descontos.Value:=(ZQRecebimento.FieldByName('saldo').AsFloat*perdesc.Value)/100
    else
      CDSCobrancaVr_descontos.Value:=0;
    CDSCobrancasaldo_corrigido.Value:=(ZQRecebimento.FieldByName('saldo').AsFloat+CDSCobrancavr_multa.Value+CDSCobrancavr_juros.Value)-CDSCobrancaVr_descontos.Value;
    CDSCobranca.Post;
    if CDSCobrancadias_atraso.Value>0 then
    begin
      sld_vencidos.Value:=sld_vencidos.Value+ZQRecebimento.FieldByName('saldo').AsFloat;
      sld_venc_outros.Value:=(sld_venc_outros.Value+CDSCobrancavr_multa.Value+CDSCobrancavr_juros.Value)-CDSCobrancaVr_descontos.Value;
      sld_venc_final.Value:=sld_venc_final.Value+CDSCobrancasaldo_corrigido.Value;

    end
    else
    begin
      sld_avencer.Value:=sld_avencer.Value+ZQRecebimento.FieldByName('saldo').AsFloat;
      sld_avenc_outros.Value:=(sld_avenc_outros.Value+CDSCobrancavr_multa.Value+CDSCobrancavr_juros.Value)-CDSCobrancaVr_descontos.Value;
      sld_avenc_final.Value:=sld_avenc_final.Value+CDSCobrancasaldo_corrigido.Value;
    end;

    ZQRecebimento.Next;
  end;
  ZQRecebimento.EnableControls;
  sld_geral.Value:=sld_vencidos.Value+sld_avencer.Value;
  sld_geral_outros.Value:=sld_venc_outros.Value+sld_avenc_outros.Value;
  sld_geral_final.Value:=sld_venc_final.Value+sld_avenc_final.Value;

  pb.Position:=0;
  pb.Visible:=false;

end;

procedure Tfrmcobranca.dxButton2Click(Sender: TObject);
begin
  Frm_RelUCobranca:=nil;
  if Frm_RelUCobranca=nil then
     Frm_RelUCobranca:= TFrm_RelUCobranca.Create(Application);
  Frm_RelUCobranca.RLLabel2.Caption:='Data Base:'+datetostr(database.Date)+'   -  Período de :'+datetostr(inicial.Date)+' até '+datetostr(final.Date);
  Frm_RelUCobranca.RLReport1.PreviewModal;
  Frm_RelUCobranca:=nil;
end;

procedure Tfrmcobranca.dxButton3Click(Sender: TObject);
begin
  close;
end;

procedure Tfrmcobranca.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ZQRecebimento.Close;
  CDSCobranca.Close;
  frmcobranca:=nil;
  Action:=CaFree;
end;

procedure Tfrmcobranca.FormShow(Sender: TObject);
begin
  database.Date:=date;
  inicial.Date:=strtodate('01/01/2000');
  perjur.SetFocus;
end;


procedure Tfrmcobranca.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(Tfrmcobranca, 'ZQRecebimento', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentojuridico', 'juridico', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentodata_juridico', 'data_juridico', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentodt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(Tfrmcobranca, 'ZQRecebimento', 'ZQRecebimentodescricao_juridico', 'descricao_juridico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
