

unit RecebBaixa_subst;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, Mask, DBCtrls,
  XBanner, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, ImgList, XNum, DBClient, XDBDate, XDBNum, ComCtrls,
  JvExComCtrls, JvDateTimePicker, XEdit, XDBEdit, ExtCtrls, XDate,
  TFlatGaugeUnit, System.ImageList;
type
  TFrmRecebBaixa_subst = class(TForm)
    Label10: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    DBEBaixaDocum: TDBEdit;
    DBEBaixaVrRec: TDBEdit;
    DXBBaixaGravar: TdxButtonArround;
    DS_RecBai: TDataSource;
    ZQRecBai: TZQuery;
    DBGBaixando: TDBGrid;
    DXBFechar: TdxButtonArround;
    ImageList1: TImageList;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    DS_Parcelas: TDataSource;
    DBGrid2: TDBGrid;
    Label5: TLabel;
    DBCBTIPDOC: TDBComboBox;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label15: TLabel;
    Label9: TLabel;
    Label17: TLabel;
    Label19: TLabel;
    EMObsbx: TMemo;
    Label20: TLabel;
    EMNovostit: TMemo;
    Label21: TLabel;
    Label22: TLabel;
    EContabil: TEdit;
    Label11: TLabel;
    GBCheque: TGroupBox;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    DBEAgencia: TDBEdit;
    DBEBanco: TDBEdit;
    DBEDono: TDBEdit;
    DBECpf_Cnpj: TDBEdit;
    XDBNumEdit1: TXDBNumEdit;
    DBText1: TDBText;
    Label28: TLabel;
    DEBConta: TDBEdit;
    GBContra: TGroupBox;
    Enomes: TEdit;
    Ecodnome: TEdit;
    RBParticipante: TRadioButton;
    RBCheque: TRadioButton;
    Label1: TLabel;
    Label23: TLabel;
    GBBanco: TGroupBox;
    DBGBanco: TDBGrid;
    Label29: TLabel;
    EPlanocontabanco: TEdit;
    RBPonteiro: TRadioButton;
    DBEidreg: TDBEdit;
    LREg: TLabel;
    CBSaldoCorrigido: TCheckBox;
    DS_RecebBxTemp: TDataSource;
    ZQRecebBxTemp: TZQuery;

    XNJuros: TXNumEdit;
    Label30: TLabel;
    Label31: TLabel;
    Label33: TLabel;
    DS_dividdo: TDataSource;
    CDS_dividido: TClientDataSet;
    DS_ratear: TDataSource;
    ZQRatear: TZQuery;
    CDS_divididoparticipante: TIntegerField;
    CDS_divididovr_base: TFloatField;
    CDS_divididovr_calc: TFloatField;
    CDS_divididocontabil: TIntegerField;
    CDS_divididotipodoc: TWideStringField;
    CDS_divididoobs: TWideMemoField;
    DS_ItensRecibo: TDataSource;
    ZQItensRecibo: TZQuery;
    CDS_divididonomeparti: TWideStringField;
    CDS_divididodescrcontabil: TWideStringField;
    CDS_divididodescrcontabil2: TWideStringField;
    Memo1: TMemo;
    DS_Div_usado: TDataSource;
    ZQDiv_usado: TZQuery;




































































































    XBanner9: TXBanner;
    XBanner1: TXBanner;
    XBanner2: TXBanner;
    XBanner3: TXBanner;












    GBAdversa: TGroupBox;
    XBanner5: TXBanner;
    EcodAdversa: TEdit;
    Eadversa: TEdit;
    GroupBox1: TGroupBox;
    XBanner4: TXBanner;
    EcodAdversaNova: TEdit;
    EAdversaNova: TEdit;
    XEEmpresend: TXEdit;
    CBTipobaixa: TComboBox;








    Label18: TLabel;
    DXBPesquisar: TdxButtonArround;





    DS_Receb_baixa: TDataSource;
    ZQReceb_Baixa: TZQuery;















    Label32: TLabel;

    ZQrecebimento: TZQuery;
    DataRecebimento: TDataSource;





























    Label35: TLabel;
    Elote: TEdit;
    btndes: TButton;
    Panel1: TPanel;
    JDEntrada: TXDateEdit;
    XNEAberto: TXNumEdit;
    XNEVencido: TXNumEdit;
    XNEMarcado: TXNumEdit;
    XNEJuros: TXNumEdit;
    XNERecebido: TXNumEdit;
    XNEDesconto: TXNumEdit;
    XNEResto: TXNumEdit;
    JDBaixa: TXDateEdit;
    XNEParcelas: TXNumEdit;
    CBPeriodoEnt: TComboBox;
    WDTBaixaEntra: TXDBDateEdit;
    DTPIni: TXDateEdit;
    DTPFim: TXDateEdit;
    Label37: TLabel;
    pini: TMaskEdit;
    Label38: TLabel;
    pfim: TMaskEdit;
    fim: TLabel;
    dpl: TLabel;
    Label39: TLabel;
    EDLoteamento: TComboBox;
    DataZQSubstituicao: TDataSource;
    ZQSubstituicao: TZQuery;
    DataZQbaixa: TDataSource;
    ZQbaixa: TZQuery;
    DataZQBaixa2: TDataSource;
    ZQBaixa2: TZQuery;



























    ProgressBar1: TProgressBar;
    barravertical: TFlatGauge;






    ZQRecebe: TZQuery;
    DS_Recebimento: TDataSource;

    Label34: TLabel;
    DBEdit1: TDBEdit;
    CDSParcelas: TClientDataSet;
    CDSParcelasTipDoc: TStringField;
    CDSParcelasDocum: TStringField;
    CDSParcelasOrdem: TStringField;
    CDSParcelasVenci: TDateField;
    CDSParcelasVrParc: TFloatField;
    CDSParcelascliente: TIntegerField;
    CDSParcelasusuario: TIntegerField;
    CDSParcelasentrada: TDateField;
    CDSParcelasbanco: TStringField;
    CDSParcelasagencia: TStringField;
    CDSParcelasdono: TStringField;
    CDSParcelascpf_cnpj: TStringField;
    CDSParcelascodcontabil: TIntegerField;
    CDSParcelasdesccontabil: TStringField;
    CDSParcelaspercent: TFloatField;
    CDSParcelasconta: TStringField;
    CDSParcelascodbanco: TIntegerField;
    CDSParcelascontabanco: TIntegerField;
    CDSParcelassomapag: TStringField;
    CDSParcelasjuros: TFloatField;
    CDSParcelasdesconto: TFloatField;
    CDSParcelasnomedoempreendimento: TStringField;
    CDSParcelasvenda_idvenda: TIntegerField;
    CDSParcelasquadralote: TStringField;
    CDSParcelasSubstituicao: TStringField;
    CDSParcelassq: TIntegerField;
    CDSParcelasorigem: TStringField;
    CDSParcelasParcelas_fixas: TStringField;
    CDSParcelasncheque: TStringField;

    procedure Atualiza_tela;
    procedure DXBFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGBaixandoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGBaixandoColEnter(Sender: TObject);
    procedure DBGBaixandoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure XNERecebidoExit(Sender: TObject);
    procedure DXBBaixaGravarClick(Sender: TObject);
    procedure DBCBTIPDOCExit(Sender: TObject);
    procedure DXBBaixaGravarEnter(Sender: TObject);
    procedure DBCBTIPDOCEnter(Sender: TObject);
    procedure DBEBaixaVrRecExit(Sender: TObject);
    procedure EContabilExit(Sender: TObject);
    procedure XDBNumEdit1Exit(Sender: TObject);
    procedure DBEBaixaDocumExit(Sender: TObject);
    procedure EContabilEnter(Sender: TObject);
    procedure DBECpf_CnpjExit(Sender: TObject);
    procedure EnomesExit(Sender: TObject);
    procedure Label1Click(Sender: TObject);
    procedure RBChequeClick(Sender: TObject);
    procedure GBBancoExit(Sender: TObject);
    procedure GBChequeExit(Sender: TObject);
    procedure EPlanocontabancoExit(Sender: TObject);
    procedure filtrabaixa;
    procedure GBContraExit(Sender: TObject);
    procedure RBParticipanteClick(Sender: TObject);
    procedure RBPonteiroClick(Sender: TObject);
    procedure XDBNumEdit1Enter(Sender: TObject);
    procedure JDEntradaExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ZQRecBaiCalcFields(DataSet: TDataSet);
    procedure CDSParcelasCalcFields(DataSet: TDataSet);
    procedure ZQRatearCalcFields(DataSet: TDataSet);
    procedure XNEJurosEnter(Sender: TObject);
    procedure CDS_divididoCalcFields(DataSet: TDataSet);
    procedure CBPeriodoEntExit(Sender: TObject);
    procedure EadversaExit(Sender: TObject);
    procedure EAdversaNovaExit(Sender: TObject);
    procedure CBTipobaixaExit(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure btndesClick(Sender: TObject);
    procedure EDLoteamentoDropDown(Sender: TObject);
    procedure DBEDonoExit(Sender: TObject);
    procedure DBGBaixandoExit(Sender: TObject);
    procedure Label9MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRecebBaixa_subst: TFrmRecebBaixa_subst;
  VArDoc, varrecpag : String;

implementation

uses tabelas, Cad_Recebimento, funcoes, ImpRecibo, PesqRecebimento_bx, uRuntimeFields;

{$R *.dfm}

// procedure de mover o formulario sem bordas
procedure MovimentaObject(Sender:TObject;Button:TMouseButton;Shift:TShiftState;X,Y:Integer;Formulario:TForm); 
 var ObjectPos,MousePosMov:TPoint;Pt:TPoint;fHandle:HWND; 
 begin 
   GetCursorPos(Pt); 
   ObjectPos.X:=Formulario.Left; 
   ObjectPos.Y:=Formulario.Top; 
   if(Sender is TForm)then 
     fHandle:=TWinControl(Sender).Handle 
   else 
     fHandle:=TWinControl(Sender).Parent.Handle; 
   while DragDetect(fHandle,ObjectPos) do 
   begin 
     GetCursorPos(MousePosMov); 
     Formulario.Left:=MousePosMov.X-X-3; 
     Formulario.Top:=MousePosMov.Y-Y-3; 
     Application.ProcessMessages;
   end; 
 end; 

procedure TFrmRecebBaixa_subst.Atualiza_tela;
Var
  Varreg, varmarc, varvenc : integer;
Begin
  if ZQRecBai.active=false then
     ZQRecBai.open;
  Varreg := ZQRecBai.RecNo;
  ZQRecBai.First;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  varmarc := 0;
  varvenc := 0;

  ProgressBar1.Max:=ZQRecBai.RecordCount;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=true;
  DS_RecBai.DataSet:=nil;


  ZQRecBai.first;
  ZQRecBai.DisableControls;
  while not ZQRecBai.Eof do
  begin
    application.ProcessMessages;
    ProgressBar1.Position:=ZQRecBai.RecNo;
    application.ProcessMessages;
    XNEAberto.Value := XNEAberto.Value + ZQRecBai.FieldByName('saldocorrig').AsFloat;
    if ZQRecBai.FieldByName('marcar').AsString = '0' Then Begin
      XNEMarcado.Value := XNEMarcado.Value + ZQRecBai.FieldByName('saldocorrig').AsFloat;
      XNEJuros.Value := XNEJuros.Value + (ZQRecBai.FieldByName('saldocorrig').AsFloat-ZQRecBai.FieldByName('saldo').AsFloat);
      inc(varmarc);
    end;
    if ZQRecBai.FieldByName('Dt_Vencimento').AsDateTime <=date Then Begin
      XNEVencido.Value := XNEVencido.Value + ZQRecBai.FieldByName('saldocorrig').AsFloat;
      inc(varvenc);
    end;
    ZQRecBai.Next;
  end;
  ZQRecBai.EnableControls;
  DS_RecBai.DataSet:=ZQRecBai;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;


  application.ProcessMessages;
  Label31.Caption := inttostr(varmarc);
  Label33.Caption := inttostr(varvenc);
  XNERecebido.Value := XNEMarcado.Value;
  ZQRecBai.RecNo := Varreg;
End;

procedure TFrmRecebBaixa_subst.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRecebBaixa_subst.FormShow(Sender: TObject);
begin
  dm_tabelas.ZQLoteamento.Open;
  DM_tabelas.ZQCobaRe.open;
  DM_Tabelas.ZQNumOrdem.open;

//  DM_Tabelas.ZQDividido.open;
{  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.Open;}
//  if  DM_tabelas.ZQRecebimento.Active=false then
//       DM_tabelas.ZQRecebimento.Open;
//
//
//  if ZQRecebBxTemp.Active=false then
//  begin
//    ZQRecebBxTemp.Close;
//    ZQRecebBxTemp.SQL.Clear;
//    ZQRecebBxTemp.SQL.Add('SELECT * FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib join receb_baixa as B ON B.refbaixa=H.refer');
//    ZQRecebBxTemp.open;
//  end;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and 1=0');
  DM_tabelas.ZQRecebimento.open;

  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReBxHi.SQL.Clear;
  DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist limit 0, 1');
  DM_tabelas.ZQReBxHi.open;


  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa limit 0, 1');
  DM_tabelas.ZQReceb_Baixa.open;


  DM_Tabelas.ZQCheque.Close;
  DM_Tabelas.ZQCheque.SQL.Clear;
  DM_Tabelas.ZQCheque.SQL.Add('select Banco,Dono,CPF_CNPJ,valor,emissao,deposito,alias,prorrogado,conta,idrecebimento,idpagamento,idcheque,idvinculo,numero,agencia,Substituicao,sq from cheque_rec limit 0, 1');
  DM_Tabelas.ZQCheque.open;


  DM_tabelas.contador_bx.open;
  DM_Tabelas.ZQMovBancaria.open;
  DM_tabelas.ZQCheque.open;
  DM_tabelas.ZQCaixa.open;


  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  JDEntrada.DateText := datetostr(Date);
  JDBaixa.DateText := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.DateText := datetostr(Date);
  DTPFim.DateText := datetostr(Date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
{  Enomes.Text   := DM_tabelas.ZQRecebimentonomecli.Value;
  Ecodnome.Text := DM_tabelas.ZQRecebimentocliente.Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimentoadversanome.Value;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimentoadversanome.Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimentoadversa.Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimentoadversa.Text;
  varrecpag := DM_Tabelas.ZQRecebimentoRecPag.Value;
 }

  CBPeriodoEnt.ItemIndex := 0;
//  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;
  elote.SetFocus;
end;

procedure TFrmRecebBaixa_subst.DBGBaixandoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);

begin
  if DBEidreg.Text = ZQRecBai.FieldByName('idrecebimento').Text Then Begin
    DBGBaixando.Canvas.Brush.Color := $006CFFFF;
    DBGBaixando.Canvas.Font.Color := $00A80000;
  end;
  DBGBaixando.DefaultDrawDataCell(Rect, DBGBaixando.columns[datacol].field, State);
  if Column.Field = ZQRecBai.FieldByName('marcar') then
  begin
    DBGBaixando.Canvas.FillRect(Rect);
    if ZQRecBai.FieldByName('marcar').AsString = '0' then
      ImageList1.Draw(DBGBaixando.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGBaixando.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;
end;

procedure TFrmRecebBaixa_subst.DBGBaixandoColEnter(Sender: TObject);
begin
  if DBGBaixando.SelectedField = ZQRecBai.FieldByName('marcar') then
    DBGBaixando.Options := DBGBaixando.Options - [dgEditing]
  else
    DBGBaixando.Options := DBGBaixando.Options + [dgEditing];
end;

procedure TFrmRecebBaixa_subst.DBGBaixandoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  { Depois de gravar a substituicao, a grade pode receber um clique antes de
    ter uma coluna/campo selecionado. Nao acessar SelectedField nesse estado. }
  if (DBGBaixando = nil) or (DBGBaixando.SelectedField = nil) or
     (ZQRecBai = nil) or (not ZQRecBai.Active) or ZQRecBai.IsEmpty then
    Exit;

  if SameText(DBGBaixando.SelectedField.FieldName, 'marcar') Then Begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').AsString = '0' Then
      ZQRecBai.FieldByName('marcar').AsString := '1'
    else
      ZQRecBai.FieldByName('marcar').AsString := '0';
  //  Atualiza_tela;
  end;
end;

procedure TFrmRecebBaixa_subst.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    close;
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmRecebBaixa_subst.XNERecebidoExit(Sender: TObject);
begin
   if XNEMarcado.Value >= XNERecebido.Value Then
     XNEResto.Value := XNEMarcado.Value - (XNERecebido.Value + XNEDesconto.Value)
   else
     XNEJuros.Value :=  XNERecebido.Value - XNEMarcado.Value;
end;

procedure TFrmRecebBaixa_subst.DXBBaixaGravarClick(Sender: TObject);
Var
  Varpago, percentbaixa, baixatot, varjuros, varx,
  varsomarepasse, varsomarepasseautomatico : Double;
  contador,varreg : Integer;
  ctbl,varordem, varnumordem1, varnumordem2 : string;
begin
  if empty(CBTipobaixa.Text) then
  begin
    showmessage('Selecione o tipo de operação !');
    CBTipobaixa.SetFocus;
    exit;
  end;
  ZQReceb_Baixa.close;
  ZQReceb_Baixa.SQL.Clear;
  ZQReceb_Baixa.SQL.Add('delete from Receb_Baixa where docum='+quotedstr(alltrim(CDSParcelasDocum.Value))+' and sq='+quotedstr(alltrim(inttostr(CDSParcelassq.value))));
  ZQReceb_Baixa.ExecSQL;

{  ZQReceb_Baixa.Filtered:=false;
  ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(CDSParcelasDocum.Value))+' and sq='+quotedstr(alltrim(inttostr(CDSParcelassq.value)));
  ZQReceb_Baixa.Filtered:=true;
  if ZQReceb_Baixa.RecordCount>0 then
     ZQReceb_Baixa.Delete;
  ZQReceb_Baixa.Filtered:=false;
  ZQReceb_Baixa.close;}

  DXBBaixaGravar.Enabled := False;
  DXBFechar.Enabled := False;
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  varnumordem1 := '';
  varnumordem2 := '';
  If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
    DM_tabelas.ZQCobaRe.Edit;
  DM_tabelas.ZQCobaRe.FieldByName('entrada').AsDateTime := strtodate(JDEntrada.DateText);
  DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').AsString := CBTipobaixa.Text;
  DM_tabelas.ZQCobaRe.Post;
  CDSParcelas.First;
  varsomarepasse := 0;
  DM_tabelas.contador_bx.close;
  DM_tabelas.contador_bx.Open;
  DM_tabelas.contador_bx.Last;
  if DM_tabelas.contador_bx.RecordCount>0 then
     contador:=DM_tabelas.contador_bx.FieldByName('idcontador_bx').AsLargeInt
  else
    contador:=0;
  BarraVertical.progress:=0;
  BarraVertical.Visible:=true;
  CDSParcelas.DisableControls;

  while not CDSParcelas.Eof do
  begin
    inc(contador);
    BarraVertical.Progress:=CDSParcelas.RecNo * 100 div CDSParcelas.RecordCount;
    if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
    begin
      if CDSParcelascodContabil.Value=0 then
         ctbl:='0'
      else
         ctbl:=CDSParcelascodContabil.Text;

      DM_tabelas.ZQReceb_Baixa.Close;
      DM_tabelas.ZQReceb_Baixa.SQL.Clear;

      DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,Juros_Vr,Desc_Vr,refbaixa,codcontabil,Obsebx,valor_parcela,substituicao,sq) VALUES ('+
                                        quotedstr(TrocaVirgPPto(TrimChar(CDSParcelasVrParc.Text,'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDBaixa.Datetext)))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                        quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(CDSParcelasDocum.Value)+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasjuros.Value),'.')))+','+
                                        quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasdesconto.Value),'.')))+','+quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(ctbl)+','+quotedstr(EMObsbx.Text)+','+
                                        quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value),'.')))+',''S'','+quotedstr(inttostr(contador))+')');
      DM_tabelas.ZQReceb_Baixa.ExecSQL;



{      DM_tabelas.ZQReceb_Baixa.Insert;
      DM_tabelas.ZQReceb_BaixaVr_rec.Value := CDSParcelasVrParc.Value;
      DM_tabelas.ZQReceb_Baixavalor_parcela.Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);
      DM_tabelas.ZQReceb_BaixaDt_rec.Value := strtodate(JDBaixa.Datetext); //DM_tabelas.ZQCobareHoje.Value;
      DM_TAbelas.ZQReceb_Baixadataref.Value := strtodate(JDEntrada.Datetext);
      DM_TAbelas.ZQReceb_Baixavencimento.Value := CDSParcelasVenci.Value;
      DM_tabelas.ZQReceb_BaixaTipdoc.Value := CDSParcelasTipDoc.Value;
      DM_tabelas.ZQReceb_BaixaDocum.Value  := CDSParcelasDocum.Value;
      DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
      DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
      DM_tabelas.ZQReceb_Baixarefbaixa.Value := DM_tabelas.ZQCobaRecotagem.Value;
      DM_tabelas.ZQReceb_Baixacodcontabil.Value := CDSParcelascodContabil.Value;
      DM_tabelas.ZQReceb_BaixaObsebx.Value := EMObsbx.Text;
      DM_tabelas.ZQReceb_BaixaDocum.Value:=CDSParcelasDocum.Value;
      DM_tabelas.ZQReceb_Baixasubstituicao.Value:='S';
      DM_tabelas.ZQReceb_Baixasq.Value := contador;
      DM_tabelas.ZQReceb_Baixa.Post;               }

    end
    else
    begin

      if CDSParcelascodContabil.Value=0 then
         ctbl:='0'
      else
         ctbl:=CDSParcelascodContabil.Text;
      DM_tabelas.ZQReceb_Baixa.Close;
      DM_tabelas.ZQReceb_Baixa.SQL.Clear;

      DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT INTO  Receb_Baixa (Vr_ab,dataref,vencimento,Tipdoc,Docum,refbaixa,codcontabil,Obsebx,valor_parcela,substituicao,sq) VALUES ('+
                                        quotedstr(TrocaVirgPPto(TrimChar(CDSParcelasVrParc.Text,'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                        quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(CDSParcelasDocum.Value)+','+
                                        quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(ctbl)+','+quotedstr(EMObsbx.Text)+','+
                                        quotedstr(TrocaVirgPPto(TrimChar(floattostr((CDSParcelasVrParc.Value+CDSParcelasdesconto.Value)-CDSParcelasjuros.value),'.')))+',''S'','+quotedstr(inttostr(contador))+')');
      DM_tabelas.ZQReceb_Baixa.ExecSQL;


{          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQReceb_BaixaVr_ab.Value := CDSParcelasVrParc.Value;
          DM_TAbelas.ZQReceb_Baixadataref.Value := strtodate(JDEntrada.Datetext);
          DM_TAbelas.ZQReceb_Baixavencimento.Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQReceb_BaixaTipdoc.Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQReceb_BaixaDocum.Value  := CDSParcelasDocum.Value;
          DM_tabelas.ZQReceb_Baixarefbaixa.Value := DM_tabelas.ZQCobaRecotagem.Value;
          DM_tabelas.ZQReceb_Baixacodcontabil.Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQReceb_BaixaObsebx.Value := EMObsbx.Text;
          DM_tabelas.ZQReceb_Baixavalor_parcela.Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);
          DM_tabelas.ZQReceb_Baixasubstituicao.Value:='S';
          DM_tabelas.ZQReceb_Baixasq.Value := contador;
          DM_tabelas.ZQReceb_Baixa.Post;}

    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').AsString='S' Then
      varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;

    if empty(varnumordem1) Then
    Begin
      DM_Tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
      varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
    end;



    if CDSParcelascodContabil.Value=0 then
       ctbl:='0'
    else
       ctbl:=CDSParcelascodContabil.Text;
    Dm_tabelas.ZQLoteamento.Filtered:=false;
    Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
    Dm_tabelas.ZQLoteamento.Filtered:=true;


    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (documento,cliente,dt_entrada,dt_vencimento,Valor,VrDoc,contabil,usuario,ordem,Tipdoc,nomeadversa,venda_idvenda,quadralote,idloteamento,saldo,refvinda,Observ,recpag,'+
                                     'adversa,numordem,substituicao,sq,Reajustado,Data_reajuste,Proximo_Reajuste,Parcelas_fixas,origem) VALUES ('+
                                     quotedstr(CDSParcelasDocum.Value)+','+quotedstr(Ecodnome.Text)+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime))+','+quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+
                                     quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNERecebido.Value),'.')))+','+
                                     quotedstr(ctbl)+','+quotedstr(inttostr(DM_tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt))+','+quotedstr(varnumordem1+'-'+CDSParcelasOrdem.Value)+','+
                                     quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(EAdversa.text)+','+quotedstr(inttostr(CDSParcelasvenda_idvenda.Value))+','+quotedstr(CDSParcelasquadralote.Value)+','+
                                     quotedstr(inttostr(Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+','+
                                     quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(EMNovostit.Text)+','+quotedstr(ZQRecBai.FieldByName('recpag').AsString)+','+quotedstr(EcodAdversaNova.text)+','+quotedstr(varnumordem1)+',''S'','+quotedstr(inttostr(contador))+','+
                                     quotedstr(ZQRecBai.FieldByName('Reajustado').AsString)+','+quotedstr(FormatDateTime('YYYY-MM-DD',ZQRecBai.FieldByName('Data_reajuste').AsDateTime))+','+quotedstr(ZQRecBai.FieldByName('Proximo_Reajuste').AsString)+','+quotedstr(ZQRecBai.FieldByName('Parcelas_fixas').AsString)+','+
                                     quotedstr(CDSParcelasorigem.Value)+')');

    DM_tabelas.ZQRecebimento.ExecSQL;
    Dm_tabelas.ZQLoteamento.Filtered:=false;

    //pegar o ultimo registro auto incremento criado
    DM_tabelas.ZQRecebimento.Close;
    ZQRecebe.Close;
    ZQRecebe.SQL.Clear;
    ZQRecebe.SQL.Add('Select Last_Insert_ID() as ID from recebimento group by id');
//    ZQRecebe.SQL.Add('Select idrecebimento from Recebimento order by idrecebimento');    
    ZQRecebe.open;
    ZQRecebe.last;
    varreg := ZQRecebe.FieldByName('ID').AsLargeInt;


{    DM_tabelas.ZQRecebimento.Insert;
    DM_tabelas.ZQRecebimentodocumento.Value := CDSParcelasDocum.Value;
    DM_tabelas.ZQRecebimentocliente.Value := strtoint(Ecodnome.Text);
    DM_tabelas.ZQRecebimentoDt_Entrada.Value := DM_tabelas.ZQCobareHoje.Value;
    DM_tabelas.ZQRecebimentoDt_Vencimento.Value := CDSParcelasVenci.Value;
    DM_tabelas.ZQRecebimentoValor.Value := CDSParcelasVrParc.Value;
    DM_tabelas.ZQRecebimentoVrDoc.Value := XNERecebido.Value;
    DM_tabelas.ZQRecebimentocontabil.Value := CDSParcelascodContabil.Value;
    DM_tabelas.ZQRecebimentousuario.Value := DM_tabelas.ZQUsuarioidusuario.Value;
    DM_tabelas.ZQRecebimentoordem.Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
    DM_tabelas.ZQRecebimentoTipDoc.Value := CDSParcelasTipDoc.Value;
    DM_Tabelas.ZQRecebimentonomeadversa.Value := EAdversaNova.text;

    DM_tabelas.ZQRecebimentovenda_idvenda.Value:=CDSParcelasvenda_idvenda.Value;
    DM_tabelas.ZQRecebimentoquadralote.Value:=CDSParcelasquadralote.Value;

    DM_tabelas.ZQRecebimentoDocumento.Value:=CDSParcelasDocum.Value;
    Dm_tabelas.ZQLoteamento.Filtered:=false;
    Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
    Dm_tabelas.ZQLoteamento.Filtered:=true;
    if Dm_tabelas.ZQLoteamento.RecordCount>0 then
       Dm_tabelas.ZQRecebimentoidloteamento.Value:=Dm_tabelas.ZQLoteamentoidloteamento.Value;
    Dm_tabelas.ZQLoteamento.Filtered:=false;

   //     if (CDSParcelasTipDoc.Value<>'CH') then
    DM_tabelas.ZQRecebimentosaldo.Value := CDSParcelasVrParc.Value;

    DM_tabelas.ZQRecebimentorefvinda.Value := DM_tabelas.ZQCobaRecotagem.Value;;
    DM_tabelas.ZQRecebimentoObserv.Value := EMNovostit.Text;
    DM_Tabelas.ZQRecebimentorecpag.Value := ZQRecBairecpag.Value;
    DM_Tabelas.ZQRecebimentoadversa.Value := strtoint(EcodAdversaNova.text);
    DM_Tabelas.ZQRecebimentonumordem.Value := strtoint(varnumordem1);

    DM_Tabelas.ZQRecebimentosubstituicao.Value:='S';
    DM_tabelas.ZQRecebimentosq.Value := contador;
    DM_Tabelas.ZQRecebimentoParcelas_fixas.Value:=CDSParcelasParcelas_fixas.Value;
    DM_Tabelas.ZQRecebimentoorigem.Value:=CDSParcelasorigem.Value;

    DM_tabelas.ZQRecebimento.Post;}

    if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString='S' Then
    Begin

        if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
        begin
          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').AsInteger := CDSParcelascodContabil.Value;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').AsDateTime := date;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').AsFloat := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCaixa.FieldByName('documento').AsString := CDSParcelasDocum.Value;
          DM_tabelas.ZQCaixa.FieldByName('empresa').AsInteger := 1;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
          DM_tabelas.ZQCaixa.FieldByName('credeb').AsString := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').AsString;
          DM_tabelas.ZQCaixa.FieldByName('vinculo').AsString := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').AsString = 'P' Then
            DM_tabelas.ZQCaixa.FieldByName('vinculo').AsString := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_tabelas.ZQCaixa.FieldByName('contracodigo').AsInteger := strtoint(Ecodnome.Text);
          DM_tabelas.ZQCaixa.FieldByName('obs').AsString := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
          IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
            DM_tabelas.ZQCaixa.FieldByName('obs').AsString :=DM_tabelas.ZQCaixa.FieldByName('obs').AsString + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
          end;
          DM_tabelas.ZQCaixa.FieldByName('obs').AsString :=DM_tabelas.ZQCaixa.FieldByName('obs').AsString +' | '+EMObsbx.Text;
          DM_tabelas.ZQCaixa.FieldByName('Substituicao').AsString := 'S';
          DM_tabelas.ZQCaixa.FieldByName('sq').AsLargeInt := contador;
          DM_tabelas.ZQCaixa.Post;

        end;
    end;
    IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    Begin

        DM_tabelas.ZQCheque.Insert;
        DM_tabelas.ZQCheque.FieldByName('Banco').AsString := CdSParcelasBanco.Value;
        DM_tabelas.ZQCheque.FieldByName('agencia').AsString := CDSParcelasagencia.Value;
        DM_tabelas.ZQCheque.FieldByName('Dono').AsString := CDSParcelasdono.Value;
        DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString := CDSParcelasCPF_CNPJ.Value;
        DM_tabelas.ZQCheque.FieldByName('valor').AsFloat := CDSParcelasVrparc.Value;
        DM_tabelas.ZQCheque.FieldByName('emissao').AsDateTime := strtodate(JDEntrada.DateText);
        DM_Tabelas.ZQCheque.FieldByName('conta').AsString := CDSParcelasConta.Value;
        DM_tabelas.ZQCheque.FieldByName('deposito').AsDateTime := CDSParcelasVenci.Value;
        DM_tabelas.ZQCheque.FieldByName('numero').AsString := CDSParcelasNcheque.Value;        
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := varreg;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').AsString := 'S';
        DM_tabelas.ZQCheque.FieldByName('sq').AsLargeInt := contador;
        DM_tabelas.ZQCheque.Post;

    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').AsString = 'S' Then Begin
        DM_Tabelas.ZQMovBancaria.Insert;
        DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').AsInteger := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
        DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').AsInteger := CDSParcelascontabanco.Value;
        DM_TAbelas.ZQMovBancaria.FieldByName('documento').AsString := CDSParcelasDocum.Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').AsDateTime := date;
        DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').AsDateTime := Time;
        DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').AsInteger := strtoint(Ecodnome.Text);
        DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').AsDateTime := CDSParcelasVenci.Value;
        DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').AsFloat := CDSParcelasVrParc.Value;
        DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
        DM_TAbelas.ZQMovBancaria.FieldByName('credeb').AsString := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').AsString;
        DM_Tabelas.ZQMovBancaria.FieldByName('obs').AsString := EMObsbx.Text;
        DM_tabelas.ZQMovBancaria.FieldByName('vinculo').AsString := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        IF ZQRecBai.FieldByName('recpag').AsString = 'P' Then
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').AsString := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').AsString := 'S';
        DM_tabelas.ZQMovBancaria.FieldByName('sq').AsLargeInt := contador;
        DM_Tabelas.ZQMovBancaria.Post;
    end;
    CDSParcelas.Next;
  end;
  CDSParcelas.EnableControls;
  ZQRecebe.close;
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where 1=0 order by DT_Vencimento ');
  DM_tabelas.ZQRecebimento.open;

  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa');
  DM_tabelas.ZQReceb_Baixa.open;

  CDSParcelas.EnableControls;
  BarraVertical.Progress:=0;
  BarraVertical.Visible:=false;


  DM_tabelas.contador_bx.insert;
  DM_tabelas.contador_bx.FieldByName('idcontador_bx').AsLargeInt:=contador;
  DM_tabelas.contador_bx.Post;
  DM_tabelas.contador_bx.close;

  if not empty(varnumordem1) Then Begin
    ZQDiv_usado.SQL.Clear;
    ZQDiv_usado.SQL.Add('select iddividido,participante,percentual,planodeconta,tipodocumento,ordem from dividido');
    ZQDiv_usado.SQL.Add(' where  ordem = '+quotedstr(ZQRecBai.FieldByName('numordem').Text));
    ZQDiv_usado.Open;
    varsomarepasseautomatico := 0;
    ZQDiv_usado.DisableControls;
    while not ZQDiv_usado.Eof do begin
      DM_Tabelas.ZQDividido.Insert;
      DM_Tabelas.ZQDividido.FieldByName('participante').AsInteger := ZQDiv_usado.FieldByName('participante').AsInteger;
      DM_Tabelas.ZQDividido.FieldByName('percentual').AsFloat := ZQDiv_usado.FieldByName('percentual').AsFloat;
      DM_Tabelas.ZQDividido.FieldByName('planodeconta').AsInteger := ZQDiv_usado.FieldByName('planodeconta').AsInteger;
      DM_Tabelas.ZQDividido.FieldByName('tipodocumento').AsString := ZQDiv_usado.FieldByName('tipodocumento').AsString;
      DM_Tabelas.ZQDividido.FieldByName('ordem').AsInteger := strtoint(varnumordem1);
      DM_Tabelas.ZQDividido.Post;
      ZQDiv_usado.Next;
    end;
  end;
  ZQDiv_usado.EnableControls;
  ZQRecBai.First;
  varpago := 0;
  Memo1.Clear;
  baixatot := XNERecebido.Value;

  //aqui que baixa o titulo
  While not ZqRecBai.Eof do
  Begin

    ZQSubstituicao.Close;
    ZQSubstituicao.SQL.clear;
    ZQSubstituicao.SQL.Add('select lancamento,conta_bancaria_cod_banco,plano_contas_codigo,documento,dt_lanc,hora_lanc,dt_conciliado,vr_lanc,credeb,saldo_lanc,saldo_conci,obs,contraquem,');
    ZQSubstituicao.SQL.Add('       contracodigo,vinculo,Substituicao,sq ');
    ZQSubstituicao.SQL.Add(' from  mov_banco where SUBSTR(vinculo,2,20)='+quotedstr(ZQRecBai.FieldByName('refvinda').Text));
//    ZQSubstituicao.SQL.Add('select * from recebimento as rc join cheque_rec as ch on ch.idrecebimento=rc.idrecebimento join caixa as cx on SUBSTR(cx.vinculo,2,20)=rc.refvinda join recbxhist as rh on rh.idrecib=rc.idrecebimento join receb_baixa as rx on rx.refbaixa=rh.refer join mov_banco as mv on SUBSTR(mv.vinculo,2,20)=rc.refvinda where idrecebimento='+quotedstr(ZQRecBaiidrecebimento.text);
    ZQSubstituicao.open;
    ZQSubstituicao.First;
    if ZQSubstituicao.RecordCount>0 then
    begin
      while not ZQSubstituicao.Eof do
      begin
        ZQSubstituicao.Delete;
      end;
    end;
    ZQbaixa.Close;
    ZQbaixa.SQL.clear;
    ZQbaixa.SQL.Add('Select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from recbxhist where idrecib='+quotedstr(ZQRecBai.FieldByName('idrecebimento').text));
    ZQbaixa.open;
    ZQbaixa.First;
    if ZQbaixa.RecordCount>0 then
    begin
      while not ZQbaixa.Eof do
      begin
        ZQBaixa2.Close;
        ZQBaixa2.SQL.Clear;
        ZQBaixa2.SQL.Add('delete from receb_baixa where refbaixa='+quotedstr(ZQBaixa.FieldByName('refer').Text));
        ZQBaixa2.ExecSQL;
        ZQbaixa.Delete;
      end;
    end;
    ZQSubstituicao.Close;
    ZQSubstituicao.SQL.clear;
    ZQSubstituicao.SQL.Add('select codigo,plano_contas_codigo,data_lan,vr_lan,obs,documento,empresa,vinculo,saldo,credeb,contraquem,contracodigo,Substituicao,sq from  caixa where SUBSTR(vinculo,2,20)='+quotedstr(ZQRecBai.FieldByName('refvinda').Text));
    ZQSubstituicao.open;
    ZQSubstituicao.First;
    if ZQSubstituicao.RecordCount>0 then
    begin
      while not ZQSubstituicao.Eof do
      begin
        ZQSubstituicao.Delete;
      end;
    end;
    ZQSubstituicao.Close;
    ZQSubstituicao.SQL.clear;
    ZQSubstituicao.SQL.Add('delete from  cheque_rec where idrecebimento='+quotedstr(ZQRecBai.FieldByName('idrecebimento').text));
    ZQSubstituicao.ExecSQL;
    ZqRecBai.delete;
  end;
  ZQSubstituicao.Close;
  ZQbaixa.Close;
  ZQbaixa2.Close;
{  CDS_dividido.Close;
  CDS_dividido.CreateDataSet;
  ZQRatear.SQL.Clear;
  ZQRatear.SQL.Add('select idrecebimento, h.valor, percent_usado, b.vr_rec, participante, percentual,planodeconta, dv.tipodocumento, re.ordem, dv.ordem from recbxhist as H');
  ZQRatear.SQL.Add(' join recebimento as re ON H.idrecib=RE.idrecebimento');
  ZQRatear.SQL.Add(' join dividido as dv ON dv.ordem=numordem');
  ZQRatear.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
  ZQRatear.SQL.Add(' join tipodocumento as D ON B.tipdoc=D.tipodoc');
  ZQRatear.SQL.Add(' where somapaga = '+quotedstr('S')+' and refer = '+DM_tabelas.ZQCobaRecotagem.Text);
  ZQRatear.Open;
  varsomarepasseautomatico := 0;
  while not ZQRatear.Eof do begin
    if CDS_dividido.Locate('participante;contabil;tipodoc',VarArrayOf([ZQRatearparticipante.Value,ZQRatearplanodeconta.Value,ZQRateartipodocumento.Value]),[]) then
      CDS_dividido.Edit
    else
      CDS_dividido.Insert;
    CDS_divididoparticipante.Value := ZQRatearparticipante.Value;
    CDS_divididovr_base.Value := CDS_divididovr_base.Value + ZQRatearvr_base.Value;
    CDS_divididovr_calc.Value := CDS_divididovr_calc.Value + ZQRatearvr_divi.Value;
    CDS_divididocontabil.Value := ZQRatearplanodeconta.Value;
    CDS_divididotipodoc.Value := ZQRateartipodocumento.Value;
    CDS_divididoobs.Value := CDS_divididoobs.Value+ZQRatearidrecebimento.Text+' | '+ZQRatearordem.Value+' | '+ZQRatearvr_base.Text+' x '+ZQRatearpercentual.Text+' = '+ZQRatearvr_divi.Text+chr(13);
    varsomarepasseautomatico := varsomarepasseautomatico + ZQRatearvr_divi.Value;
    ZQRatear.Next;
  end;
  CDS_dividido.First;
  if CDS_dividido.RecordCount>0 Then Begin
    DM_Tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
  end;
  while not CDS_dividido.Eof do begin
    DM_tabelas.ZQRecebimento.Insert;
    DM_tabelas.ZQRecebimentodocumento.Value := 'Rateio Bx - '+DM_tabelas.ZQCobaRecotagem.Text;
    DM_tabelas.ZQRecebimentocliente.Value :=  CDS_divididoparticipante.Value;
    DM_tabelas.ZQRecebimentoDt_Entrada.Value := strtodate(JDEntrada.DateText);
    DM_tabelas.ZQRecebimentoDt_Vencimento.Value := strtodate(JDEntrada.DateText);
    DM_tabelas.ZQRecebimentocontabil.Value := CDS_divididocontabil.Value;
    DM_tabelas.ZQRecebimentoObserv.Value := EMNovostit.Text+chr(13)+CDS_divididoobs.Value;
    DM_Tabelas.ZQRecebimentorecpag.Value := 'P';
    DM_tabelas.ZQRecebimentoTipDoc.Value := CDS_divididotipodoc.Value;
    DM_tabelas.ZQRecebimentoordem.Value := DM_Tabelas.ZQNumOrdemidnumordem.Text+'-'+inttostr(CDS_dividido.RecNo)+'/'+inttostr(CDS_dividido.RecordCount);
    DM_tabelas.ZQRecebimentorefvinda.Value := DM_tabelas.ZQCobaRecotagem.Value;
    DM_tabelas.ZQRecebimentoVrDoc.Value :=  CDS_divididovr_base.Value;
    DM_tabelas.ZQRecebimentoValor.Value :=  CDS_divididovr_calc.Value;
    DM_tabelas.ZQRecebimentosaldo.Value :=  DM_tabelas.ZQRecebimentoValor.Value;
    DM_Tabelas.ZQRecebimentoadversa.Value := DM_Tabelas.ZQConfiguracoespaticipantedefault.Value;

    DM_Tabelas.ZQRecebimentonomeadversa.Value := EAdversaNova.text;

    DM_tabelas.ZQRecebimento.Post;
    CDS_dividido.Next;
  end;}
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where idrecebimento='+quotedstr(varreg.ToString)+' order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;


  DM_Tabelas.ZQRecebimento.Locate('idrecebimento',varreg,[]);
  FreeAndNil(FrmImpRecibo);
  if FrmImpRecibo=nil then
     FrmImpRecibo:=TFrmImpRecibo.Create(self);
  FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
  FrmImpRecibo.Nomecli.Text := Enomes.Text;
  FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;

  ///
  ZQRecBai.Refresh;
  DBGBaixando.Refresh;
  DM_tabelas.ZQReceb_Baixa.Refresh;
  DM_Tabelas.ZQReBxHi.Refresh;
  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  XNEJuros.Value:=0;
  XNERecebido.Value:=0;
  XNEDesconto.Value:=0;
  XNEResto.Value:=0;
  XNEParcelas.Value:=1;
  CBPeriodoEnt.Text:='HOJE';
  EMObsbx.Clear;
  EMNovostit.Clear;


//  Close;
   // imprime o recibo
  //  FrmImpRecibo.showmodal;
///
{  if perguntaSN('Gerar Recibo de Quitação ?','S') Then Begin
    FrmRelRecibodeQuita.ZQForma.SQL.Clear;
    FrmRelRecibodeQuita.ZQForma.SQL.Add('Select * from receb_baixa where refbaixa='+DM_tabelas.ZQCobaRecotagem.Text);
    FrmRelRecibodeQuita.ZQForma.Open;
    FrmRelRecibodeQuita.RLLabel19.Caption := Transform(XNERecebido.Value,'###,###,###,##0.00');
    FrmRelRecibodeQuita.RLLabel54.Caption := Transform(XNERecebido.Value,'###,###,###,##0.00');
    FrmRelRecibodeQuita.RLMemo1.Lines := EMObsbx.Lines;
    FrmRelRecibodeQuita.RLMemo2.Lines := Memo1.Lines;
    FrmRelRecibodeQuita.RLMemo3.Lines := FrmRelRecibodeQuita.RLMemo2.Lines;
    FrmRelRecibodeQuita.RLLabel18.Caption := DM_tabelas.ZQCobaRecotagem.Text;
    FrmRelRecibodeQuita.RLLabel39.Caption := DM_tabelas.ZQCobaRecotagem.Text;
    ZQItensRecibo.Close;
    ZQItensRecibo.SQL[1] := 'where refer='+DM_tabelas.ZQCobaRecotagem.Text;
    ZQItensRecibo.Open;
    if FrmRelRecibodeQuita=nil then
      FrmRelRecibodeQuita:=TFrmRelRecibodeQuita.Create(Application);
    FrmRelRecibodeQuita.RLReport1.PreviewModal;
    FreeAndNil(FrmRelRecibodeQuita);
  end; }
  FreeAndNil(FrmImpRecibo);
  ZQRecBai.Filtered := False;
  DM_tabelas.ZQRecebimento.Refresh;
  Elote.SetFocus;
  DXBFechar.Enabled := true;
end;

procedure TFrmRecebBaixa_subst.DBCBTIPDOCExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if Empty(DBCBTIPDOC.Text) Then Begin
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipDoc.Text,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').AsString='S' Then Begin
      if (datetostr(CDSParcelasVenci.Value)<>JDEntrada.DateText) and ((CDSParcelasTipDoc.Value<>'DP')) Then BEgin
        showmessage('Este tipo de documento só aceita pagamento a vista!!!');
        EContabil.SetFocus;
        Exit;
      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').AsString='S' then Begin
      GBBanco.Visible := True;
      DBGBanco.SetFocus;
      exit;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S' then Begin
      GBCheque.Visible := True;
      DBEBanco.SetFocus;
      exit;
    end;
  end;
end;

procedure TFrmRecebBaixa_subst.DXBBaixaGravarEnter(Sender: TObject);
begin
  CDSParcelas.First;
  CDSParcelas.DisableControls;
  While not CDSParcelas.Eof do Begin
    if CDSParcelascodContabil.Value=0 Then Begin
      EContabil.SetFocus;
      Exit;
    end;
    if empty(CDSParcelasTipDoc.Value) Then Begin
      DBCBTIPDOC.SetFocus;
      Exit;
    end;
    CDSParcelas.Next;
  End;
  CDSParcelas.EnableControls;
end;

procedure TFrmRecebBaixa_subst.DBCBTIPDOCEnter(Sender: TObject);
Var
  VarCheque,VarTipDoc, VarDocum, VarBanco, Varconta, VarAgencia, VarCPF_CNPJ, VarDono : string;
begin
  if CDSParcelas.Recno>1  Then begin
    CDSParcelas.Prior;
    VarTipDoc := CDSParcelasTipdoc.Value;
    VarDocum := CDSParcelasDocum.Value;
    VarBanco := CDSParcelasBanco.Value;
    VarAgencia := CDSParcelasagencia.Value;
    varconta := CDSParcelasconta.Value;
    VarCPF_CNPJ := CDSParcelascpf_cnpj.Value;
    VarDono := CDSParcelasDono.Value;
    Varcheque:= CDSParcelasNcheque.Value;

    CDSParcelas.Next;
    CDSParcelas.Edit;
    if Empty(CDSParcelasTipDoc.Value) Then
      CDSParcelasTipdoc.Value := VarTipDoc;
    if Empty(CDSParcelasDocum.Value) Then
      CDSParcelasDocum.Value :=  VarDocum;
    if Empty(CDSParcelasBanco.Value) Then
      CDSParcelasBanco.Value := VarBanco ;
    if Empty(CDSParcelasagencia.Value) Then
      CDSParcelasagencia.Value := VarAgencia;
    if Empty(CDSParcelascpf_cnpj.Value) then
      CDSParcelascpf_cnpj.Value := VarCPF_CNPJ;
    if Empty(CDSParcelasDono.Value) then
      CDSParcelasDono.Value := VarDono;
    if Empty(CDSParcelasConta.Value) then
      CDSParcelasConta.Value := VarConta;
    if Empty(CDSParcelasNcheque.Value) then
      CDSParcelasNcheque.Value := VarCheque;

  end
  else
  Begin
    CDSParcelas.Edit;
    CDSParcelasDono.Value := Enomes.Text;
  End;
end;

procedure TFrmRecebBaixa_subst.DBEBaixaVrRecExit(Sender: TObject);
Var
  varregis : Integer;
  varsoma, vardif, varjur, vardesc : Double;
begin
  if Empty(WDTBaixaEntra.Text) Then
    CDSParcelasVenci.Value := date;
  varsoma := 0;
  varjur := XNEJuros.Value;
  vardesc := XNEDesconto.Value;
  varregis := CDSParcelas.RecNo;
  CDSParcelas.First;
  CDSParcelas.DisableControls;
  while not (CDSParcelas.Eof) and (CDSParcelas.RecNo <= varregis) do Begin
    VarDoc := VarDoc + quotedstr(CDSParcelasDocum.Value)+',';
    varsoma := varsoma +CDSParcelasVrParc.Value;
    CDSParcelas.Edit;
    if CDSParcelasVrParc.Value>=varjur then
      CDSParcelasjuros.Value := varjur
    else
      CDSParcelasjuros.Value := CDSParcelasVrParc.Value;
    varjur := varjur - CDSParcelasjuros.Value;
    if CDSParcelasVrParc.Value>= vardesc Then
      CDSParcelasdesconto.Value := vardesc
    else
      CDSParcelasdesconto.Value := CDSParcelasVrParc.Value;
    vardesc := vardesc - CDSParcelasdesconto.Value;
    CDSParcelas.Post;
    CDSParcelas.next;
  end;
  CDSParcelas.EnableControls;
  VarDoc := copy(VarDoc,1,length(VarDoc)-1);
  CDSParcelas.RecNo := varregis;
  vardif := ExRound(XNERecebido.Value,2) - ExRound(varsoma,2);
  if CDSParcelas.RecNO< XNEParcelas.Value Then Begin
    CDSParcelas.RecNO := varregis+1;
    CDSParcelas.DisableControls;
   // CDSParcelas.first;
    while not CDSParcelas.Eof do Begin
      CDSParcelas.Edit;
      CDSParcelasVrParc.Value := ExRound(vardif,2) / (ExRound(XNEParcelas.Value,2)-ExRound(varregis,2));
      vardesc := vardesc - CDSParcelasdesconto.Value;
      CDSParcelas.Next;
    end;
    CDSParcelas.EnableControls;
  end;
  CDSParcelas.RecNo := varregis;
  DXBBaixaGravar.Enabled := True;
  DXBBaixaGravar.SetFocus;
end;

procedure TFrmRecebBaixa_subst.EContabilExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if empty(Econtabil.text) Then Begin
      showmessage('A descrição contábil não pode ficar em branco....');
      Econtabil.SetFocus;
      exit;
    end;
    if not AchaPlanoDeConta(450, 30, 'S', EContabil.Text, '') then Begin
      EContabil.SetFocus;
      exit;
    end;
    CDSParcelas.EnableControls;
    EContabil.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    if not (CDSParcelas.State in [DsInsert, DsEdit]) Then
      CDSParcelas.Edit;
    CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
    CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
  end;
//    CDSParcelasTipDoc.Value := FrmAchaPlanodeConta.ZQTempPlanoContasdoccomum.Value;
end;

procedure TFrmRecebBaixa_subst.XDBNumEdit1Exit(Sender: TObject);
begin
  CDSParcelasVrParc.Value := XNERecebido.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmRecebBaixa_subst.DBEBaixaDocumExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
{    if (not empty(DBEBaixaDocum.Text)) and (DM_tabelas.ZQRecebimento.Locate('documento',DBEBaixaDocum.Text,[])) Then Begin
      showmessage('Este documento já foi lançado anteriormente... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    End;
    if DM_tabelas.ZQCheque.Locate('CH_Conta',DBEBaixaDocum.Text,[]) Then Begin
      showmessage('Este cheque já foi lançado anteriormente... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (pos(quotedstr(CDSParcelasDocum.Value),VarDoc)>0) Then Begin
      showmessage('Este cheque já foi lançado nesta inclusão... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (empty(CDSParcelasDocum.Value)) Then Begin
      showmessage('o campo Nº do cheque deve ser preenchido...');
      DBEBaixaDocum.SetFocus;
      exit;
    end;}
  end;
end;

procedure TFrmRecebBaixa_subst.EContabilEnter(Sender: TObject);
begin
  if empty(Econtabil.Text) Then Begin
    Econtabil.Text := DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').Text;
  end;
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;
  DBText1.Caption := CDSParcelasOrdem.Value;
end;

procedure TFrmRecebBaixa_subst.DBECpf_CnpjExit(Sender: TObject);
Var
  variavel : string;
  varx : integer;
begin
  If not DXBFechar.Focused Then Begin
    if not empty(CDSParcelascpf_cnpj.Value) then Begin
      variavel := '';
      for varx :=1 to length(CDSParcelascpf_cnpj.Value) do
        if pos(copy(CDSParcelascpf_cnpj.Value,varx,1),'0123456789')>0 then
          variavel := variavel + copy(CDSParcelascpf_cnpj.Value,varx,1);
      if not Verif_doc(variavel,True) then
        DBECpf_Cnpj.SetFocus
    end;
  end;
end;

procedure TFrmRecebBaixa_subst.EnomesExit(Sender: TObject);
begin
  If not DXBFechar.Focused Then Begin
    if not AchaCLiente(450, 140,Enomes.Text) Then Begin
      Enomes.SetFocus;
      Exit;
    End;
    Ecodnome.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
    Enomes.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
  End;
end;

procedure TFrmRecebBaixa_subst.Label1Click(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
end;

procedure TFrmRecebBaixa_subst.RBChequeClick(Sender: TObject);
begin
  Label1.Visible := true;
  DTPIni.Visible := true;
  Label23.Visible := true;
  DTPFim.Visible := true;
  Enomes.Visible := False;
end;

procedure TFrmRecebBaixa_subst.GBBancoExit(Sender: TObject);
begin
  GBBanco.Visible := False;
  if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S' then Begin
    CDSParcelasagencia.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString;
    CDSParcelasConta.Value := DM_TAbelas.ZQContaBancaria.FieldByName('n_conta').AsString;
    CDSParcelasbanco.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString;
    CDSParcelasdono.Value := DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').AsString;
    GBCheque.Visible := True;
    DBEBanco.SetFocus;
    exit;
  end;
end;

procedure TFrmRecebBaixa_subst.GBChequeExit(Sender: TObject);
begin
  GBCheque.Visible := False;
end;

procedure TFrmRecebBaixa_subst.EPlanocontabancoExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not AchaPlanoDeConta(450, 30, 'S', EPlanocontabanco.Text, '') then Begin
      EPlanocontabanco.SetFocus;
      exit;
    end;
    EPlanocontabanco.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    CDSParcelascontabanco.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
  end;
end;

procedure TFrmRecebBaixa_subst.filtrabaixa;
var
  vartipo : string;
  posi:integer;
Begin
  Label37.Visible:=true;
  pini.Visible:=true;
  Label38.Visible:=true;
  pfim.Visible:=true;

  if ( not empty(alltrim(Elote.Text))) then
  begin
    ZQRecBai.close;
    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    begin
       ZQRecBai.Close;
       ZQRecBai.SQL.Clear;
       ZQRecBai.SQL.Add('update recebimento set marcar= case when (dt_vencimento<= :dt) then ''0'' when (dt_vencimento> :dt) then ''1'' end  ');
       ZQRecBai.SQL.Add(' where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+Ecodnome.Text+')');
       if not empty(EcodAdversa.text) Then
       Begin
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
       end;
       if not empty(EDLoteamento.Text) then
       begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          ZQRecBai.SQL.Add(' and (idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+')');
       end;
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and (Dt_Vencimento>= :dt1) and (Dt_Vencimento<= :dt2)');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
       ZQRecBai.ParamByName('dt').AsDate:=date;
       ZQRecBai.ExecSQL;


       ZQRecBai.Close;
       ZQRecBai.SQL.Clear;
       ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
       ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
       ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
       ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+Ecodnome.Text+')');
       if not empty(EcodAdversa.text) Then
       Begin
         ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
       end;
       if not empty(EDLoteamento.Text) then
       begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          ZQRecBai.SQL.Add(' and (idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+')');
       end;

       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
       ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
       ZQRecBai.Open;

    end
    else if RBCheque.Checked then
    Begin
      Label37.Visible:=false;
      pini.Visible:=false;
      pini.Clear;
      Label38.Visible:=false;
      pfim.Visible:=false;
      pfim.Clear;
      vartipo := '';
      DM_TAbelas.ZQTipodoc.First;
      DM_TAbelas.ZQTipodoc.DisableControls;
      while not DM_Tabelas.ZQTipodoc.Eof do begin
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
        DM_Tabelas.ZQTipodoc.Next;
      end;
      DM_TAbelas.ZQTipodoc.EnableControls;
      // Evita gerar IN () quando nenhum tipo de documento de cheque estiver configurado.
      if vartipo = '' then
        vartipo := quotedstr('')
      else
        vartipo := copy(vartipo,1,length(vartipo)-1);



      ZQRecBai.Close;
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('update recebimento set marcar= case when (dt_vencimento<= :dt) then ''0'' when (dt_vencimento> :dt) then ''1'' end  ');
      ZQRecBai.SQL.Add(' where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
      ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
      ZQRecBai.ParamByName('dt').AsDate:=date;
      ZQRecBai.ExecSQL;

      ZQRecBai.Close;
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
      ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
      ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
      ZQRecBai.Open;
    end;
//    ZQRecBai.Open;
    XEEmpresend.Text := ZQRecBai.FieldByName('nomeempreend').AsString;
    LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';

{    if RBParticipante.Checked Then
    begin
      while not ZQRecBai.Eof do begin
        application.ProcessMessages;
        ZQRecBai.Edit;
        if ZQRecBaiDt_Vencimento.Value<=date Then
         ZQRecBaimarcar.Value := '0'
         else
          ZQRecBaimarcar.Value := '1';
        ZQRecBai.Post;
        ZQRecBai.Next;
      end;
    end;}

    ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]);
    if RBPonteiro.Checked Then Begin
       ZQRecBai.Edit;
       ZQRecBai.FieldByName('marcar').AsString := '0';
       ZQRecBai.Post;
    end;
  end
  else
  begin

    ZQRecBai.close;
    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    begin
       ZQRecBai.Close;
       ZQRecBai.SQL.Clear;
       ZQRecBai.SQL.Add('update recebimento set marcar= case when (dt_vencimento<= :dt) then ''0'' when (dt_vencimento> :dt) then ''1'' end  ');
       ZQRecBai.SQL.Add(' where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null))');
       if not empty(EcodAdversa.text) Then
       Begin
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
       end;
       if not empty(EDLoteamento.Text) then
       begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          ZQRecBai.SQL.Add(' and (idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+')');
       end;
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and (Dt_Vencimento>= :dt1) and (Dt_Vencimento<= :dt2)');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
       ZQRecBai.ParamByName('dt').AsDate:=date;
       ZQRecBai.ExecSQL;


       ZQRecBai.Close;
       ZQRecBai.SQL.Clear;
       ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
       ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
       ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
       ZQRecBai.SQL.Add(' from recebimento where recpag='+quotedstr(varrecpag)+' and ((saldo>0) and  (not saldo is null)) ');
       if not empty(EcodAdversa.text) Then
       Begin
          ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
       end;
       if not empty(EDLoteamento.Text) then
       begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          ZQRecBai.SQL.Add(' and (idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+')');
       end;

       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
       ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
       ZQRecBai.Open;

    end
    else if RBCheque.Checked then
    Begin

      Label37.Visible:=false;
      pini.Visible:=false;
      pini.Clear;
      Label38.Visible:=false;
      pfim.Visible:=false;
      pfim.Clear;
      vartipo := '';
      DM_TAbelas.ZQTipodoc.First;
      while not DM_Tabelas.ZQTipodoc.Eof do begin
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
        DM_Tabelas.ZQTipodoc.Next;
      end;
      vartipo := copy(vartipo,1,length(vartipo)-1);



      ZQRecBai.Close;
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('update recebimento set marcar= case when (dt_vencimento<= :dt) then ''0'' when (dt_vencimento> :dt) then ''1'' end  ');
      ZQRecBai.SQL.Add(' where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
      if not empty(EcodAdversa.text) Then
      Begin
        ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
      end;
      if not empty(EDLoteamento.Text) then
      begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          ZQRecBai.SQL.Add(' and (idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+')');
      end;

      if pfim.Text<>'  /  /    ' then
      begin
         ZQRecBai.SQL.Add(' and (dt_vencimento between :dt5 and :dt6)');
         ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
         ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
      end;
      ZQRecBai.ParamByName('dt').AsDate:=date;
      ZQRecBai.ExecSQL;


      ZQRecBai.Close;
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
      ZQRecBai.SQL.Add(' from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
      if not empty(EcodAdversa.text) Then
      Begin
        ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
      end;
       if not empty(EDLoteamento.Text) then
       begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          ZQRecBai.SQL.Add(' and (idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+')');
       end;
      if pfim.Text<>'  /  /    ' then
      begin
         ZQRecBai.SQL.Add(' and (dt_vencimento between :dt5 and :dt6)');
         ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
         ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
      end;
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
      ZQRecBai.Open;
    end;
//    ZQRecBai.Open;
    XEEmpresend.Text := ZQRecBai.FieldByName('nomeempreend').AsString;
    LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';

{    if RBParticipante.Checked Then
    begin
      while not ZQRecBai.Eof do begin
        application.ProcessMessages;
        ZQRecBai.Edit;
        if ZQRecBaiDt_Vencimento.Value<=date Then
         ZQRecBaimarcar.Value := '0'
         else
          ZQRecBaimarcar.Value := '1';
        ZQRecBai.Post;
        ZQRecBai.Next;
      end;
    end;}

    ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]);
    if RBPonteiro.Checked Then Begin
       ZQRecBai.Edit;
       ZQRecBai.FieldByName('marcar').AsString := '0';
       ZQRecBai.Post;
    end;

  end;
end;

procedure TFrmRecebBaixa_subst.GBContraExit(Sender: TObject);
begin
  if not RBPonteiro.Checked Then
    Filtrabaixa;
end;

procedure TFrmRecebBaixa_subst.RBParticipanteClick(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
  Enomes.SetFocus;
end;

procedure TFrmRecebBaixa_subst.RBPonteiroClick(Sender: TObject);
Var
  VArreg : integer;
begin
  if ZQRecBai.active=true then
  begin
    Label1.Visible := False;
    DTPIni.Visible := False;
    Label23.Visible := False;
    DTPFim.Visible := False;
    Enomes.Visible := True;
    varreg := ZQRecBai.RecNo;
    ZQRecBai.First;
    while not ZQRecBai.Eof do begin
      ZQRecBai.Edit;
      ZQRecBai.FieldByName('marcar').AsString := '1';
      ZQRecBai.Post;
      ZQRecBai.Next;
    end;
    ZQRecBai.Edit;
    ZQRecBai.RecNo := varreg;
    ZQRecBai.FieldByName('marcar').AsString := '0';
    ZQRecBai.Post;
    XNEMarcado.Value := ZQRecBai.FieldByName('saldo').AsFloat;
    Atualiza_tela;
  end;
end;

procedure TFrmRecebBaixa_subst.XDBNumEdit1Enter(Sender: TObject);
begin
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;

end;

procedure TFrmRecebBaixa_subst.JDEntradaExit(Sender: TObject);
begin
  if not (DXBFechar.Focused or GBContra.Focused or CBSaldoCorrigido.Focused or RBPonteiro.Focused or RBCheque.Focused or RBParticipante.Focused ) Then Begin
//    Atualiza_tela;
    JDBaixa.DateText:=JDEntrada.DateText;
    XNEAberto.SetFocus;
  end;
end;

procedure TFrmRecebBaixa_subst.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Elote.clear;
  dm_tabelas.ZQLoteamento.close;
  JDEntrada.Enabled := False;
  JDBaixa.Enabled   := False; 
  ZQRecBai.Close;
  ZQRecebBxTemp.Close;
  ZQrecebimento.close;
{  ZQrecebimento.SQL.Clear;
  ZQrecebimento.SQL.Add('update recebimento set saldo=0 where saldo<=5');
  ZQrecebimento.ExecSQL;
  ZQrecebimento.Close;}
  ZQRecebBxTemp.close;
  dm_tabelas.ZQLoteamento.close;
  DM_tabelas.ZQCobaRe.close;
  DM_Tabelas.ZQNumOrdem.close;
  DM_tabelas.ZQReBxHi.close;
  DM_Tabelas.ZQDividido.close;
  DM_tabelas.contador_bx.close;
  DM_Tabelas.ZQMovBancaria.close;
  DM_tabelas.ZQCheque.close;
  DM_tabelas.ZQCaixa.close;
  FrmRecebBaixa_subst:=nil;
  action:=CaFree;
end;


procedure TFrmRecebBaixa_subst.ZQRecBaiCalcFields(DataSet: TDataSet);
var
  vardI, vardT : Tdate;
  varsal : Double;
begin
    if (ZQRecBai.FieldByName('saldo').AsFloat>0) and (ZQRecBai.FieldByName('recpag').AsString='R') and (CBSaldoCorrigido.Checked)  Then Begin
    varsal := ZQRecBai.FieldByName('Valor').AsFloat;

    ZQRecebBxTemp.Close;
    ZQRecebBxTemp.SQL.Clear;
    ZQRecebBxTemp.SQL.Add('SELECT H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,H.sq,H.valor_parcela,');
    ZQRecebBxTemp.SQL.Add('       R.documento,R.cliente,R.dt_entrada,R.dt_vencimento,R.Valor,R.VrDoc,R.contabil,R.usuario,R.ordem,R.Tipdoc,R.nomeadversa,R.venda_idvenda,R.quadralote,R.idloteamento,R.saldo,R.refvinda,R.Observ,R.recpag,');
    ZQRecebBxTemp.SQL.Add('       R.juros,R.descontos,R.Data_Quitacao,');
    ZQRecebBxTemp.SQL.Add('       R.marcar,R.adversa,R.numordem,R.substituicao,R.sq,R.Parcelas_fixas,R.origem,R.adversa,R.recpag,R.numordem,R.idrecebimento,R.empresa,R.custodaparcela,R.origem,R.adversa,R.recpag,R.numordem,R.RefBaixa,');
    ZQRecebBxTemp.SQL.Add('       B.idreceb_baixa,B.Dt_rec,B.Vr_rec,B.TipDoc,B.Docum,B.Juros_Vr,B.Desc_Vr,B.RefBaixa,B.obsebx,B.Codcontabil,B.dataref,B.vencimento,B.substituicao,B.vr_ab,B.sq,B.valor_parcela ');
    ZQRecebBxTemp.SQL.Add(' FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib join receb_baixa as B ON B.refbaixa=H.refer');
    ZQRecebBxTemp.SQL.Add(' where H.idrecib='+quotedstr(ZQRecBai.FieldByName('idrecebimento').Text));
    ZQRecebBxTemp.open;
    ZQRecebBxTemp.First;
    ZQRecebBxTemp.First;
    vardI := ZQRecBai.FieldByName('Dt_Vencimento').AsDateTime;
    vardT := strtodate(FrmRecebBaixa_subst.JDEntrada.DateText);
    ZQRecebBxTemp.DisableControls;
    while (not ZQRecebBxTemp.Eof) and (vardT>=ZQRecebBxTemp.FieldByName('dataref').AsDateTime) do begin
      varsal := Calcjuros(vardI,ZQRecebBxTemp.FieldByName('dataref').AsDateTime,varsal,XNJuros.Value);
      varsal := varsal-(ZQRecebBxTemp.FieldByName('Vr_rec').AsFloat*ZQRecebBxTemp.FieldByName('percent_usado').AsFloat/100);
      vardI := ZQRecebBxTemp.FieldByName('dataref').AsDateTime;
      ZQRecebBxTemp.Next;
    end;
    ZQRecebBxTemp.EnableControls;
    ZQRecBai.FieldByName('saldocorrig').AsFloat := Calcjuros(vardI,vardT,varsal,XNJuros.Value);
  end
  else if ZQRecBai.FieldByName('saldo').AsFloat=0 Then
    ZQRecBai.FieldByName('saldocorrig').AsFloat := 0
  else
    ZQRecBai.FieldByName('saldocorrig').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;
end;

procedure TFrmRecebBaixa_subst.CDSParcelasCalcFields(DataSet: TDataSet);
begin
  CDSParcelasPercent.Value := CDSParcelasVrParc.Value*100/XNERecebido.Value;
//  CDSParcelasjuros.Value := XNEJuros.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmRecebBaixa_subst.ZQRatearCalcFields(DataSet: TDataSet);
begin
    ZQRatear.FieldByName('vr_base').AsFloat := ZQRatear.FieldByName('vr_rec').AsFloat * ZQRatear.FieldByName('percent_usado').AsFloat / 100;
    ZQRatear.FieldByName('vr_divi').AsFloat := ZQRatear.FieldByName('vr_base').AsFloat * ZQRatear.FieldByName('percentual').AsFloat /100;

end;

procedure TFrmRecebBaixa_subst.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').AsString = 'R' Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrmRecebBaixa_subst.CDS_divididoCalcFields(DataSet: TDataSet);
begin
  CDS_divididodescrcontabil2.Value := CDS_divididocontabil.Text+' - '+CDS_divididodescrcontabil.Value;
end;

procedure TFrmRecebBaixa_subst.CBPeriodoEntExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma : double;
  varvenci : Tdatetime;
  vartipo : string;
begin
  if CBPeriodoEnt.ItemIndex <0 Then Begin
    CBPeriodoEnt.SetFocus;
    exit;
  end;
  varvenci :=  strtodate(JDEntrada.DateText);
  ZQRecBai.Filtered := false;  
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;

  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  CDSParcelas.IndexFieldNames := 'ordem';
  varfinal := strtoint(floattostr(XNEParcelas.Value));
  if StrIsNumero(CBPeriodoEnt.Text) Then Begin
    vardias := strtoint(CBPeriodoEnt.Text);
    vartipo := 'N';
  end
  else Begin
    if CBPeriodoEnt.ItemIndex<0 Then Begin
      CBPeriodoEnt.SetFocus;
      Exit;
    end;
    vartipo := 'S';
    case CBPeriodoEnt.ItemIndex of
      0 : vardias := 0;
      1 : vardias := 1;
      2 : vardias := 7;
      3 : vardias := 15;
      4 : vardias := 1;
      5 : vardias := 2;
      6 : vardias := 3;
      7 : vardias := 4;
      8 : vardias := 6;
      9 : vardias := 12;
      10: vardias := 24;
    end;
    if CBPeriodoEnt.ItemIndex<=2 Then
      vartipo := 'N';
  end;
  DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger,[]);
  varsoma := 0;
  for varvezes := 1 to varfinal do Begin
    CDSParcelas.Insert;
    CDSParcelasOrdem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
    CDSParcelasjuros.Value :=  ExRound(XNEJuros.Value,2)/ExRound(varfinal,2);
    CDSParcelasVrParc.Value := ExRound(XNERecebido.Value,2)/ExRound(varfinal,2);
    if ZQRecBai.FieldByName('recpag').AsString='R' Then
      CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').AsInteger
    else
      CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').AsInteger;
      CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
    if varvezes = 1 Then
       CDSParcelasVenci.Value := strtodate(JDEntrada.DateText)
    else Begin
    if vartipo = 'S' then
        CDSParcelasVenci.Value := IncMonth(strtodate(JDEntrada.DateText),(varvezes*vardias)-(vardias))
      else
        CDSParcelasVenci.Value := varvenci+vardias;
      varvenci := CDSParcelasVenci.Value;
    end;

    CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').AsString;
    CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').AsInteger;
    CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').AsString;
    CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
    CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').AsString;
    CDSParcelassq.Value := ZQRecBai.FieldByName('sq').AsLargeInt;
    CDSParcelasorigem.Value:=ZQRecBai.FieldByName('origem').AsString;
    CDSParcelasParcelas_fixas.Value:=ZQRecBai.FieldByName('Parcelas_fixas').AsString;
    CDSParcelas.Post;
    varsoma := varsoma + CDSParcelasVrParc.Value;
  end;
  CDSParcelas.First;
  CDSParcelas.Edit;
  CDSParcelasVrParc.Value := CDSParcelasVrParc.Value+((XNERecebido.Value)-varsoma);
  CDSParcelas.Post;
  ZQRecBai.Filtered := false;
end;

procedure TFrmRecebBaixa_subst.EadversaExit(Sender: TObject);
var
posi:integer;
begin
  if not DXBFechar.Focused Then
  Begin
    if not empty(Eadversa.Text) Then
    Begin
      if AchaCLiente(400, 300, Eadversa.Text) Then
      Begin
        Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        EcodAdversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        EcodAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
        if not empty(EDLoteamento.Text) then
        begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          DM_tabelas.ZQRecebimento.Close;
          DM_tabelas.ZQRecebimento.SQL.Clear;
          DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
          if pfim.Text='  /  /    ' then
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+' and adversa='+quotedstr(EcodAdversa.Text)+' and saldo>0');
            DM_tabelas.ZQRecebimento.open;
          end
          else
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+' and adversa='+quotedstr(EcodAdversa.Text)+' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2 and saldo>0');
            DM_tabelas.ZQRecebimento.ParamByName('dt1').Value:=strtodate(pini.text);
            DM_tabelas.ZQRecebimento.ParamByName('dt2').Value:=strtodate(pfim.text);
            DM_tabelas.ZQRecebimento.open;
          end;
        end
        else
        begin
          DM_tabelas.ZQRecebimento.Close;
          DM_tabelas.ZQRecebimento.SQL.Clear;
          DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
          if pfim.Text='  /  /    ' then
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and adversa='+quotedstr(EcodAdversa.Text)+' and saldo>0');
            DM_tabelas.ZQRecebimento.open;
          end
          else
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and adversa='+quotedstr(EcodAdversa.Text)+' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2 and saldo>0');
            DM_tabelas.ZQRecebimento.ParamByName('dt1').Value:=strtodate(pini.text);
            DM_tabelas.ZQRecebimento.ParamByName('dt2').Value:=strtodate(pfim.text);
            DM_tabelas.ZQRecebimento.open;
          end;
        end;
        DM_tabelas.ZQRecebimento.Recordcount;
//        JDEntrada.Date := Date;
        JDBaixa.DateText   := JDEntrada.DateText;
        DBText1.Caption := '';
        XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;

        DM_tabelas.ZQCobaRe.Insert;
        DTPIni.Datetext := datetostr(date);
        DTPFim.Datetext := datetostr(date);
        DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
        DM_tabelas.ZQCobaRe.post;

        GBCheque.Visible := False;
        DXBBaixaGravar.Enabled := false;
        EMObsbx.Clear;
        EMNovostit.Clear;
        EContabil.Text := '';
        XNEParcelas.Value := 1;
        XNEAberto.Value := 0;
        XNEVencido.Value := 0;
        XNEMarcado.Value := 0;
        XNERecebido.Value := 0;
        XNEDesconto.Value := 0;
        XNEJuros.Value := 0;
        DBCBTipDoc.Clear;
        DM_tabelas.ZQTipodoc.First;
        DM_tabelas.ZQTipodoc.DisableControls;
        while not DM_tabelas.ZQTipodoc.Eof do
        Begin
          DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
          DM_tabelas.ZQTipodoc.Next;
        end;
        DM_tabelas.ZQTipodoc.EnableControls;
        Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
        Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
        Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
        EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
        EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
        EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
        XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;
        varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;
      end
      else
        JDEntrada.SetFocus;
      filtrabaixa;  
    end
    else
      EcodAdversa.Text := '';

  end;
end;
procedure TFrmRecebBaixa_subst.EAdversaNovaExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not empty(EAdversaNova.Text) Then Begin
      if AchaCLiente(400, 300, EAdversaNova.Text) Then Begin
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        EcodAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
      end
      else
        EAdversaNova.SetFocus;
    end
    else
      EAdversaNova.Text := '';
  end;
end;

procedure TFrmRecebBaixa_subst.CBTipobaixaExit(Sender: TObject);
begin
  JDEntrada.SetFocus;
end;

procedure TFrmRecebBaixa_subst.DXBPesquisarClick(Sender: TObject);
begin
  FrmPesqRecebimento_bx.Top := FrmRecebBaixa_subst.Top+100;
  FrmPesqRecebimento_bx.Left := FrmRecebBaixa_subst.Left;
  FrmPesqRecebimento_bx.showmodal;


  JDEntrada.DateText := datetostr(Date);
  JDBaixa.DateText   := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.DateText := DATETOSTR(date);
  DTPFim.DateText := DATETOSTR(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text   := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


  CBPeriodoEnt.ItemIndex := 0;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True; 
  Atualiza_tela;


  DBGBaixando.SetFocus;
end;

procedure TFrmRecebBaixa_subst.EloteExit(Sender: TObject);
var
  vartipo : string;

begin
 if not empty(Elote.Text) then
 begin
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and quadralote='+quotedstr(Elote.text)+' and saldo>0');
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.Recordcount;
//  JDEntrada.DateTextText := datetostr(Date);
  JDBaixa.DateText   := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.DateText := datetostr(date);
  DTPFim.DateText := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
  DM_tabelas.ZQCobaRe.post;

  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;

  Atualiza_tela;


  DBGBaixando.SetFocus;
 end;

end;

procedure TFrmRecebBaixa_subst.btndesClick(Sender: TObject);
var
  vartipo : string;
begin

  vartipo := '';
  DM_TAbelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_Tabelas.ZQTipodoc.Eof do begin
     if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
          vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
     DM_Tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  vartipo := copy(vartipo,1,length(vartipo)-1);
  if empty(Elote.text) then
     Elote.text:=ZQRecBai.FieldByName('quadralote').AsString;

  ZQRecBai.Close;
  ZQRecBai.SQL.Clear;
  ZQRecBai.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end  ');
//  ZQRecBai.SQL.Add(' where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
  ZQRecBai.SQL.Add(' where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
  if not empty(EcodAdversa.text) Then
  Begin
    ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
  end;

  if pfim.Text<>'  /  /    ' then
  begin
    ZQRecBai.SQL.Add('  and (dt_vencimento between :dt1 and :dt2)');
    ZQRecBai.ParamByName('dt1').AsString:=FormatDateTime('YYYY-MM-DD',strtodate(PIni.text));
    ZQRecBai.ParamByName('dt2').AsString:=FormatDateTime('YYYY-MM-DD',strtodate(PFim.text));
  end;
//  showmessage(zqrecbai.sql.text);
  ZQRecBai.ExecSQL;

  ZQRecBai.Close;
  ZQRecBai.SQL.Clear;
//  ZQRecBai.SQL.Add('select * from recebimento where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
  ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,observ_estorno,tip ');
  ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
  if not empty(EcodAdversa.text) Then
  Begin
    ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
  end;
  if pfim.Text<>'  /  /    ' then
  begin
    ZQRecBai.SQL.Add('  and (dt_vencimento between :dt1 and :dt2)');
    ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
    ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
  end;
  ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
  ZQRecBai.Open;
  ZQRecBai.first;


{  ZQRecBai.First;
  while not ZQRecBai.Eof do
  begin
    application.ProcessMessages;
    ZQRecBai.Edit;
    if  ZQRecBaimarcar.Value = '0' Then
      ZQRecBaimarcar.Value := '1'
    else
      ZQRecBaimarcar.Value := '0';
    ZQRecBai.post;
    ZQRecBai.Next;
  end;}
  Atualiza_tela;
end;

procedure TFrmRecebBaixa_subst.EDLoteamentoDropDown(Sender: TObject);
var
abri:string;
begin
  abri:='F';
  EDLoteamento.Clear;
  if DM_tabelas.ZQLoteamento.Active=false then
  begin
    DM_tabelas.ZQLoteamento.open;
    abri:='T';
  end;
  DM_tabelas.ZQLoteamento.First;
  DM_tabelas.ZQLoteamento.DisableControls;
  while not  DM_tabelas.ZQLoteamento.Eof do
  begin
    EDLoteamento.Items.Add( DM_tabelas.ZQLoteamento.FieldByName('idloteamento').AsString+'-'+DM_tabelas.ZQLoteamento.FieldByName('apelido').AsString);
    DM_tabelas.ZQLoteamento.Next;
  end;
  DM_tabelas.ZQLoteamento.EnableControls;
  if abri='T' then
     DM_tabelas.ZQLoteamento.Close;
  abri:='F';
end;

procedure TFrmRecebBaixa_subst.DBEDonoExit(Sender: TObject);
begin
 DBEBaixaDocum.setfocus;
end;

procedure TFrmRecebBaixa_subst.DBGBaixandoExit(Sender: TObject);
begin
  Atualiza_tela;
end;

procedure TFrmRecebBaixa_subst.Label9MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  //mover formulario sem bordas
  MovimentaObject(sender,button,shift,x,y,FrmRecebBaixa_subst);
end;


procedure TFrmRecebBaixa_subst.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQRecBai', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaisaldocorrig', 'saldocorrig', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaidocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaicliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaisaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaimarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBairefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaicontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaicustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBairecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBainumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBainomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaivenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBainumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBainomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaisomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaidescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecBai', 'ZQRecBaiData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTemprefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTemppercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTemprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTemporigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTemprecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebBxTemp', 'ZQRecebBxTempsubstituicao_1', 'substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQRatear', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearvr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRateartipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearvr_base', 'vr_base', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearvr_divi', 'vr_divi', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRatear', 'ZQRatearordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQItensRecibo', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciborefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciborefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciborecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensRecibonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQItensRecibo', 'ZQItensReciboSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQDiv_usado', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQDiv_usado', 'ZQDiv_usadoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQDiv_usado', 'ZQDiv_usadoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQDiv_usado', 'ZQDiv_usadopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQDiv_usado', 'ZQDiv_usadoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQDiv_usado', 'ZQDiv_usadotipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQDiv_usado', 'ZQDiv_usadoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQrecebimento', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQrecebimento', 'ZQrecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQbaixa', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQbaixa', 'ZQbaixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQBaixa2', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2idreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Dt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Vr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Docum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Juros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Desc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2obsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Codcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2vencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2Substituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2sq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2VR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQBaixa2', 'ZQBaixa2valor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa_subst, 'ZQRecebe', False);
  RegisterRuntimeField(TFrmRecebBaixa_subst, 'ZQRecebe', 'ZQRecebeID', 'ID', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.


