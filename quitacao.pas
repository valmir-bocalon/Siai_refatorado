

unit quitacao;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, wwdbdatetimepicker, StdCtrls, Mask, DBCtrls,
  XBanner, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, ImgList, XNum, DBClient, XDBDate, XDBNum, ComCtrls,
  JvExComCtrls, JvDateTimePicker, XEdit, XDBEdit, ExtCtrls, JvExControls,
  JvSimIndicator, TFlatGaugeUnit, XDate, System.ImageList;
type
  TFrmquitacao = class(TForm)
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
    XNEAberto: TXNumEdit;
    XNEVencido: TXNumEdit;
    XNEMarcado: TXNumEdit;
    Label4: TLabel;
    DS_Parcelas: TDataSource;
    CDSParcelas: TClientDataSet;
    CDSParcelasTipDoc: TWideStringField;
    CDSParcelasDocum: TWideStringField;
    CDSParcelasOrdem: TWideStringField;
    CDSParcelasVenci: TDateField;
    CDSParcelasVrParc: TFloatField;
    CDSParcelascliente: TIntegerField;
    CDSParcelasusuario: TIntegerField;
    CDSParcelasentrada: TDateField;
    CDSParcelasbanco: TWideStringField;
    CDSParcelasagencia: TWideStringField;
    CDSParcelasdono: TWideStringField;
    CDSParcelascpf_cnpj: TWideStringField;
    DBGrid2: TDBGrid;
    XNEParcelas: TXNumEdit;
    Label5: TLabel;
    DBCBTIPDOC: TDBComboBox;
    Label6: TLabel;
    Label7: TLabel;
    XNEDesconto: TXNumEdit;
    XNEJuros: TXNumEdit;
    XNERecebido: TXNumEdit;
    Label8: TLabel;
    Label15: TLabel;
    Label9: TLabel;
    Label17: TLabel;
    Label19: TLabel;
    XNEResto: TXNumEdit;
    EMObsbx: TMemo;
    Label20: TLabel;
    EMNovostit: TMemo;
    Label21: TLabel;
    CDSParcelascodcontabil: TIntegerField;
    Label22: TLabel;
    EContabil: TEdit;
    CDSParcelasdesccontabil: TWideStringField;
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
    CDSParcelaspercent: TFloatField;
    XDBNumEdit1: TXDBNumEdit;
    DBText1: TDBText;
    CDSParcelasconta: TWideStringField;
    Label28: TLabel;
    DEBConta: TDBEdit;
    GBContra: TGroupBox;
    Enomes: TEdit;
    Ecodnome: TEdit;
    RBParticipante: TRadioButton;
    RBCheque: TRadioButton;
    Label1: TLabel;
    Label23: TLabel;
    CDSParcelascodbanco: TIntegerField;
    GBBanco: TGroupBox;
    DBGBanco: TDBGrid;
    Label29: TLabel;
    EPlanocontabanco: TEdit;
    CDSParcelascontabanco: TIntegerField;
    CDSParcelassomapag: TWideStringField;
    RBPonteiro: TRadioButton;
    DBEidreg: TDBEdit;
    LREg: TLabel;
    CDSParcelasjuros: TFloatField;
    CDSParcelasdesconto: TFloatField;
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




































































































    CBPeriodoEnt: TComboBox;
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
    XEEmpresend: TXEdit;





    CDSParcelasnomedoempreendimento: TWideStringField;



    CDSParcelasvenda_idvenda: TIntegerField;
    CDSParcelasquadralote: TWideStringField;
    Label18: TLabel;
    DXBPesquisar: TdxButtonArround;




    CDSParcelasSubstituicao: TWideStringField;

    CDSParcelassq: TIntegerField;
    DS_Receb_baixa: TDataSource;
    ZQReceb_Baixa: TZQuery;















    Label32: TLabel;

    CBTipoBaixa: TComboBox;
    Datarecebimento_historico_excluidos: TDataSource;
    recebimento_historico_excluidos: TZQuery;






























    btndes: TButton;
    Label34: TLabel;
    zqquitacao: TZQuery;
    DS_zqquitacao: TDataSource;










    barravertical: TFlatGauge;
    Label35: TLabel;
    Elote: TEdit;

    ZQVenda: TZQuery;
    DS_Venda: TDataSource;
    ZQEmpree: TZQuery;
    DS_Empree: TDataSource;
    DS_Entrada: TDataSource;
    ZQEntrada: TZQuery;






























    ZQParcela: TZQuery;




























    DS_Parcela: TDataSource;






























    CDSParcelasnumero_cheque: TWideStringField;
    Label36: TLabel;
    DBEdit1: TDBEdit;
    xvequta: TXNumEdit;
    Label37: TLabel;
    pini: TMaskEdit;
    Label38: TLabel;
    pfim: TMaskEdit;
    JDEntrada: TXDateEdit;
    WDTBaixaEntra: TXDBDateEdit;
    JDBaixa: TXDateEdit;
    vecto1: TXDateEdit;
    DTPIni: TXDateEdit;
    DTPFim: TXDateEdit;














    EAdversaNova: TEdit;
    EcodAdversaNova: TEdit;

    CDSParcelasorigem: TWideStringField;
    CDSParcelasParcelas_fixas: TWideStringField;
    DS_Recebimento: TDataSource;
    ZQRecebe: TZQuery;

    ProgressBar1: TProgressBar;
    Panel1: TPanel;

    qlotes: TMaskEdit;
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
    procedure btndesClick(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure XNERecebidoEnter(Sender: TObject);
    procedure DBGBaixandoExit(Sender: TObject);
    procedure Label9MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }

    function ObterIdLoteamentoDaVenda(const AVendaId: Int64;
      const AApelido: string): Int64;
    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frmquitacao: TFrmquitacao;
  ql,VArDoc, varrecpag : String;

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

function TFrmquitacao.ObterIdLoteamentoDaVenda(const AVendaId: Int64;
  const AApelido: string): Int64;
var
  LConsulta: TZQuery;
begin
  Result := 0;
  if DM_Tabelas = nil then
    Exit;

  LConsulta := TZQuery.Create(nil);
  try
    LConsulta.Connection := DM_Tabelas.zconeccao;

    { The sale/imovel relationship is the canonical source for the lot. }
    if AVendaId > 0 then
    begin
      LConsulta.SQL.Text :=
        'select i.loteamento_idloteamento from venda v ' +
        'inner join imovel i on i.idimovel=v.imovel ' +
        'where v.idvenda=:idvenda';
      LConsulta.ParamByName('idvenda').AsLargeInt := AVendaId;
      LConsulta.Open;
      if not LConsulta.IsEmpty then
        Result := LConsulta.FieldByName('loteamento_idloteamento').AsLargeInt;
      LConsulta.Close;
    end;

    { Keep compatibility with old titles without a sale, but never use an
      unopened or filtered global lookup dataset as the source of the ID. }
    if (Result = 0) and (Trim(AApelido) <> '') then
    begin
      LConsulta.SQL.Text :=
        'select idloteamento from loteamento where apelido=:apelido ' +
        'order by idloteamento limit 1';
      LConsulta.ParamByName('apelido').AsString := Trim(AApelido);
      LConsulta.Open;
      if not LConsulta.IsEmpty then
        Result := LConsulta.FieldByName('idloteamento').AsLargeInt;
    end;
  finally
    LConsulta.Free;
  end;
end;


procedure TFrmquitacao.Atualiza_tela;
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

procedure TFrmquitacao.DXBFecharClick(Sender: TObject);
begin
  xvequta.Value:=0;
  Elote.Clear;
  Close;
end;

procedure TFrmquitacao.FormShow(Sender: TObject);
begin
  dm_tabelas.ZQLoteamento.Open;
  DM_Tabelas.ZQNumOrdem.open;
  DM_tabelas.ZQCobaRe.open;
  DM_tabelas.ZQCaixa.open;
  DM_Tabelas.ZQMovBancaria.open;

  DM_Tabelas.ZQContaBancaria.close;
  DM_Tabelas.ZQContaBancaria.SQL.Clear;
  DM_Tabelas.ZQContaBancaria.SQL.Add('Select  idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,nometitular,aberta,senha_interna,senha_conta,ativa,');
  DM_Tabelas.ZQContaBancaria.SQL.Add('        jurosemboleto,taxadiaria,valoroupercent,carteira,boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,Apelido,cod_transmissao,');
  DM_Tabelas.ZQContaBancaria.SQL.Add('        complemento,moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote ');
  DM_Tabelas.ZQContaBancaria.SQL.Add(' from conta_bancaria');
  DM_Tabelas.ZQContaBancaria.Open;

  DM_tabelas.ZQTipodoc.close;
  DM_tabelas.ZQTipodoc.SQL.Clear;
  DM_tabelas.ZQTipodoc.sql.Add('select idtipodocumento,tipodoc,descricao,vend_receb,vend_caixa,receb_receb,receb_caixa,dados_chequ,so_avista,lancabanco,somapaga,lancinclus from tipodocumento');
  DM_tabelas.ZQTipodoc.open;
  DM_tabelas.ZQTipodoc.First;

  DM_tabelas.ZQPlanoDeContas.close;
  DM_tabelas.ZQPlanoDeContas.SQL.clear;
  DM_tabelas.ZQPlanoDeContas.SQL.Add('Select codigo,doccomum,mascara,descricao,classificacao,valor,cod_reduzido,cred_debi,usoativo,CDPlano,DespFixVar from Plano_contas');
  DM_tabelas.ZQPlanoDeContas.open;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and 1=0');
  DM_tabelas.ZQRecebimento.open;
//  if  DM_tabelas.ZQRecebimento.Active=false then
//       DM_tabelas.ZQRecebimento.Open;



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


//  ZQRecBai.Close;

  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  JDEntrada.Datetext := datetostr(Date);
  JDBaixa.Datetext := datetostr(Date);
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;

//  if empty(DM_tabelas.ZQRecebimentocliente.Text) Then Begin
//    Close;
//    Exit;
//  end;

  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
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
  DM_Tabelas.ZQTipoDoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_Tabelas.ZQTipoDoc.EnableControls;
{  Enomes.Text   := DM_tabelas.ZQRecebimentonomecli.Value;
  Ecodnome.Text := DM_tabelas.ZQRecebimentocliente.Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimentoadversanome.Value;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimentoadversanome.Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimentoadversa.Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimentoadversa.Text;
  varrecpag := DM_Tabelas.ZQRecebimentoRecPag.Value;}


  CBPeriodoEnt.ItemIndex := 4;
  //filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;
//  JDEntrada.SetFocus;
  Elote.SetFocus;
end;

procedure TFrmquitacao.DBGBaixandoDrawColumnCell(Sender: TObject;
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

procedure TFrmquitacao.DBGBaixandoColEnter(Sender: TObject);
begin
  if DBGBaixando.SelectedField = ZQRecBai.FieldByName('marcar') then
    DBGBaixando.Options := DBGBaixando.Options - [dgEditing]
  else
    DBGBaixando.Options := DBGBaixando.Options + [dgEditing];
end;

procedure TFrmquitacao.DBGBaixandoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DBGBaixando.SelectedField.fieldname = 'marcar' Then Begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').AsString = '0' Then
      ZQRecBai.FieldByName('marcar').AsString := '1'
    else
      ZQRecBai.FieldByName('marcar').AsString := '0';
    ZQRecBai.Post;      
   // Atualiza_tela;
  end;
end;

procedure TFrmquitacao.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    close;
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmquitacao.XNERecebidoExit(Sender: TObject);
begin
   if XNEMarcado.Value >= XNERecebido.Value Then
     XNEResto.Value := XNEMarcado.Value - (XNERecebido.Value + XNEDesconto.Value)
   else
     XNEJuros.Value :=  XNERecebido.Value - XNEMarcado.Value;
end;

procedure TFrmquitacao.DXBBaixaGravarClick(Sender: TObject);
Var
  Varpago, percentbaixa, baixatot, varjuros, varx,
  varsomarepasse, varsomarepasseautomatico : Double;
  posi,posi2,contador,varreg,LRegistroAtual : Integer;
  LIdLoteamento: Int64;
  ctbl,varordem, varnumordem1, varnumordem2 : string;
begin

  if empty(CBTipobaixa.Text) then
  begin
    mensagem('Selecione o tipo de operação !');
    CBTipobaixa.SetFocus;
    exit;
  end;  
  if empty(EMNovostit.Text) then
  begin
    mensagem('Digite o Histórico da Quitação !');
    EMNovostit.SetFocus;
    exit;
  end;

  { Validate every generated title before writing the baixa. This prevents a
    partial quitacao when an old record has neither a sale nor a valid lot. }
  LRegistroAtual := CDSParcelas.RecNo;
  CDSParcelas.DisableControls;
  try
    CDSParcelas.First;
    while not CDSParcelas.Eof do
    begin
      LIdLoteamento := ObterIdLoteamentoDaVenda(
        CDSParcelasvenda_idvenda.AsLargeInt,
        CDSParcelasnomedoempreendimento.AsString);
      if LIdLoteamento = 0 then
      begin
        ShowMessage('Nao foi possivel identificar o loteamento da venda ' +
          CDSParcelasvenda_idvenda.AsString +
          '. A quitacao nao foi gravada.');
        Exit;
      end;
      CDSParcelas.Next;
    end;
  finally
    if LRegistroAtual > 0 then
      CDSParcelas.RecNo := LRegistroAtual;
    CDSParcelas.EnableControls;
  end;

  DXBBaixaGravar.Enabled := False;
  DXBFechar.Enabled := False;

  ZQRecBai.Filtered:=false;  
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;

  ql:=ZQRecBai.FieldByName('quadralote').AsString;
  
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  varnumordem1 := '';
  varnumordem2 := '';
  If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
    DM_tabelas.ZQCobaRe.Edit;
  DM_tabelas.ZQCobaRe.FieldByName('entrada').AsDateTime := strtodate(JDEntrada.Datetext);
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

  BarraVertical.Progress:=0;
  BarraVertical.Visible:=true;
  CDSParcelas.DisableControls;
  while not CDSParcelas.Eof do
  begin
    BarraVertical.Progress:=CDSParcelas.RecNo * 100 div CDSParcelas.RecordCount;
    inc(contador);

    if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP')  or (CDSParcelasTipDoc.Value='TB')  or (CDSParcelasTipDoc.Value='BB') or (CDSParcelasTipDoc.Value='CH')  or (CDSParcelasTipDoc.Value='DI') then
    begin

      if CDSParcelascodContabil.Value=0 then
         ctbl:='0'
      else
         ctbl:=CDSParcelascodContabil.Text;

      DM_tabelas.ZQReceb_Baixa.Close;
      DM_tabelas.ZQReceb_Baixa.SQL.Clear;

      // com juros e descontos
{      DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,Juros_Vr,Desc_Vr,refbaixa,codcontabil,Obsebx,valor_parcela,substituicao,sq) VALUES ('+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value-CDSParcelasdesconto.Value+CDSParcelasjuros.value),'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDBaixa.Datetext)))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                              quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(CDSParcelasDocum.Value)+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasjuros.Value),'.')))+','+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasdesconto.Value),'.')))+','+quotedstr(DM_tabelas.ZQCobaRecotagem.Text)+','+quotedstr(ctbl)+','+quotedstr(EMObsbx.Text)+','+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value),'.')))+',''Q'','+quotedstr(inttostr(contador))+')');}


      //sem juros e sem descontos igual era vide abaixo
      DM_tabelas.ZQReceb_Baixa.SQL.Add('INSERT INTO  Receb_Baixa (Vr_rec,Dt_rec,dataref,vencimento,Tipdoc,Docum,refbaixa,codcontabil,Obsebx,valor_parcela,substituicao,sq) VALUES ('+
//                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value+CDSParcelasjuros.value-CDSParcelasdesconto.Value),'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDBaixa.Datetext)))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDBaixa.Datetext)))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                              quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(CDSParcelasDocum.Value)+','+
                                              quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(ctbl)+','+quotedstr(EMObsbx.Text)+','+
                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+',''Q'','+quotedstr(inttostr(contador))+')');                                              
//                                              quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value),'.')))+',''Q'','+quotedstr(inttostr(contador))+')');
                                              
      DM_tabelas.ZQReceb_Baixa.ExecSQL;

      {  DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQReceb_BaixaVr_rec.Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
        DM_tabelas.ZQReceb_Baixavalor_parcela.Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);
        DM_tabelas.ZQReceb_BaixaDt_rec.Value := strtodate(JDBaixa.Datetext); //DM_tabelas.ZQCobareHoje.Value;
        DM_TAbelas.ZQReceb_Baixadataref.Value := strtodate(JDEntrada.Datetext);
        DM_TAbelas.ZQReceb_Baixavencimento.Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQReceb_BaixaTipdoc.Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQReceb_BaixaDocum.Value  := CDSParcelasDocum.Value;
//        DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
//        DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
        DM_tabelas.ZQReceb_Baixarefbaixa.Value := DM_tabelas.ZQCobaRecotagem.Value;
        DM_tabelas.ZQReceb_Baixacodcontabil.Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQReceb_BaixaObsebx.Value := EMObsbx.Text;
        DM_tabelas.ZQReceb_BaixaDocum.Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQReceb_Baixasubstituicao.Value:='Q';
        DM_tabelas.ZQReceb_Baixasq.Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;}
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
                                        //quotedstr(TrocaVirgPPto(TrimChar(floattostr((CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value),'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                        quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.Datetext)))+','+
                                        quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(CDSParcelasDocum.Value)+','+
                                        quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(ctbl)+','+quotedstr(EMObsbx.Text)+','+
                                        quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+',''Q'','+quotedstr(inttostr(contador))+')');
                                        //quotedstr(TrocaVirgPPto(TrimChar(floattostr((CDSParcelasVrParc.Value+CDSParcelasdesconto.Value)-CDSParcelasjuros.value),'.')))+',''Q'','+quotedstr(inttostr(contador))+')');
      DM_tabelas.ZQReceb_Baixa.ExecSQL;



      {  DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQReceb_BaixaVr_ab.Value := ((CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value);
        DM_tabelas.ZQReceb_Baixavalor_parcela.Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);
        DM_TAbelas.ZQReceb_Baixadataref.Value := strtodate(JDEntrada.Datetext);
        DM_TAbelas.ZQReceb_Baixavencimento.Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQReceb_BaixaTipdoc.Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQReceb_Baixarefbaixa.Value := DM_tabelas.ZQCobaRecotagem.Value;
        DM_tabelas.ZQReceb_Baixacodcontabil.Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQReceb_BaixaObsebx.Value := EMObsbx.Text;
        DM_tabelas.ZQReceb_BaixaDocum.Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQReceb_Baixasubstituicao.Value:='Q';
        DM_tabelas.ZQReceb_Baixasq.Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;
       }
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

    LIdLoteamento := ObterIdLoteamentoDaVenda(
      CDSParcelasvenda_idvenda.AsLargeInt,
      CDSParcelasnomedoempreendimento.AsString);

    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;

    DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (documento,cliente,dt_entrada,dt_vencimento,Valor,VrDoc,contabil,usuario,ordem,Tipdoc,nomeadversa,venda_idvenda,quadralote,idloteamento,saldo,refvinda,Observ,recpag,'+
                                     'adversa,numordem,substituicao,sq,Parcelas_fixas,origem) VALUES ('+
                                     quotedstr(CDSParcelasDocum.Value)+','+quotedstr(Ecodnome.Text)+','+quotedstr(FormatDateTime('YYYY-MM-DD',DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime))+','+quotedstr(FormatDateTime('YYYY-MM-DD',CDSParcelasVenci.Value))+','+
//                                     quotedstr(TrocaVirgPPto(TrimChar(floattostr((CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNERecebido.Value),'.')))+','+
                                     quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(XNERecebido.Value),'.')))+','+
                                     quotedstr(ctbl)+','+quotedstr(inttostr(DM_tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt))+','+quotedstr(varnumordem1+'-'+CDSParcelasOrdem.Value)+','+
                                     quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(EAdversa.text)+','+quotedstr(inttostr(CDSParcelasvenda_idvenda.Value))+','+quotedstr(CDSParcelasquadralote.Value)+','+
                                     quotedstr(IntToStr(LIdLoteamento))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDSParcelasVrParc.Value),'.')))+','+
                                     quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(EMNovostit.Text)+','+quotedstr(ZQRecBai.FieldByName('recpag').AsString)+','+quotedstr(EcodAdversa.Text)+','+quotedstr(varnumordem1)+',''Q'','+quotedstr(inttostr(contador))+','+
                                     quotedstr(CDSParcelasParcelas_fixas.Value)+','+quotedstr(CDSParcelasorigem.Value)+')');
    DM_tabelas.ZQRecebimento.ExecSQL;
    DM_tabelas.ZQRecebimento.Close;

    //pegar o ultimo registro auto incremento criado
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
    DM_tabelas.ZQRecebimentoValor.Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
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

    DM_Tabelas.ZQRecebimentosubstituicao.Value:='Q';
    DM_tabelas.ZQRecebimentosq.Value := contador;
    DM_tabelas.ZQRecebimento.Post; }

    if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString='S' Then
    Begin
        if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
        begin
          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').AsInteger := CDSParcelascodContabil.Value;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').AsDateTime := date;
//          DM_tabelas.ZQCaixavr_lan.Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').AsFloat := (CDSParcelasVrParc.Value);
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
          DM_tabelas.ZQCaixa.FieldByName('Substituicao').AsString := 'Q';
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
        DM_tabelas.ZQCheque.FieldByName('emissao').AsDateTime := strtodate(JDEntrada.Datetext);
        DM_Tabelas.ZQCheque.FieldByName('conta').AsString := CDSParcelasConta.Value;
        DM_Tabelas.ZQCheque.FieldByName('numero').AsString := CDSParcelasnumero_cheque.Value;
        DM_tabelas.ZQCheque.FieldByName('deposito').AsDateTime := CDSParcelasVenci.Value;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := varreg;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').AsInteger := varreg;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').AsString := 'Q';
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
        DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').AsString := 'Q';
        DM_tabelas.ZQMovBancaria.FieldByName('sq').AsLargeInt := contador;
        DM_Tabelas.ZQMovBancaria.Post;

        //      CalcSaldo('','',False);

    end;
    CDSParcelas.Next;
  end;

  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa');
  DM_tabelas.ZQReceb_Baixa.open;

  CDSParcelas.EnableControls;

  // histórico do QUITAÇÃO 06/08/2010
  CDSParcelas.Prior;
  zqQUITACAO.close;
  zqQUITACAO.Open;
  zqQUITACAO.Insert;
  ZQquitacao.FieldByName('idvenda').AsLargeInt:=CDSParcelasvenda_idvenda.Value;
  ZQquitacao.FieldByName('datacessao').AsDateTime:=strtodate(JDEntrada.Datetext);
  ZQquitacao.FieldByName('hora').AsDateTime:=time;
  ZQquitacao.FieldByName('historico').AsString:=EMNovostit.Text;
  zqQUITACAO.post;
  zqQUITACAO.close;
  CDSParcelas.Last;
  // até aqui

  BarraVertical.Progress:=0;
//  BarraVertical.Visible:=false;
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
    ZQDiv_usado.EnableControls;
  end;
  ZQRecebe.close;
  ZQRecBai.First;
  varpago := 0;
  Memo1.Clear;
  baixatot := XNERecebido.Value;
  if DM_tabelas.ZQReBxHi.Active=false then
     DM_tabelas.ZQReBxHi.Open;
//  DM_tabelas.ZQRecebimento.Close;
//  DM_tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
//  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
//  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
//  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
//  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where quadralote='+quotedstr(ZQRecBaiquadralote.Value)+' order by DT_Vencimento ');
//  DM_tabelas.ZQRecebimento.open;

  //EXCLUIR AS PARCELAS ANTIGAS COM CHEQUES E O CAIXA
//  BarraVertical.Maximum:=ZqRecBai.RecordCount;

  BarraVertical.Progress:=0;
  BarraVertical.Visible:=true;
  ZqRecBai.First;
  ZQReceb_Baixa.open;
  recebimento_historico_excluidos.open;
  if DM_tabelas.ZQCheque.Active=false then
     DM_tabelas.ZQCheque.open;
  if DM_tabelas.ZQReBxHi.Active=false then
     DM_tabelas.ZQReBxHi.Open;

  DM_tabelas.ZQCheque.MasterSource:=nil;
  DM_tabelas.ZQCheque.MasterFields:='';
  DM_tabelas.ZQCheque.LinkedFields:='';
  ZqRecBai.DisableControls;
  While not ZqRecBai.Eof do
  Begin
    {BarraVertical.Progress:=ZqRecBai.RecNo * 100 div ZqRecBai.RecordCount;
    DM_tabelas.ZQCheque.First;
    if DM_tabelas.ZQCheque.RecordCount>0 then
    begin
      while not DM_tabelas.ZQCheque.Eof do
      begin
        if DM_tabelas.ZQChequeidrecebimento.Value=ZqRecBaiidrecebimento.Value then
        begin
           DM_tabelas.ZQCheque.Delete;
        end
        else
        begin
           DM_tabelas.ZQCheque.Next;
        end;
      end;
    end;  }

    application.ProcessMessages;
    BarraVertical.progress:=ZqRecBai.RecNo * 100 div ZqRecBai.RecordCount;
    DM_tabelas.ZQCheque.close;
    DM_tabelas.ZQCheque.sql.Clear;
    DM_tabelas.ZQCheque.SQL.Add('delete from cheque_rec where idrecebimento='+quotedstr(ZQRecBai.FieldByName('idrecebimento').Text));
    DM_tabelas.ZQCheque.ExecSQL;

//    DM_tabelas.ZQMovBancariavinculo.Value := 'R'+DM_tabelas.ZQCobaRecotagem.Text;

  //  DM_tabelas.ZQCaixavinculo.Value := 'P'+DM_tabelas.ZQCobaRecotagem.Text;
    ZQReceb_Baixa.Close;
    ZQReceb_Baixa.SQL.Clear;
    ZQReceb_Baixa.SQL.add(' delete from Receb_Baixa  where docum='+quotedstr(alltrim(ZQRecBai.FieldByName('documento').AsString))+' and sq='+quotedstr(alltrim(inttostr(ZQRecBai.FieldByName('sq').AsLargeInt))));
    ZQReceb_Baixa.ExecSQL;


{    ZQReceb_Baixa.Filtered:=false;
    ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(ZqRecBaiDocumento.Value))+' and sq='+quotedstr(alltrim(inttostr(ZqRecBaisq.value)));
    ZQReceb_Baixa.Filtered:=true;
    ZQReceb_Baixa.First;
    if ZQReceb_Baixa.RecordCount>0 then
    begin
      while not ZQReceb_Baixa.Eof do
      begin
        if ZQReceb_Baixavr_ab.Value>0 then
        begin
          ZQReceb_Baixa.Delete;
        end
        else
        begin
          ZQReceb_Baixa.Next;
        end;
      end;
    end;
    ZQReceb_Baixa.Filtered:=false;}

    DM_tabelas.ZQReBxHi.Close;
    DM_tabelas.ZQReBxHi.SQL.clear;
    DM_tabelas.ZQReBxHi.sql.Add('delete  from RecBxHist where  idrecib='+quotedstr(ZQRecBai.FieldByName('idrecebimento').Text));
    DM_tabelas.ZQReBxHi.ExecSQL;


  {  DM_tabelas.ZQReBxHi.First;
    if DM_tabelas.ZQReBxHi.RecordCount>0 then
    begin
      while not DM_tabelas.ZQReBxHi.Eof do
      begin
        if DM_tabelas.ZQReBxHiidrecib.Value=ZqRecBaiidrecebimento.Value then
        begin
           DM_tabelas.ZQReBxHi.Delete;
        end
        else
        begin
           DM_tabelas.ZQReBxHi.Next;
        end;
      end;
    end;}

    recebimento_historico_excluidos.Insert;
    recebimento_historico_excluidos.FieldByName('idrecebimento').AsLargeInt:= ZQRecBai.FieldByName('idrecebimento').AsLargeInt;
    recebimento_historico_excluidos.FieldByName('documento').AsString:= ZQRecBai.FieldByName('documento').AsString;
    recebimento_historico_excluidos.FieldByName('cliente').AsInteger:= ZQRecBai.FieldByName('cliente').AsInteger;
    recebimento_historico_excluidos.FieldByName('usuario').AsInteger:= ZQRecBai.FieldByName('usuario').AsInteger;
    recebimento_historico_excluidos.FieldByName('Dt_Entrada').AsDateTime:= ZQRecBai.FieldByName('Dt_Entrada').AsDateTime;
    recebimento_historico_excluidos.FieldByName('Dt_Vencimento').AsDateTime:= ZQRecBai.FieldByName('Dt_Vencimento').AsDateTime;
    recebimento_historico_excluidos.FieldByName('Valor').AsFloat:= ZQRecBai.FieldByName('Valor').AsFloat;
    recebimento_historico_excluidos.FieldByName('Observ').AsString:= ZQRecBai.FieldByName('Observ').AsString;
    recebimento_historico_excluidos.FieldByName('VrDoc').AsFloat:= ZQRecBai.FieldByName('VrDoc').AsFloat;
    recebimento_historico_excluidos.FieldByName('ordem').AsString:= ZQRecBai.FieldByName('ordem').AsString;
    recebimento_historico_excluidos.FieldByName('TipDoc').AsString:= ZQRecBai.FieldByName('TipDoc').AsString;
    recebimento_historico_excluidos.FieldByName('saldo').AsFloat:= ZQRecBai.FieldByName('saldo').AsFloat;
    recebimento_historico_excluidos.FieldByName('marcar').AsString:= ZQRecBai.FieldByName('marcar').AsString;
    recebimento_historico_excluidos.FieldByName('RefBaixa').AsInteger:= ZQRecBai.FieldByName('RefBaixa').AsInteger;
    recebimento_historico_excluidos.FieldByName('refvinda').AsInteger:= ZQRecBai.FieldByName('refvinda').AsInteger;
    recebimento_historico_excluidos.FieldByName('contabil').AsInteger:= ZQRecBai.FieldByName('contabil').AsInteger;
    recebimento_historico_excluidos.FieldByName('empresa').AsInteger:= ZQRecBai.FieldByName('empresa').AsInteger;
    recebimento_historico_excluidos.FieldByName('custodaparcela').AsFloat:= ZQRecBai.FieldByName('custodaparcela').AsFloat;
    recebimento_historico_excluidos.FieldByName('origem').AsString:= ZQRecBai.FieldByName('origem').AsString;
    recebimento_historico_excluidos.FieldByName('adversa').AsInteger:=ZQRecBai.FieldByName('adversa').AsInteger;
    recebimento_historico_excluidos.FieldByName('recpag').AsString:= ZQRecBai.FieldByName('recpag').AsString;
    recebimento_historico_excluidos.FieldByName('numordem').AsInteger:= ZQRecBai.FieldByName('numordem').AsInteger;
    recebimento_historico_excluidos.FieldByName('idloteamento').AsInteger:= ZQRecBai.FieldByName('idloteamento').AsInteger;
    recebimento_historico_excluidos.FieldByName('venda_idvenda').AsInteger:= ZQRecBai.FieldByName('venda_idvenda').AsInteger;
    recebimento_historico_excluidos.FieldByName('quadralote').AsString:= ZQRecBai.FieldByName('quadralote').AsString;
    recebimento_historico_excluidos.FieldByName('numboleto').AsString:= ZQRecBai.FieldByName('numboleto').AsString;
    recebimento_historico_excluidos.FieldByName('Substituicao').AsString:=ZQRecBai.FieldByName('Substituicao').AsString;
    recebimento_historico_excluidos.FieldByName('sq').AsLargeInt:= ZQRecBai.FieldByName('sq').AsLargeInt;
    recebimento_historico_excluidos.FieldByName('somar').AsString:= ZQRecBai.FieldByName('somar').AsString;
    recebimento_historico_excluidos.FieldByName('motivo').AsString:='Q';
    recebimento_historico_excluidos.Post;
    ZqRecBai.Next;
  end;
  ZqRecBai.EnableControls;
  recebimento_historico_excluidos.close;
  recebimento_historico_excluidos.open;
  recebimento_historico_excluidos.last;
  DM_tabelas.ZQCheque.close;
  DM_tabelas.ZQCheque.sql.Clear;
  DM_tabelas.ZQCheque.SQL.Add('select Banco,Dono,CPF_CNPJ,valor,emissao,deposito,alias,prorrogado,conta,idrecebimento,idpagamento,idcheque,idvinculo,numero,agencia,Substituicao,sq from cheque_rec');
  DM_tabelas.ZQCheque.open;

  ZQReceb_Baixa.Close;
  ZQReceb_Baixa.SQL.Clear;
  ZQReceb_Baixa.SQL.add(' select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa ');
  ZQReceb_Baixa.open;

  DM_tabelas.ZQReBxHi.Close;
  DM_tabelas.ZQReBxHi.SQL.clear;
  DM_tabelas.ZQReBxHi.sql.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist ');
  DM_tabelas.ZQReBxHi.open;


  DM_tabelas.ZQCheque.MasterSource:=DM_tabelas.DS_Recebimento;
  DM_tabelas.ZQCheque.MasterFields:='idrecebimento';
  DM_tabelas.ZQCheque.LinkedFields:='idvinculo';



{  ZqRecBai.First;
  While not ZqRecBai.Eof do
  Begin
    ZqRecBai.Delete;
  end;
 }

  ql:=ZQRecBai.FieldByName('quadralote').AsString;
  ZQRecBai.close;
  ZQRecBai.SQL.Clear;
  ZQRecBai.SQL.Add('delete from recebimento where marcar=''0'' and  saldo>''0'' and quadralote='+quotedstr(ql));
  ZQRecBai.ExecSQL;


  ZQRecBai.close;
  ZQRecBai.SQL.Clear;
  ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao ');
  ZQRecBai.SQL.Add(' from recebimento where marcar=''0'' and  saldo>''0'' and quadralote='+quotedstr(ql));
  ZQRecBai.Open;


  recebimento_historico_excluidos.close;




  BarraVertical.Progress:=0;
//  BarraVertical.Visible:=false;
  CDS_dividido.Close;
  CDS_dividido.CreateDataSet;
  ZQRatear.SQL.Clear;
  ZQRatear.SQL.Add('select idrecebimento, h.valor, percent_usado, b.vr_rec, participante, percentual,planodeconta, dv.tipodocumento, re.ordem, dv.ordem from recbxhist as H');
  ZQRatear.SQL.Add(' join recebimento as re ON H.idrecib=RE.idrecebimento');
  ZQRatear.SQL.Add(' join dividido as dv ON dv.ordem=numordem');
  ZQRatear.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
  ZQRatear.SQL.Add(' join tipodocumento as D ON B.tipdoc=D.tipodoc');
  ZQRatear.SQL.Add(' where somapaga = '+quotedstr('S')+' and refer = '+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text);
  ZQRatear.Open;
  varsomarepasseautomatico := 0;
  ZqRatear.DisableControls;
  while not ZQRatear.Eof do begin
    if CDS_dividido.Locate('participante;contabil;tipodoc',VarArrayOf([ZQRatear.FieldByName('participante').AsInteger,ZQRatear.FieldByName('planodeconta').AsInteger,ZQRatear.FieldByName('tipodocumento').AsString]),[]) then
      CDS_dividido.Edit
    else
      CDS_dividido.Insert;
    CDS_divididoparticipante.Value := ZQRatear.FieldByName('participante').AsInteger;
    CDS_divididovr_base.Value := CDS_divididovr_base.Value + ZQRatear.FieldByName('vr_base').AsFloat;
    CDS_divididovr_calc.Value := CDS_divididovr_calc.Value + ZQRatear.FieldByName('vr_divi').AsFloat;
    CDS_divididocontabil.Value := ZQRatear.FieldByName('planodeconta').AsInteger;
    CDS_divididotipodoc.Value := ZQRatear.FieldByName('tipodocumento').AsString;
    CDS_divididoobs.Value := CDS_divididoobs.Value+ZQRatear.FieldByName('idrecebimento').Text+' | '+ZQRatear.FieldByName('ordem').AsString+' | '+ZQRatear.FieldByName('vr_base').Text+' x '+ZQRatear.FieldByName('percentual').Text+' = '+ZQRatear.FieldByName('vr_divi').Text+chr(13);
    varsomarepasseautomatico := varsomarepasseautomatico + ZQRatear.FieldByName('vr_divi').AsFloat;
    ZQRatear.Next;
  end;
  zqRatear.EnableControls;
  CDS_dividido.First;
  if CDS_dividido.RecordCount>0 Then Begin
    DM_Tabelas.ZQNumOrdem.Append;
    DM_Tabelas.ZQNumOrdem.Post;
  end;
  CDS_dividido.DisableControls;
  while not CDS_dividido.Eof do begin



    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;

    DM_tabelas.ZQRecebimento.SQL.Add('INSERT INTO  Recebimento (documento,cliente,dt_entrada,dt_vencimento,Valor,VrDoc,contabil,ordem,Tipdoc,nomeadversa,saldo,refvinda,Observ,recpag,'+
                                     'adversa) VALUES ('+
                                     quotedstr('Rateio Bx - '+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(CDS_divididoparticipante.Text)+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.datetext)))+','+quotedstr(FormatDateTime('YYYY-MM-DD',strtodate(JDEntrada.datetext)))+','+
                                     quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDS_divididovr_calc.Value),'.')))+','+quotedstr(TrocaVirgPPto(TrimChar(floattostr(CDS_divididovr_base.Value),'.')))+','+
                                     quotedstr(CDS_divididocontabil.Text)+','+quotedstr(DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text+'-'+inttostr(CDS_dividido.RecNo)+'/'+inttostr(CDS_dividido.RecordCount))+','+
                                     quotedstr(CDSParcelasTipDoc.Value)+','+quotedstr(EAdversaNova.text)+','+
                                     quotedstr(TrocaVirgPPto(TrimChar(floattostr(DM_tabelas.ZQRecebimento.FieldByName('Valor').AsFloat),'.')))+','+
                                     quotedstr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text)+','+quotedstr(EMNovostit.Text+chr(13)+CDS_divididoobs.Value)+','+quotedstr('P')+','+
                                     quotedstr(DM_Tabelas.ZQConfiguracoes.FieldByName('paticipantedefault').Text)+')');
    DM_tabelas.ZQRecebimento.ExecSQL;

       // se dar erro, voltar o abaixo e exluir o de cima

//    DM_tabelas.ZQRecebimento.Append;
//    DM_tabelas.ZQRecebimentodocumento.Value := 'Rateio Bx - '+DM_tabelas.ZQCobaRecotagem.Text;
//    DM_tabelas.ZQRecebimentocliente.Value :=  CDS_divididoparticipante.Value;
//    DM_tabelas.ZQRecebimentoDt_Entrada.Value := strtodate(JDEntrada.datetext);
//    DM_tabelas.ZQRecebimentoDt_Vencimento.Value := strtodate(JDEntrada.datetext);
//    DM_tabelas.ZQRecebimentocontabil.Value := CDS_divididocontabil.Value;
//    DM_tabelas.ZQRecebimentoObserv.Value := EMNovostit.Text+chr(13)+CDS_divididoobs.Value;
//    DM_Tabelas.ZQRecebimentorecpag.Value := 'P';
//    DM_tabelas.ZQRecebimentoTipDoc.Value := CDS_divididotipodoc.Value;
//    DM_tabelas.ZQRecebimentoordem.Value := DM_Tabelas.ZQNumOrdemidnumordem.Text+'-'+inttostr(CDS_dividido.RecNo)+'/'+inttostr(CDS_dividido.RecordCount);
//    DM_tabelas.ZQRecebimentorefvinda.Value := DM_tabelas.ZQCobaRecotagem.Value;
//    DM_tabelas.ZQRecebimentoVrDoc.Value :=  CDS_divididovr_base.Value;
//    DM_tabelas.ZQRecebimentoValor.Value :=  CDS_divididovr_calc.Value;
//    DM_tabelas.ZQRecebimentosaldo.Value :=  DM_tabelas.ZQRecebimentoValor.Value;
//    DM_Tabelas.ZQRecebimentoadversa.Value := DM_Tabelas.ZQConfiguracoespaticipantedefault.Value;
//
//    DM_Tabelas.ZQRecebimentonomeadversa.Value := EAdversaNova.text;
//
//    DM_tabelas.ZQRecebimento.Post;
    CDS_dividido.Next;
  end;
  CDS_dividido.EnableControls;
//  varreg := DM_Tabelas.ZQRecebimentoidRecebimento.Value;
  varreg := ZQRecBai.FieldByName('idrecebimento').AsLargeInt;
//  DM_tabelas.ZQRecebimento.Refresh;
{  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;}
//  DM_tabelas.ZQRecebimento.Refresh;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where quadralote='+quotedstr(ql)+' order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;

  DM_Tabelas.ZQRecebimento.Locate('idrecebimento',varreg,[]);

  if FrmImpRecibo=nil then
     FrmImpRecibo:=TFrmImpRecibo.Create(Self);
  FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
  FrmImpRecibo.Nomecli.Text := Enomes.Text;
  FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;

  ///
 // ZQRecBai.Refresh;




  // alteração do valor de venda 28/08/2010
  ZqRecBai.First;

  posi:=pos('-',ql);
  dec(posi);

  posi2:=pos('-',ql);
  inc(posi2);

  ZQEmpree.close;
  ZQEmpree.SQL.Clear;
  ZQEmpree.SQL.Add('Select  loteamento_idloteamento,quadra,lote,lado,area,confrontofrente,confrontofundo,confrontodireita,confrontoesquerda,esquinanome,valorvenal,Obervacaoloteamento,');
  ZQEmpree.SQL.Add('        cornomapa,idimovel,disponivel,medidafrente,medidafundo,medidadedireita,esquinalinha,esquinacurva,medidaesquerda,matri,matricula,proposta,valorCusto ');
  ZQEmpree.SQL.Add(' from imovel where quadra='+quotedstr(copy(ql,1,posi))+' and lote='+quotedstr(copy(ql,posi2,3)));
  ZQEmpree.open;
  ZQVenda.open;
  ZQEntrada.Close;
  ZQEntrada.SQL.Clear;
  ZQEntrada.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQEntrada.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQEntrada.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  ZQEntrada.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  ZQEntrada.SQL.Add(', sum(valor) as entrada from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-E-%''  group by quadralote  order by ordem');
  ZQEntrada.open;

  ZQParcela.Close;
  ZQParcela.SQL.Clear;
  ZQParcela.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQParcela.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQParcela.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  ZQParcela.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  ZQParcela.SQL.Add(', sum(valor) as parcela from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-P-%''  group by quadralote  order by ordem');
  ZQParcela.open;

  ZQVenda.Edit;
  ZQVenda.FieldByName('valorvenda').AsFloat:=ZQEntrada.FieldByName('entrada').AsFloat+ZQParcela.FieldByName('parcela').AsFloat;
  ZQVenda.Post;

  ZQEmpree.close;
  ZQVenda.close;
  ZQEntrada.close;
  ZQParcela.close;
  ql:='';
  // até aqui


  DBGBaixando.Refresh;
  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  XNEJuros.Value:=0;
  XNERecebido.Value:=0;
  XNEDesconto.Value:=0;
  XNEResto.Value:=0;
  XNEParcelas.Value:=1;
  CBPeriodoEnt.Text:='MENSAL';
  EMObsbx.Clear;
  EMNovostit.Clear;
  ql:=elote.Text;
  Elote.Clear;




//  Close;

   // imprime o recibo
  FrmImpRecibo.EValor.Text:=alltrim(transform(XNERecebido.value,'###,###,##0.00'));
  FrmImpRecibo.Nomecli.Text:=Eadversa.Text;
  FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;

  FrmImpRecibo.Adversas.Text := Enomes.Text;
  FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;




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
//  DM_tabelas.ZQRecebimento.Refresh;

  DM_tabelas.ZQReceb_Baixa.Refresh;
  DM_Tabelas.ZQReBxHi.Refresh;
  xvequta.Value:=0;

  JDEntrada.SetFocus;
  DXBFechar.Enabled := true;
  ZQRecBai.Filtered:=false;
  ZQRecBai.close;
  mensagem('Operação Terminada!');
end;

procedure TFrmquitacao.DBCBTIPDOCExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if Empty(DBCBTIPDOC.Text) Then Begin
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipDoc.Text,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').AsString='S' Then Begin
      if (datetostr(CDSParcelasVenci.Value)<>vecto1.Datetext) and ((CDSParcelasTipDoc.Value<>'DP')) Then BEgin
        mensagem('Este tipo de documento só aceita pagamento a vista!!!');
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

procedure TFrmquitacao.DXBBaixaGravarEnter(Sender: TObject);
begin
  CDSParcelas.First;
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
end;

procedure TFrmquitacao.DBCBTIPDOCEnter(Sender: TObject);
Var
  VarTipDoc, VarDocum, VarBanco, Varconta, VarAgencia, VarCPF_CNPJ, VarDono : string;
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
  end
  else
  Begin
    CDSParcelas.Edit;
    CDSParcelasDono.Value := Enomes.Text;
  End;
end;

procedure TFrmquitacao.DBEBaixaVrRecExit(Sender: TObject);
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
  vardif := XNERecebido.Value - varsoma;
  if vardif>0 then
  begin
    if CDSParcelas.RecNO< XNEParcelas.Value Then Begin
      CDSParcelas.RecNO := varregis+1;
      CDSParcelas.DisableControls;
      while not CDSParcelas.Eof do Begin
        CDSParcelas.Edit;
        CDSParcelasVrParc.Value := truncar(vardif / (XNEParcelas.Value-varregis),2);
        vardesc := vardesc - CDSParcelasdesconto.Value;
        CDSParcelas.Next;
      end;
      CDSParcelas.EnableControls;
    end;
  end;
  CDSParcelas.RecNo := varregis;
  DXBBaixaGravar.Enabled := True;
  DXBBaixaGravar.SetFocus;
end;

procedure TFrmquitacao.EContabilExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if empty(Econtabil.text) Then Begin
      mensagem('A descrição contábil não pode ficar em branco....');
      Econtabil.SetFocus;
      exit;
    end;
    if not AchaPlanoDeConta(450, 30, 'S', EContabil.Text, '') then Begin
      EContabil.SetFocus;
      exit;
    end;
    EContabil.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    if not (CDSParcelas.State in [DsInsert, DsEdit]) Then
      CDSParcelas.Edit;
    CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
    CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
  end;
//    CDSParcelasTipDoc.Value := FrmAchaPlanodeConta.ZQTempPlanoContasdoccomum.Value;
end;

procedure TFrmquitacao.XDBNumEdit1Exit(Sender: TObject);
begin
  CDSParcelasVrParc.Value := XNERecebido.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmquitacao.DBEBaixaDocumExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
{    if (not empty(DBEBaixaDocum.Text)) and (DM_tabelas.ZQRecebimento.Locate('documento',DBEBaixaDocum.Text,[])) Then Begin
      mensagem('Este documento já foi lançado anteriormente... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    End;
    if DM_tabelas.ZQCheque.Locate('CH_Conta',DBEBaixaDocum.Text,[]) Then Begin
      mensagem('Este cheque já foi lançado anteriormente... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (pos(quotedstr(CDSParcelasDocum.Value),VarDoc)>0) Then Begin
      mensagem('Este cheque já foi lançado nesta inclusão... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (empty(CDSParcelasDocum.Value)) Then Begin
      mensagem('o campo Nº do cheque deve ser preenchido...');
      DBEBaixaDocum.SetFocus;
      exit;
    end;}
  end;
end;

procedure TFrmquitacao.EContabilEnter(Sender: TObject);
begin
  if empty(Econtabil.Text) Then Begin
    Econtabil.Text := DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').Text;
  end;
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;
  DBText1.Caption := CDSParcelasOrdem.Value;
end;

procedure TFrmquitacao.DBECpf_CnpjExit(Sender: TObject);
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

procedure TFrmquitacao.EnomesExit(Sender: TObject);
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

procedure TFrmquitacao.Label1Click(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
end;

procedure TFrmquitacao.RBChequeClick(Sender: TObject);
begin
  Label1.Visible := true;
  DTPIni.Visible := true;
  Label23.Visible := true;
  DTPFim.Visible := true;
  Enomes.Visible := False;
end;

procedure TFrmquitacao.GBBancoExit(Sender: TObject);
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

procedure TFrmquitacao.GBChequeExit(Sender: TObject);
begin
  GBCheque.Visible := False;
end;

procedure TFrmquitacao.EPlanocontabancoExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then
  Begin
    if not AchaPlanoDeConta(450, 30, 'S', EPlanocontabanco.Text, '') then
    Begin
      EPlanocontabanco.SetFocus;
      exit;
    end;
    EPlanocontabanco.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    CDSParcelascontabanco.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
  end;
end;

procedure TFrmquitacao.filtrabaixa;
var
  vartipo : string;
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
       ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao ');
       ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+Ecodnome.Text+')');
       if not empty(EcodAdversa.text) Then
       Begin
         ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
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
      ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao ');
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
       ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao ');
       ZQRecBai.SQL.Add(' from recebimento where recpag='+quotedstr(varrecpag)+' and ((saldo>0) and  (not saldo is null)) ');
       if not empty(EcodAdversa.text) Then
       Begin
          ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));

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
      // Evita gerar IN () quando nenhum tipo de documento de cheque estiver configurado.
      if vartipo = '' then
        vartipo := quotedstr('')
      else
        vartipo := copy(vartipo,1,length(vartipo)-1);



      ZQRecBai.Close;
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('update recebimento set marcar= case when (dt_vencimento<= :dt) then ''0'' when (dt_vencimento> :dt) then ''1'' end  ');
      ZQRecBai.SQL.Add(' where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
      if not empty(EcodAdversa.text) Then
      Begin
        ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
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
      ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao ');
      ZQRecBai.SQL.Add(' from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
      if not empty(EcodAdversa.text) Then
      Begin
        ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
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

procedure TFrmquitacao.GBContraExit(Sender: TObject);
begin
  if not RBPonteiro.Checked Then
    Filtrabaixa;
end;

procedure TFrmquitacao.RBParticipanteClick(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
  Enomes.SetFocus;
end;

procedure TFrmquitacao.RBPonteiroClick(Sender: TObject);
Var
  VArreg : integer;
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

procedure TFrmquitacao.XDBNumEdit1Enter(Sender: TObject);
begin
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;

end;

procedure TFrmquitacao.JDEntradaExit(Sender: TObject);
begin
  if not (DXBFechar.Focused or GBContra.Focused or CBSaldoCorrigido.Focused or RBPonteiro.Focused or RBCheque.Focused or RBParticipante.Focused ) Then Begin
   // Atualiza_tela;
    JDBaixa.Datetext:=JDEntrada.Datetext;
    vecto1.Datetext:=JDEntrada.Datetext;
//    DBGBaixando.SetFocus;
  end;
end;

procedure TFrmquitacao.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  xvequta.Value:=0;
  // Detach local grids before closing their datasets.  Closing a bound query
  // while the form is being freed can otherwise dispatch DataChange events to
  // controls that are already in the close sequence.
  DS_RecBai.DataSet := nil;
  DS_RecebBxTemp.DataSet := nil;
  dm_tabelas.ZQLoteamento.close;
  JDEntrada.Enabled := False;
  JDBaixa.Enabled   := False;
  ZQRecBai.Close;
  ZQRecebBxTemp.Close;
  DM_Tabelas.ZQNumOrdem.close;
  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQCobaRe.close;
  DM_Tabelas.ZQContaBancaria.close;
  DM_tabelas.ZQCaixa.close;
  DM_tabelas.ZQCheque.close;
  DM_Tabelas.ZQMovBancaria.close;
  panel1.Visible:=true;

//  DM_tabelas.ZQRecebimento.Close;
//  DM_tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
//  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
//  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
//  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
//  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento order by DT_Vencimento');
//  DM_tabelas.ZQRecebimento.open;


  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where quadralote='+quotedstr(qlotes.Text)+ ' order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.last;
  qlotes.Text:='';
  panel1.Visible:=false;
  Frmquitacao:=nil;
  Action:=CaFree;
end;

procedure TFrmquitacao.ZQRecBaiCalcFields(DataSet: TDataSet);
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
    vardI := ZQRecBai.FieldByName('Dt_Vencimento').AsDateTime;
    vardT := strtodate(FrmQUITACAO.JDEntrada.Datetext);
    while (not ZQRecebBxTemp.Eof) and (vardT>=ZQRecebBxTemp.FieldByName('dataref').AsDateTime) do begin
      varsal := Calcjuros(vardI,ZQRecebBxTemp.FieldByName('dataref').AsDateTime,varsal,XNJuros.Value);
      varsal := varsal-(ZQRecebBxTemp.FieldByName('Vr_rec').AsFloat*ZQRecebBxTemp.FieldByName('percent_usado').AsFloat/100);
      vardI := ZQRecebBxTemp.FieldByName('dataref').AsDateTime;
      ZQRecebBxTemp.Next;
    end;
    ZQRecBai.FieldByName('saldocorrig').AsFloat := Calcjuros(vardI,vardT,varsal,XNJuros.Value);
  end
  else if ZQRecBai.FieldByName('saldo').AsFloat=0 Then
    ZQRecBai.FieldByName('saldocorrig').AsFloat := 0
  else
    ZQRecBai.FieldByName('saldocorrig').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;
end;

procedure TFrmquitacao.CDSParcelasCalcFields(DataSet: TDataSet);
begin
  CDSParcelasPercent.Value := CDSParcelasVrParc.Value*100/XNERecebido.Value;
//  CDSParcelasjuros.Value := XNEJuros.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmquitacao.ZQRatearCalcFields(DataSet: TDataSet);
begin
    ZQRatear.FieldByName('vr_base').AsFloat := ZQRatear.FieldByName('vr_rec').AsFloat * ZQRatear.FieldByName('percent_usado').AsFloat / 100;
    ZQRatear.FieldByName('vr_divi').AsFloat := ZQRatear.FieldByName('vr_base').AsFloat * ZQRatear.FieldByName('percentual').AsFloat /100;

end;

procedure TFrmquitacao.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').AsString = 'R' Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrmquitacao.CDS_divididoCalcFields(DataSet: TDataSet);
begin
  CDS_divididodescrcontabil2.Value := CDS_divididocontabil.Text+' - '+CDS_divididodescrcontabil.Value;
end;

procedure TFrmquitacao.CBPeriodoEntExit(Sender: TObject);
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
  varvenci :=  strtodate(vecto1.Datetext);
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
    CDSParcelasjuros.Value :=  truncar(XNEJuros.Value/varfinal,2);
    CDSParcelasVrParc.Value := truncar(XNERecebido.Value/varfinal,2);
    if ZQRecBai.FieldByName('recpag').AsString='R' Then
      CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').AsInteger
    else
      CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').AsInteger;
      CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
    if CDSParcelascodcontabil.Value=0 then
       CDSParcelascodcontabil.Value:=1;
    if empty(CDSParcelasTipDoc.Value) then
       CDSParcelasTipDoc.Value:='BO';

    if varvezes = 1 Then
       CDSParcelasVenci.Value := strtodate(vecto1.Datetext)
    else Begin
    if vartipo = 'S' then
        CDSParcelasVenci.Value := IncMonth(strtodate(vecto1.Datetext),(varvezes*vardias)-(vardias))
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

procedure TFrmquitacao.EadversaExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not empty(Eadversa.Text) Then Begin
      if AchaCLiente(400, 300, Eadversa.Text) Then Begin
        Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        EcodAdversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
      end
      else
        Eadversa.SetFocus;
    end
    else
      EcodAdversa.Text := '';
    filtrabaixa;
//    Atualiza_tela;
  end;
end;

procedure TFrmquitacao.EAdversaNovaExit(Sender: TObject);
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

procedure TFrmquitacao.CBTipobaixaExit(Sender: TObject);
begin
  JDEntrada.SetFocus;
end;

procedure TFrmquitacao.DXBPesquisarClick(Sender: TObject);
begin
  FrmPesqRecebimento_bx.Top := FrmQUITACAO.Top+100;
  FrmPesqRecebimento_bx.Left := FrmQUITACAO.Left;
  FrmPesqRecebimento_bx.showmodal;


  JDEntrada.Datetext := datetostr(Date);
  JDBaixa.Datetext   := datetostr(Date);
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
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
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text   := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True; 
  Atualiza_tela;


  DBGBaixando.SetFocus;
end;

procedure TFrmquitacao.btndesClick(Sender: TObject);
var
vartipo:string;
begin
  vartipo := '';
  DM_TAbelas.ZQTipodoc.First;
  while not DM_Tabelas.ZQTipodoc.Eof do begin
     if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
          vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
     DM_Tabelas.ZQTipodoc.Next;
  end;
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
    ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
    ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
  end;
  ZQRecBai.ExecSQL;

  ZQRecBai.Close;
  ZQRecBai.SQL.Clear;
//  ZQRecBai.SQL.Add('select * from recebimento where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
  ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao ');
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

procedure TFrmquitacao.EloteExit(Sender: TObject);
var
  vartipo1 : string;

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
  DM_tabelas.ZQRecebimento.RecordCount;

  ZQRecBai.Filtered:=false;
  ql:='';
  JDEntrada.Datetext := datetostr(Date);
  JDBaixa.Datetext   := JDEntrada.Datetext;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
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
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text   := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;
  Atualiza_tela;
  JDEntrada.setfocus;

//  DBGBaixando.SetFocus;
 end;

end;

procedure TFrmquitacao.XNERecebidoEnter(Sender: TObject);
begin
  if xvequta.Value>0 then
     XNERecebido.Value:=xvequta.Value;
end;

procedure TFrmquitacao.DBGBaixandoExit(Sender: TObject);
begin
  Atualiza_tela;
end;

procedure TFrmquitacao.Label9MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  //mover formulario sem bordas
  MovimentaObject(sender,button,shift,x,y,Frmquitacao);
end;


procedure TFrmquitacao.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQRecBai', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisaldocorrig', 'saldocorrig', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaidocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaicliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaimarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBairefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaicontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaicustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBairecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  // nomeadversa is returned by every ZQRecBai query. Keeping it as data
  // prevents the participant lookup from clearing the first visible record.
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainomeadversa', 'nomeadversa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaivenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaidescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQRecebBxTemp', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemprefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemppercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemporigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemprecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempsubstituicao_1', 'substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempjuros_1', 'juros_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQRatear', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearvr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRateartipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearvr_base', 'vr_base', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearvr_divi', 'vr_divi', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQItensRecibo', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciborefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciborefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciborecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibovalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibojuros_1', 'juros_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQDiv_usado', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadotipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'recebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmotivo', 'motivo', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'zqquitacao', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQVenda', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreelote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreelado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreematri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQEntrada', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQParcela', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQRecebe', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebe', 'ZQRecebeID', 'ID', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.


