

unit quitacao;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, wwdbdatetimepicker, StdCtrls, Mask, DBCtrls,
  XBanner, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, ImgList, XNum, DBClient, XDBDate, XDBNum, ToolEdit, ComCtrls,
  JvExComCtrls, JvDateTimePicker, XEdit, XDBEdit, ExtCtrls, JvExControls,
  JvSimIndicator, TFlatGaugeUnit, XDate;
type
  TFrmquitacao = class(TForm)
    Label10: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    DBEBaixaDocum: TDBEdit;
    DBEBaixaVrRec: TDBEdit;
    DXBBaixaGravar: TdxButton;
    DS_RecBai: TDataSource;
    ZQRecBai: TZQuery;
    DBGBaixando: TDBGrid;
    DXBFechar: TdxButton;
    ImageList1: TImageList;
    Label2: TLabel;
    Label3: TLabel;
    XNEAberto: TXNumEdit;
    XNEVencido: TXNumEdit;
    XNEMarcado: TXNumEdit;
    Label4: TLabel;
    DS_Parcelas: TDataSource;
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
    CDSParcelasdesccontabil: TStringField;
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
    CDSParcelasconta: TStringField;
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
    CDSParcelassomapag: TStringField;
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
    CDS_divididotipodoc: TStringField;
    CDS_divididoobs: TMemoField;
    DS_ItensRecibo: TDataSource;
    ZQItensRecibo: TZQuery;
    CDS_divididonomeparti: TStringField;
    CDS_divididodescrcontabil: TStringField;
    CDS_divididodescrcontabil2: TStringField;
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
    EcodAdversaNova: TEdit;
    EAdversaNova: TEdit;
    XEEmpresend: TXEdit;





    CDSParcelasnomedoempreendimento: TStringField;



    CDSParcelasvenda_idvenda: TIntegerField;
    CDSParcelasquadralote: TStringField;
    Label18: TLabel;
    DXBPesquisar: TdxButton;




    CDSParcelasSubstituicao: TStringField;

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






























    CDSParcelasnumero_cheque: TStringField;
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
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frmquitacao: TFrmquitacao;
  VArDoc, varrecpag : String;

implementation

uses tabelas, Cad_Recebimento, funcoes, ImpRecibo, PesqRecebimento_bx, uRuntimeFields;

{$R *.dfm}

procedure TFrmquitacao.Atualiza_tela;
Var
  Varreg, varmarc, varvenc : integer;
Begin

  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  varmarc := 0;
  varvenc := 0;
  if ZQRecBai.Active=true then
  begin
    Varreg := ZQRecBai.RecNo;
    ZQRecBai.First;
    while not ZQRecBai.Eof do begin
      XNEAberto.Value := XNEAberto.Value + ZQRecBai.FieldByName('saldocorrig').Value;
      if ZQRecBai.FieldByName('marcar').Value = '0' Then Begin
        XNEMarcado.Value := XNEMarcado.Value + ZQRecBai.FieldByName('saldocorrig').Value;
        XNEJuros.Value := XNEJuros.Value + (ZQRecBai.FieldByName('saldocorrig').Value-ZQRecBai.FieldByName('saldo').Value);
        inc(varmarc);
      end;
      if ZQRecBai.FieldByName('Dt_Vencimento').Value <=date Then Begin
        XNEVencido.Value := XNEVencido.Value + ZQRecBai.FieldByName('saldocorrig').Value;
        inc(varvenc);
      end;
      ZQRecBai.Next;
    end;
    ZQRecBai.RecNo := Varreg;
  end;
  Label31.Caption := inttostr(varmarc);
  Label33.Caption := inttostr(varvenc);
  XNERecebido.Value := XNEMarcado.Value;
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
  DM_Tabelas.ZQContaBancaria.SQL.Add('Select * from conta_bancaria');
  DM_Tabelas.ZQContaBancaria.Open;

  DM_tabelas.ZQTipodoc.close;
  DM_tabelas.ZQTipodoc.SQL.Clear;
  DM_tabelas.ZQTipodoc.sql.Add('select * from tipodocumento');
  DM_tabelas.ZQTipodoc.open;
  DM_tabelas.ZQTipodoc.First;

  DM_tabelas.ZQPlanoDeContas.close;
  DM_tabelas.ZQPlanoDeContas.SQL.clear;
  DM_tabelas.ZQPlanoDeContas.SQL.Add('Select * from Plano_contas');
  DM_tabelas.ZQPlanoDeContas.open;


{  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.Open;}
  if  DM_tabelas.ZQRecebimento.Active=false then
       DM_tabelas.ZQRecebimento.Open;
  

  if ZQRecebBxTemp.Active=false then
  begin
    ZQRecebBxTemp.Close;
    ZQRecebBxTemp.SQL.Clear;
    ZQRecebBxTemp.SQL.Add('SELECT * FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib join receb_baixa as B ON B.refbaixa=H.refer');
    ZQRecebBxTemp.open;
  end;

  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReBxHi.SQL.Clear;
  DM_tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist');
  DM_tabelas.ZQReBxHi.open;


  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('Select * from Receb_Baixa');
  DM_tabelas.ZQReceb_Baixa.open;


  DM_Tabelas.ZQCheque.Close;
  DM_Tabelas.ZQCheque.SQL.Clear;
  DM_Tabelas.ZQCheque.SQL.Add('select * from cheque_rec');
  DM_Tabelas.ZQCheque.open;


//  ZQRecBai.Close;

  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  JDEntrada.Datetext := datetostr(Date);
  JDBaixa.Datetext := datetostr(Date);
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text   := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;


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
    if ZQRecBai.FieldByName('marcar').Value = '0' then
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
    if  ZQRecBai.FieldByName('marcar').Value = '0' Then
      ZQRecBai.FieldByName('marcar').Value := '1'
    else
      ZQRecBai.FieldByName('marcar').Value := '0';
    Atualiza_tela;
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
  posi,posi2,contador,varreg : Integer;
  ql,varordem, varnumordem1, varnumordem2 : string;
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

  DXBBaixaGravar.Enabled := False;
  DXBFechar.Enabled := False;

  ZQRecBai.Filtered:=false;  
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;

  ql:=ZQRecBai.FieldByName('quadralote').Value;
  
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  varnumordem1 := '';
  varnumordem2 := '';
  If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
    DM_tabelas.ZQCobaRe.Edit;
  DM_tabelas.ZQCobaRe.FieldByName('entrada').Value := strtodate(JDEntrada.Datetext);
  DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').Value := CBTipobaixa.Text;
  DM_tabelas.ZQCobaRe.Post;
  CDSParcelas.First;
  varsomarepasse := 0;
  DM_tabelas.contador_bx.close;
  DM_tabelas.contador_bx.Open;
  DM_tabelas.contador_bx.Last;
  if DM_tabelas.contador_bx.RecordCount>0 then
     contador:=DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value
  else
    contador:=0;

  BarraVertical.Progress:=0;
  BarraVertical.Visible:=true;
  CDSParcelas.DisableControls;
  while not CDSParcelas.Eof do
  begin
    BarraVertical.Progress:=CDSParcelas.RecNo * 100 div CDSParcelas.RecordCount;
    inc(contador);
    if (CBTipobaixa.Text<>'QUITAÇÃO')  then
    begin
      DM_tabelas.ZQReceb_Baixa.Insert;
      DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
      DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);
      DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := strtodate(JDBaixa.Datetext); //DM_tabelas.ZQCobareHoje.Value;
      DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.Datetext);
      DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
//      DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
 //     DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
      DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
      DM_tabelas.ZQReceb_Baixa.Post;

    end
    else
    begin
      if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
      begin
        DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
        DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);        
        DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := strtodate(JDBaixa.Datetext); //DM_tabelas.ZQCobareHoje.Value;
        DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.Datetext);
        DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
//        DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
//        DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='Q';
        DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;

      end
      else
      begin
        DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQreceb_baixa.FieldByName('VR_AB').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
        DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);        
        DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.Datetext);
        DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='Q';
        DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;

      end;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').Value='S' Then
      varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;
   // if DM_tabelas.ZQTipodocreceb_receb.Value = 'S' Then
   // Begin
      if (CBTipobaixa.Text<>'QUITAÇÃO') then
      begin
        if empty(varnumordem1) Then
        Begin
          DM_Tabelas.ZQNumOrdem.Insert;
          DM_Tabelas.ZQNumOrdem.Post;
          varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
        end;
        DM_tabelas.ZQRecebimento.Insert;
        DM_tabelas.ZQRecebimento.FieldByName('documento').Value := CDSParcelasDocum.Value;
        DM_tabelas.ZQRecebimento.FieldByName('cliente').Value := strtoint(Ecodnome.Text);
        DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := DM_tabelas.ZQCobaRe.FieldByName('hoje').Value;
        DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQRecebimento.FieldByName('Valor').Value := CDSParcelasVrParc.Value;
        DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNERecebido.Value;
        DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQRecebimento.FieldByName('usuario').Value := DM_tabelas.ZQUsuario.FieldByName('idusuario').Value;
        DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
        DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;

        DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

        DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
        DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;
        DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;

        Dm_tabelas.ZQLoteamento.Filtered:=false;
        Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
        Dm_tabelas.ZQLoteamento.Filtered:=true;
        if Dm_tabelas.ZQLoteamento.RecordCount>0 then
           Dm_tabelas.ZQRecebimento.FieldByName('idloteamento').Value:=Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
         Dm_tabelas.ZQLoteamento.Filtered:=false;

        DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMNovostit.Text;
        DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := ZQRecBai.FieldByName('recpag').Value;
        DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := strtoint(EcodAdversaNova.text);
        DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := strtoint(varnumordem1);
        DM_tabelas.ZQRecebimento.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
        DM_tabelas.ZQRecebimento.Post;
      end
      else if (CBTipobaixa.Text='QUITAÇÃO')  then
      begin
        if empty(varnumordem1) Then
        Begin
          DM_Tabelas.ZQNumOrdem.Insert;
          DM_Tabelas.ZQNumOrdem.Post;
          varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
        end;
        DM_tabelas.ZQRecebimento.Insert;
        DM_tabelas.ZQRecebimento.FieldByName('documento').Value := CDSParcelasDocum.Value;
        DM_tabelas.ZQRecebimento.FieldByName('cliente').Value := strtoint(Ecodnome.Text);
        DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := DM_tabelas.ZQCobaRe.FieldByName('hoje').Value;
        DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQRecebimento.FieldByName('Valor').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
        DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNERecebido.Value;
        DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQRecebimento.FieldByName('usuario').Value := DM_tabelas.ZQUsuario.FieldByName('idusuario').Value;
        DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
        DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;

        DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

        DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
        DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;

        DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;
        Dm_tabelas.ZQLoteamento.Filtered:=false;
        Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
        Dm_tabelas.ZQLoteamento.Filtered:=true;
        if Dm_tabelas.ZQLoteamento.RecordCount>0 then
           Dm_tabelas.ZQRecebimento.FieldByName('idloteamento').Value:=Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
         Dm_tabelas.ZQLoteamento.Filtered:=false;

   //     if (CDSParcelasTipDoc.Value<>'CH') then
           DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := CDSParcelasVrParc.Value;

        DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;;
        DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMNovostit.Text;
        DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := ZQRecBai.FieldByName('recpag').Value;
        DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := strtoint(EcodAdversaNova.text);
        DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := strtoint(varnumordem1);

        DM_Tabelas.ZQRecebimento.FieldByName('Substituicao').Value:='Q';
        DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
        DM_tabelas.ZQRecebimento.Post;
        
      end;
    //end;
    if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').Value='S' Then
    Begin
      if (CBTipobaixa.Text<>'QUITAÇÃO')then
      begin
        DM_tabelas.ZQCaixa.Insert;
        DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
        DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
        DM_tabelas.ZQCaixa.FieldByName('documento').Value := CDSParcelasDocum.Value;
        DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
        DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
        DM_tabelas.ZQCaixa.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
        DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
          DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
        DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
        IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
          DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
        end;
        DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value +' | '+EMObsbx.Text;
        DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
        DM_tabelas.ZQCaixa.Post;
      end
      else
      begin
        if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
        begin
          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := (CDSParcelasVrParc.Value+CDSParcelasjuros.Value)-CDSParcelasdesconto.Value;
          DM_tabelas.ZQCaixa.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
          DM_tabelas.ZQCaixa.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
          IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
            DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
          end;
          DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value +' | '+EMObsbx.Text;
          DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value := 'Q';
          DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCaixa.Post;

        end;
      end;
    end;
    IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then
    Begin
      if (CBTipobaixa.Text<>'QUITAÇÃO') then
      begin
        DM_tabelas.ZQCheque.Insert;
        DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
        DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
        DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
        DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
        DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
        DM_tabelas.ZQCheque.FieldByName('emissao').Value := strtodate(JDEntrada.Datetext);
        DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
        DM_Tabelas.ZQCheque.FieldByName('numero').Value := CDSParcelasnumero_cheque.Value;        
        DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
        DM_tabelas.ZQCheque.Post;
      end
      else if (CBTipobaixa.Text='QUITAÇÃO') then
      begin
        DM_tabelas.ZQCheque.Insert;
        DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
        DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
        DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
        DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
        DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
        DM_tabelas.ZQCheque.FieldByName('emissao').Value := strtodate(JDEntrada.Datetext);
        DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
        DM_Tabelas.ZQCheque.FieldByName('numero').Value := CDSParcelasnumero_cheque.Value;
        DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value := 'Q';
        DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
        DM_tabelas.ZQCheque.Post;

      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').Value = 'S' Then Begin
      if (CBTipobaixa.Text<>'QUITAÇÃO') then
      begin
        DM_Tabelas.ZQMovBancaria.Insert;
        DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := CDSParcelascontabanco.Value;
        DM_TAbelas.ZQMovBancaria.FieldByName('documento').Value := CDSParcelasDocum.Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
        DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
        DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
        DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := CDSParcelasVenci.Value;
        DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').Value := CDSParcelasVrParc.Value;
        DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
        DM_TAbelas.ZQMovBancaria.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('obs').Value := EMObsbx.Text;
        DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
        DM_Tabelas.ZQMovBancaria.Post;
        //      CalcSaldo('','',False);
      end
      else if (CBTipobaixa.Text='QUITAÇÃO') then
      begin
        DM_Tabelas.ZQMovBancaria.Insert;
        DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := CDSParcelascontabanco.Value;
        DM_TAbelas.ZQMovBancaria.FieldByName('documento').Value := CDSParcelasDocum.Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
        DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
        DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
        DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := CDSParcelasVenci.Value;
        DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').Value := CDSParcelasVrParc.Value;
        DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
        DM_TAbelas.ZQMovBancaria.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
        DM_Tabelas.ZQMovBancaria.FieldByName('obs').Value := EMObsbx.Text;
        DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
        DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value := 'Q';
        DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
        DM_Tabelas.ZQMovBancaria.Post;

        //      CalcSaldo('','',False);
      end;
    end;
    CDSParcelas.Next;
  end;
  CDSParcelas.EnableControls;
  // histórico do QUITAÇÃO 06/08/2010
  CDSParcelas.Prior;
  zqQUITACAO.close;
  zqQUITACAO.Open;
  zqQUITACAO.Insert;
  ZQquitacao.FieldByName('idvenda').Value:=CDSParcelasvenda_idvenda.Value;
  ZQquitacao.FieldByName('datacessao').Value:=strtodate(JDEntrada.Datetext);
  ZQquitacao.FieldByName('hora').Value:=time;
  ZQquitacao.FieldByName('historico').Value:=EMNovostit.Text;
  zqQUITACAO.post;
  zqQUITACAO.close;
  CDSParcelas.Last;
  // até aqui

  BarraVertical.Progress:=0;
//  BarraVertical.Visible:=false;
  DM_tabelas.contador_bx.insert;
  DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value:=contador;
  DM_tabelas.contador_bx.Post;
  DM_tabelas.contador_bx.close;

  if not empty(varnumordem1) Then Begin
    ZQDiv_usado.SQL.Clear;
    ZQDiv_usado.SQL.Add('select * from dividido');
    ZQDiv_usado.SQL.Add(' where  ordem = '+quotedstr(ZQRecBai.FieldByName('numordem').Text));
    ZQDiv_usado.Open;
    varsomarepasseautomatico := 0;
    while not ZQDiv_usado.Eof do begin
      DM_Tabelas.ZQDividido.Insert;
      DM_Tabelas.ZQDividido.FieldByName('participante').Value := ZQDiv_usado.FieldByName('participante').Value;
      DM_Tabelas.ZQDividido.FieldByName('percentual').Value := ZQDiv_usado.FieldByName('percentual').Value;
      DM_Tabelas.ZQDividido.FieldByName('planodeconta').Value := ZQDiv_usado.FieldByName('planodeconta').Value;
      DM_Tabelas.ZQDividido.FieldByName('tipodocumento').Value := ZQDiv_usado.FieldByName('tipodocumento').Value;
      DM_Tabelas.ZQDividido.FieldByName('ordem').Value := strtoint(varnumordem1);
      DM_Tabelas.ZQDividido.Post;
      ZQDiv_usado.Next;
    end;
  end;
  ZQRecBai.First;
  varpago := 0;
  Memo1.Clear;
  baixatot := XNERecebido.Value;


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
  While not ZqRecBai.Eof do
  Begin
    BarraVertical.Progress:=ZqRecBai.RecNo * 100 div ZqRecBai.RecordCount;
    DM_tabelas.ZQCheque.First;
    if DM_tabelas.ZQCheque.RecordCount>0 then
    begin
      while not DM_tabelas.ZQCheque.Eof do
      begin
        if DM_tabelas.ZQCheque.FieldByName('idrecebimento').Value=ZQRecBai.FieldByName('idrecebimento').Value then
        begin
           DM_tabelas.ZQCheque.Delete;
        end
        else
        begin
           DM_tabelas.ZQCheque.Next;
        end;
      end;
    end;
//    DM_tabelas.ZQMovBancariavinculo.Value := 'R'+DM_tabelas.ZQCobaRecotagem.Text;

  //  DM_tabelas.ZQCaixavinculo.Value := 'P'+DM_tabelas.ZQCobaRecotagem.Text;

    ZQReceb_Baixa.Filtered:=false;
    ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(ZQRecBai.FieldByName('documento').Value))+' and sq='+quotedstr(alltrim(inttostr(ZQRecBai.FieldByName('sq').value)));
    ZQReceb_Baixa.Filtered:=true;
    ZQReceb_Baixa.First;
    if ZQReceb_Baixa.RecordCount>0 then
    begin
      while not ZQReceb_Baixa.Eof do
      begin
        if ZQreceb_baixa.FieldByName('VR_AB').Value>0 then
        begin
          ZQReceb_Baixa.Delete;
        end
        else
        begin
          ZQReceb_Baixa.Next;
        end;
      end;
    end;
    ZQReceb_Baixa.Filtered:=false;
    DM_tabelas.ZQReBxHi.First;
    if DM_tabelas.ZQReBxHi.RecordCount>0 then
    begin
      while not DM_tabelas.ZQReBxHi.Eof do
      begin
        if DM_tabelas.ZQReBxHi.FieldByName('idrecib').Value=ZQRecBai.FieldByName('idrecebimento').Value then
        begin
           DM_tabelas.ZQReBxHi.Delete;
        end
        else
        begin
           DM_tabelas.ZQReBxHi.Next;
        end;
      end;
    end;

    recebimento_historico_excluidos.Insert;
    recebimento_historico_excluidos.FieldByName('idrecebimento').value:= ZQRecBai.FieldByName('idrecebimento').value;
    recebimento_historico_excluidos.FieldByName('documento').value:= ZQRecBai.FieldByName('documento').value;
    recebimento_historico_excluidos.FieldByName('cliente').value:= ZQRecBai.FieldByName('cliente').value;
    recebimento_historico_excluidos.FieldByName('usuario').value:= ZQRecBai.FieldByName('usuario').value;
    recebimento_historico_excluidos.FieldByName('Dt_Entrada').value:= ZQRecBai.FieldByName('Dt_Entrada').value;
    recebimento_historico_excluidos.FieldByName('Dt_Vencimento').value:= ZQRecBai.FieldByName('Dt_Vencimento').value;
    recebimento_historico_excluidos.FieldByName('Valor').value:= ZQRecBai.FieldByName('Valor').value;
    recebimento_historico_excluidos.FieldByName('Observ').value:= ZQRecBai.FieldByName('Observ').value;
    recebimento_historico_excluidos.FieldByName('VrDoc').value:= ZQRecBai.FieldByName('VrDoc').value;
    recebimento_historico_excluidos.FieldByName('ordem').value:= ZQRecBai.FieldByName('ordem').value;
    recebimento_historico_excluidos.FieldByName('TipDoc').value:= ZQRecBai.FieldByName('TipDoc').value;
    recebimento_historico_excluidos.FieldByName('saldo').value:= ZQRecBai.FieldByName('saldo').value;
    recebimento_historico_excluidos.FieldByName('marcar').value:= ZQRecBai.FieldByName('marcar').value;
    recebimento_historico_excluidos.FieldByName('RefBaixa').value:= ZQRecBai.FieldByName('RefBaixa').value;
    recebimento_historico_excluidos.FieldByName('refvinda').value:= ZQRecBai.FieldByName('refvinda').value;
    recebimento_historico_excluidos.FieldByName('contabil').value:= ZQRecBai.FieldByName('contabil').value;
    recebimento_historico_excluidos.FieldByName('empresa').value:= ZQRecBai.FieldByName('empresa').value;
    recebimento_historico_excluidos.FieldByName('custodaparcela').value:= ZQRecBai.FieldByName('custodaparcela').value;
    recebimento_historico_excluidos.FieldByName('origem').value:= ZQRecBai.FieldByName('origem').value;
    recebimento_historico_excluidos.FieldByName('adversa').value:=ZQRecBai.FieldByName('adversa').value;
    recebimento_historico_excluidos.FieldByName('recpag').value:= ZQRecBai.FieldByName('recpag').value;
    recebimento_historico_excluidos.FieldByName('numordem').value:= ZQRecBai.FieldByName('numordem').value;
    recebimento_historico_excluidos.FieldByName('idloteamento').value:= ZQRecBai.FieldByName('idloteamento').value;
    recebimento_historico_excluidos.FieldByName('venda_idvenda').value:= ZQRecBai.FieldByName('venda_idvenda').value;
    recebimento_historico_excluidos.FieldByName('quadralote').value:= ZQRecBai.FieldByName('quadralote').value;
    recebimento_historico_excluidos.FieldByName('numboleto').value:= ZQRecBai.FieldByName('numboleto').value;
    recebimento_historico_excluidos.FieldByName('Substituicao').value:=ZQRecBai.FieldByName('Substituicao').value;
    recebimento_historico_excluidos.FieldByName('sq').value:= ZQRecBai.FieldByName('sq').value;
    recebimento_historico_excluidos.FieldByName('somar').value:= ZQRecBai.FieldByName('somar').value;
    recebimento_historico_excluidos.FieldByName('motivo').Value:='Q';
    recebimento_historico_excluidos.Post;
    ZqRecBai.Next;
  end;
  recebimento_historico_excluidos.close;
  recebimento_historico_excluidos.open;
  recebimento_historico_excluidos.last;


  ZqRecBai.First;
  While not ZqRecBai.Eof do
  Begin
    ZqRecBai.Delete;
  end;
  ZQReceb_Baixa.close;

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
  while not ZQRatear.Eof do begin
    if CDS_dividido.Locate('participante;contabil;tipodoc',VarArrayOf([ZQRatear.FieldByName('participante').Value,ZQRatear.FieldByName('planodeconta').Value,ZQRatear.FieldByName('tipodocumento').Value]),[]) then
      CDS_dividido.Edit
    else
      CDS_dividido.Insert;
    CDS_divididoparticipante.Value := ZQRatear.FieldByName('participante').Value;
    CDS_divididovr_base.Value := CDS_divididovr_base.Value + ZQRatear.FieldByName('vr_base').Value;
    CDS_divididovr_calc.Value := CDS_divididovr_calc.Value + ZQRatear.FieldByName('vr_divi').Value;
    CDS_divididocontabil.Value := ZQRatear.FieldByName('planodeconta').Value;
    CDS_divididotipodoc.Value := ZQRatear.FieldByName('tipodocumento').Value;
    CDS_divididoobs.Value := CDS_divididoobs.Value+ZQRatear.FieldByName('idrecebimento').Text+' | '+ZQRatear.FieldByName('ordem').Value+' | '+ZQRatear.FieldByName('vr_base').Text+' x '+ZQRatear.FieldByName('percentual').Text+' = '+ZQRatear.FieldByName('vr_divi').Text+chr(13);
    varsomarepasseautomatico := varsomarepasseautomatico + ZQRatear.FieldByName('vr_divi').Value;
    ZQRatear.Next;
  end;
  CDS_dividido.First;
  if CDS_dividido.RecordCount>0 Then Begin
    DM_Tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
  end;
  while not CDS_dividido.Eof do begin
    DM_tabelas.ZQRecebimento.Insert;
    DM_tabelas.ZQRecebimento.FieldByName('documento').Value := 'Rateio Bx - '+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
    DM_tabelas.ZQRecebimento.FieldByName('cliente').Value :=  CDS_divididoparticipante.Value;
    DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := strtodate(JDEntrada.datetext);
    DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := strtodate(JDEntrada.datetext);
    DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDS_divididocontabil.Value;
    DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMNovostit.Text+chr(13)+CDS_divididoobs.Value;
    DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := 'P';
    DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDS_divididotipodoc.Value;
    DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text+'-'+inttostr(CDS_dividido.RecNo)+'/'+inttostr(CDS_dividido.RecordCount);
    DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
    DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value :=  CDS_divididovr_base.Value;
    DM_tabelas.ZQRecebimento.FieldByName('Valor').Value :=  CDS_divididovr_calc.Value;
    DM_tabelas.ZQRecebimento.FieldByName('saldo').Value :=  DM_tabelas.ZQRecebimento.FieldByName('Valor').Value;
    DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := DM_Tabelas.ZQConfiguracoes.FieldByName('paticipantedefault').Value;

    DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

    DM_tabelas.ZQRecebimento.Post;
    CDS_dividido.Next;
  end;
  varreg := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
  DM_tabelas.ZQRecebimento.Refresh;
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;
//  DM_tabelas.ZQRecebimento.Refresh;
  DM_Tabelas.ZQRecebimento.Locate('idrecebimento',varreg,[]);
  FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
  FrmImpRecibo.Nomecli.Text := Enomes.Text;
  FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;

  ///
  ZQRecBai.Refresh;




  // alteração do valor de venda 28/08/2010
  ZqRecBai.First;
  posi:=pos('-',ql);
  dec(posi);

  posi2:=pos('-',ql);
  inc(posi2);

  ZQEmpree.close;
  ZQEmpree.SQL.Clear;
  ZQEmpree.SQL.Add('Select * from imovel where quadra='+quotedstr(copy(ql,1,posi))+' and lote='+quotedstr(copy(ql,posi2,3)));
  ZQEmpree.open;
  ZQVenda.open;
  ZQEntrada.Close;
  ZQEntrada.SQL.Clear;
  ZQEntrada.SQL.Add('Select *, sum(valor) as entrada from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-E-%''  group by quadralote  order by ordem');
  ZQEntrada.open;

  ZQParcela.Close;
  ZQParcela.SQL.Clear;
  ZQParcela.SQL.Add('Select *, sum(valor) as parcela from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-P-%''  group by quadralote  order by ordem');
  ZQParcela.open;

  ZQVenda.Edit;
  ZQVenda.FieldByName('valorvenda').Value:=ZQEntrada.FieldByName('entrada').Value+ZQParcela.FieldByName('parcela').Value;
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
    FrmRelRecibodeQuita.RLReport1.Preview;
  end; }
  ZQRecBai.Filtered := False;
  DM_tabelas.ZQRecebimento.Refresh;

  DM_tabelas.ZQReceb_Baixa.Refresh;
  DM_Tabelas.ZQReBxHi.Refresh;
  xvequta.Value:=0;

  JDEntrada.SetFocus;
  DXBFechar.Enabled := true;
end;

procedure TFrmquitacao.DBCBTIPDOCExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if Empty(DBCBTIPDOC.Text) Then Begin
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipDoc.Text,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').Value='S' Then Begin
      if (datetostr(CDSParcelasVenci.Value)<>vecto1.Datetext) and ((CDSParcelasTipDoc.Value<>'DP')) Then BEgin
        mensagem('Este tipo de documento só aceita pagamento a vista!!!');
        EContabil.SetFocus;
        Exit;
      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').Value='S' then Begin
      GBBanco.Visible := True;
      DBGBanco.SetFocus;
      exit;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value='S' then Begin
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
  VarDoc := copy(VarDoc,1,length(VarDoc)-1);
  CDSParcelas.RecNo := varregis;
  vardif := XNERecebido.Value - varsoma;
  if CDSParcelas.RecNO< XNEParcelas.Value Then Begin
    CDSParcelas.RecNO := varregis+1;
    while not CDSParcelas.Eof do Begin
      CDSParcelas.Edit;
      CDSParcelasVrParc.Value := truncar(vardif / (XNEParcelas.Value-varregis),2);
      vardesc := vardesc - CDSParcelasdesconto.Value;
      CDSParcelas.Next;
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
    EContabil.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').Value;
    if not (CDSParcelas.State in [DsInsert, DsEdit]) Then
      CDSParcelas.Edit;
    CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
    CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').Value;
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
    Enomes.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
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
  if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value='S' then Begin
    CDSParcelasagencia.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').Value;
    CDSParcelasConta.Value := DM_TAbelas.ZQContaBancaria.FieldByName('n_conta').Value;
    CDSParcelasbanco.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').Value;
    CDSParcelasdono.Value := DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').Value;
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
  if not DXBFechar.Focused Then Begin
    if not AchaPlanoDeConta(450, 30, 'S', EPlanocontabanco.Text, '') then Begin
      EPlanocontabanco.SetFocus;
      exit;
    end;
    EPlanocontabanco.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').Value;
    CDSParcelascontabanco.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
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
  if ( empty(alltrim(Elote.Text))) then
  begin
    ZQRecBai.close;
    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
       ZQRecBai.SQL.Add('select * from recebimento where (recpag='+quotedstr(varrecpag)+') and (saldo>0 and  not (saldo is null)) and (cliente ='+Ecodnome.Text+')');
    if not empty(EcodAdversa.text) Then
    Begin
        ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
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
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
        DM_Tabelas.ZQTipodoc.Next;
      end;
      vartipo := copy(vartipo,1,length(vartipo)-1);
      ZQRecBai.SQL.Add('select * from recebimento where tipdoc in ('+vartipo+') and (recpag='+quotedstr(varrecpag)+') and (saldo>0 and  not (saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
      ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
    end;
  end
  else
  begin
    ZQRecBai.close;
    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    begin
//       ZQRecBai.SQL.Add('select * from recebimento where ( quadralote='+quotedstr(alltrim(Equadra.text))+'-'+quotedstr(alltrim(Elote.Text))+') and (recpag='+quotedstr(varrecpag)+') and (saldo>0) and  (not saldo is null));
       ZQRecBai.SQL.Add('select * from recebimento where quadralote='+quotedstr(alltrim(Elote.Text))+' and recpag='+quotedstr(varrecpag)+' and saldo>0 ');
//    if not empty(EcodAdversa.text) Then Begin
//       ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;

       ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
    end
    else if RBCheque.Checked then Begin
      Label37.Visible:=false;
      pini.Visible:=false;
      pini.Clear;
      Label38.Visible:=false;
      pfim.Visible:=false;
      pfim.Clear;

      vartipo := '';
      DM_TAbelas.ZQTipodoc.First;
      while not DM_Tabelas.ZQTipodoc.Eof do begin
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
        DM_Tabelas.ZQTipodoc.Next;
      end;
      vartipo := copy(vartipo,1,length(vartipo)-1);
       ZQRecBai.SQL.Add('select * from recebimento where quadralote='+quotedstr(alltrim(Elote.Text))+' and recpag='+quotedstr(varrecpag)+' and saldo>0 ');
//      ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
    end;
  end;
  ZQRecBai.Open;
  XEEmpresend.Text := ZQRecBai.FieldByName('nomeempreend').Value;
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  if RBParticipante.Checked Then
    while not ZQRecBai.Eof do begin
      ZQRecBai.Edit;
//      if ZQRecBaiDt_Vencimento.Value<=date Then
        ZQRecBai.FieldByName('marcar').Value := '0';
  //    else
    //    ZQRecBaimarcar.Value := '1';
      ZQRecBai.Post;
      ZQRecBai.Next;
    end;
  ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value,[]);
  if RBPonteiro.Checked Then Begin
    ZQRecBai.Edit;
    ZQRecBai.FieldByName('marcar').Value := '0';
    ZQRecBai.Post;
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
    ZQRecBai.FieldByName('marcar').Value := '1';
    ZQRecBai.Post;
    ZQRecBai.Next;
  end;
  ZQRecBai.Edit;
  ZQRecBai.RecNo := varreg;
  ZQRecBai.FieldByName('marcar').Value := '0';
  ZQRecBai.Post;
  XNEMarcado.Value := ZQRecBai.FieldByName('saldo').Value;
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
    Atualiza_tela;
    JDBaixa.Datetext:=JDEntrada.Datetext;
    vecto1.Datetext:=JDEntrada.Datetext;
    DBGBaixando.SetFocus;
  end;
end;

procedure TFrmquitacao.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  xvequta.Value:=0;
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
  
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.last;
end;

procedure TFrmquitacao.ZQRecBaiCalcFields(DataSet: TDataSet);
var
  vardI, vardT : Tdate;
  varsal : Double;
begin
    if (ZQRecBai.FieldByName('saldo').Value>0) and (ZQRecBai.FieldByName('recpag').Value='R') and (CBSaldoCorrigido.Checked)  Then Begin
    varsal := ZQRecBai.FieldByName('Valor').value;
    ZQRecebBxTemp.First;
    ZQRecebBxTemp.Refresh;
    vardI := ZQRecBai.FieldByName('Dt_Vencimento').Value;
    vardT := strtodate(FrmQUITACAO.JDEntrada.Datetext);
    while (not ZQRecebBxTemp.Eof) and (vardT>=ZQRecebBxTemp.FieldByName('dataref').Value) do begin
      varsal := Calcjuros(vardI,ZQRecebBxTemp.FieldByName('dataref').Value,varsal,XNJuros.Value);
      varsal := varsal-(ZQRecebBxTemp.FieldByName('Vr_rec').Value*ZQRecebBxTemp.FieldByName('percent_usado').Value/100);
      vardI := ZQRecebBxTemp.FieldByName('dataref').Value;
      ZQRecebBxTemp.Next;
    end;
    ZQRecBai.FieldByName('saldocorrig').Value := Calcjuros(vardI,vardT,varsal,XNJuros.Value);
  end
  else if ZQRecBai.FieldByName('saldo').Value=0 Then
    ZQRecBai.FieldByName('saldocorrig').Value := 0
  else
    ZQRecBai.FieldByName('saldocorrig').Value := ZQRecBai.FieldByName('saldo').Value;
end;

procedure TFrmquitacao.CDSParcelasCalcFields(DataSet: TDataSet);
begin
  CDSParcelasPercent.Value := CDSParcelasVrParc.Value*100/XNERecebido.Value;
//  CDSParcelasjuros.Value := XNEJuros.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmquitacao.ZQRatearCalcFields(DataSet: TDataSet);
begin
    ZQRatear.FieldByName('vr_base').Value := ZQRatear.FieldByName('vr_rec').Value * ZQRatear.FieldByName('percent_usado').Value / 100;
    ZQRatear.FieldByName('vr_divi').Value := ZQRatear.FieldByName('vr_base').Value * ZQRatear.FieldByName('percentual').Value /100;

end;

procedure TFrmquitacao.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').Value = 'R' Then
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
  DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value,[]);
  varsoma := 0;
  for varvezes := 1 to varfinal do Begin
    CDSParcelas.Insert;
    CDSParcelasOrdem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
    CDSParcelasjuros.Value :=  truncar(XNEJuros.Value/varfinal,2);
    CDSParcelasVrParc.Value := truncar(XNERecebido.Value/varfinal,2);
    if ZQRecBai.FieldByName('recpag').Value='R' Then
      CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').Value
    else
      CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').Value;
      CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').Value;
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

    CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').Value;
    CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').Value;
    CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').Value;
    CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
    CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').Value;
    CDSParcelassq.Value := ZQRecBai.FieldByName('sq').Value;
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
        Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
        EcodAdversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
      end
      else
        Eadversa.SetFocus;
    end
    else
      EcodAdversa.Text := '';
    filtrabaixa;
    Atualiza_tela;
  end;
end;

procedure TFrmquitacao.EAdversaNovaExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not empty(EAdversaNova.Text) Then Begin
      if AchaCLiente(400, 300, EAdversaNova.Text) Then Begin
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
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
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text   := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True; 
  Atualiza_tela;


  DBGBaixando.SetFocus;
end;

procedure TFrmquitacao.btndesClick(Sender: TObject);
begin
  ZQRecBai.First;
  while not ZQRecBai.Eof do
  begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').Value = '0' Then
      ZQRecBai.FieldByName('marcar').Value := '1'
    else
      ZQRecBai.FieldByName('marcar').Value := '0';
    ZQRecBai.post;
    ZQRecBai.Next;
  end;
  Atualiza_tela;
end;

procedure TFrmquitacao.EloteExit(Sender: TObject);
var
  vartipo : string;

begin
 if not empty(Elote.Text) then
 begin
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from recebimento where recpag=''R'' and quadralote='+quotedstr(Elote.text)+' and saldo>0');
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.Recordcount;
  JDEntrada.Datetext := datetostr(date);
  JDBaixa.Datetext   := datetostr(date);
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text   := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;
  Atualiza_tela;


  DBGBaixando.SetFocus;
 end;

end;

procedure TFrmquitacao.XNERecebidoEnter(Sender: TObject);
begin
  if xvequta.Value>0 then
     XNERecebido.Value:=xvequta.Value;
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
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaidocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaicliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaimarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBairefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaicontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaicustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBairecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainomeadversa', 'nomeadversa', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaivenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBainomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaisomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaidescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecBai', 'ZQRecBaiData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemporigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTemprecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRecebBxTemp', 'ZQRecebBxTempsubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRateartipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearordem_1', 'ordem_1', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearvr_base', 'vr_base', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQRatear', 'ZQRatearvr_divi', 'vr_divi', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciborefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciborecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibovalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibojuros_1', 'juros_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensRecibodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQItensRecibo', 'ZQItensReciboData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQDiv_usado', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadotipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQDiv_usado', 'ZQDiv_usadoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQReceb_Baixa', 'ZQReceb_Baixavalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'recebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmotivo', 'motivo', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'zqquitacao', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaonome_parte', 'nome_parte', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaohistorico', 'historico', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaohora', 'hora', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'zqquitacao', 'zqquitacaocontrole', 'controle', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQVenda', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreelote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreelado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreematri', 'matri', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQEntrada', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradasomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmquitacao, 'ZQParcela', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParceladocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmquitacao, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.


