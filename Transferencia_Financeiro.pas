

unit Transferencia_Financeiro;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, wwdbdatetimepicker, StdCtrls, Mask, DBCtrls,
  XBanner, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, ImgList, XNum, DBClient, XDBDate, XDBNum, ToolEdit, ComCtrls,
  JvExComCtrls, JvDateTimePicker, XEdit, XDBEdit, ExtCtrls;
type
  TFrmtransferencia = class(TForm)
    Label10: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    DBEBaixaDocum: TDBEdit;
    DBEBaixaVrRec: TDBEdit;
    WDTBaixaEntra: TwwDBDateTimePicker;
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
    JDEntrada: TJvDateTimePicker;
    RBParticipante: TRadioButton;
    RBCheque: TRadioButton;
    DTPIni: TJvDateTimePicker;
    DTPFim: TJvDateTimePicker;
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
    CBTipobaixa: TComboBox;





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
    JDBaixa: TJvDateTimePicker;

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
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frmtransferencia: TFrmtransferencia;
  VArDoc, varrecpag : String;

implementation

uses tabelas, Cad_Recebimento, funcoes, ImpRecibo, PesqRecebimento_bx, uRuntimeFields;

{$R *.dfm}

procedure TFrmtransferencia.Atualiza_tela;
Var
  Varreg, varmarc, varvenc : integer;
Begin
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
  Label31.Caption := inttostr(varmarc);
  Label33.Caption := inttostr(varvenc);
  XNERecebido.Value := XNEMarcado.Value;
  ZQRecBai.RecNo := Varreg;
End;

procedure TFrmtransferencia.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmtransferencia.FormShow(Sender: TObject);
begin
  ZQRecebBxTemp.open;
  dm_tabelas.ZQLoteamento.Open;
  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  JDEntrada.Date := Date;
  JDBaixa.Date := Date;  
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Date := date;
  DTPFim.Date := date;
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


  CBPeriodoEnt.ItemIndex := 0;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;
  CBTipobaixa.SetFocus;
end;

procedure TFrmtransferencia.DBGBaixandoDrawColumnCell(Sender: TObject;
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

procedure TFrmtransferencia.DBGBaixandoColEnter(Sender: TObject);
begin
  if DBGBaixando.SelectedField = ZQRecBai.FieldByName('marcar') then
    DBGBaixando.Options := DBGBaixando.Options - [dgEditing]
  else
    DBGBaixando.Options := DBGBaixando.Options + [dgEditing];
end;

procedure TFrmtransferencia.DBGBaixandoMouseUp(Sender: TObject;
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

procedure TFrmtransferencia.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    close;
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmtransferencia.XNERecebidoExit(Sender: TObject);
begin
   if XNEMarcado.Value >= XNERecebido.Value Then
     XNEResto.Value := XNEMarcado.Value - (XNERecebido.Value + XNEDesconto.Value)
   else
     XNEJuros.Value :=  XNERecebido.Value - XNEMarcado.Value;
end;

procedure TFrmtransferencia.DXBBaixaGravarClick(Sender: TObject);
Var
  Varpago, percentbaixa, baixatot, varjuros, varx,
  varsomarepasse, varsomarepasseautomatico : Double;
  contador,varreg : Integer;
  varordem, varnumordem1, varnumordem2 : string;
begin
  if empty(CBTipobaixa.Text) then
  begin
    showmessage('Selecione o tipo de operação !');
    CBTipobaixa.SetFocus;
    exit;
  end;  

  ZQReceb_Baixa.open;
  ZQReceb_Baixa.Filtered:=false;
  ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(CDSParcelasDocum.AsString))+' and sq='+quotedstr(alltrim(inttostr(CDSParcelassq.value)));
  ZQReceb_Baixa.Filtered:=true;
  if ZQReceb_Baixa.RecordCount>0 then
     ZQReceb_Baixa.Delete;
  ZQReceb_Baixa.Filtered:=false;
  ZQReceb_Baixa.close;

  DXBBaixaGravar.Enabled := False;
  DXBFechar.Enabled := False;
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  varnumordem1 := '';
  varnumordem2 := '';
  If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
    DM_tabelas.ZQCobaRe.Edit;
  DM_tabelas.ZQCobaRe.FieldByName('entrada').Value := JDEntrada.Date;
  DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').Value := CBTipobaixa.Text;
  DM_tabelas.ZQCobaRe.Post;
  CDSParcelas.First;
  varsomarepasse := 0;
  DM_tabelas.contador_bx.Open;
  DM_tabelas.contador_bx.Last;
  if DM_tabelas.contador_bx.RecordCount>0 then
     contador:=DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value
  else
    contador:=0;     

  while not CDSParcelas.Eof do begin
    inc(contador);
    if (CBTipobaixa.Text<>'SUBSTITUIÇÃO')  then
    begin
      DM_tabelas.ZQReceb_Baixa.Insert;
      DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSParcelasVrParc.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := JDBaixa.Date; //DM_tabelas.ZQCobareHoje.Value;
      DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := JDEntrada.Date;
      DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
      DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSParcelasdesconto.Value;
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
        DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSParcelasVrParc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := JDBaixa.Date; //DM_tabelas.ZQCobareHoje.Value;
        DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := JDEntrada.Date;
        DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
        DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSParcelasdesconto.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='S';
        DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;        
        DM_tabelas.ZQReceb_Baixa.Post;
      end
      else
      begin
        DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQreceb_baixa.FieldByName('VR_AB').Value := CDSParcelasVrParc.Value;
  //      DM_tabelas.ZQReceb_BaixaDt_rec.Value := DM_tabelas.ZQCobareHoje.Value;
        DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := JDEntrada.Date;
        DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
//        DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
//        DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='S';
        DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;
      end;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').Value='S' Then
      varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;
    if DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').Value = 'S' Then
    Begin
      if (CBTipobaixa.Text<>'SUBSTITUIÇÃO') then
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
        DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := CDSParcelasVrParc.Value;

        DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
        DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;
        DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;

        Dm_tabelas.ZQLoteamento.Filtered:=false;
        Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.AsString);
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
      else if (CBTipobaixa.Text='SUBSTITUIÇÃO')  then
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

        DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
        DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;

        DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;
        Dm_tabelas.ZQLoteamento.Filtered:=false;
        Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.AsString);
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

        DM_Tabelas.ZQRecebimento.FieldByName('Substituicao').Value:='S';
        DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
        DM_tabelas.ZQRecebimento.Post;
      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').Value='S' Then
    Begin
      if (CBTipobaixa.Text<>'SUBSTITUIÇÃO')then
      begin
        DM_tabelas.ZQCaixa.Insert;
        DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
        DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSParcelasVrparc.Value;
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
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSParcelasVrparc.Value;
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
          DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value := 'S';
          DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCaixa.Post;
        end;
      end;
    end;
    IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then
    Begin
      if (CBTipobaixa.Text<>'SUBSTITUIÇÃO') then
      begin
        DM_tabelas.ZQCheque.Insert;
        DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
        DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
        DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
        DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
        DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
        DM_tabelas.ZQCheque.FieldByName('emissao').Value := JDEntrada.Date;
        DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
        DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;        
        DM_tabelas.ZQCheque.Post;
      end
      else if (CBTipobaixa.Text='SUBSTITUIÇÃO') then
      begin
        DM_tabelas.ZQCheque.Insert;
        DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
        DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
        DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
        DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
        DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
        DM_tabelas.ZQCheque.FieldByName('emissao').Value := JDEntrada.Date;
        DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
        DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value := 'S';
        DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
        DM_tabelas.ZQCheque.Post;
      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').Value = 'S' Then Begin
      if (CBTipobaixa.Text<>'SUBSTITUIÇÃO') then
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
      else if (CBTipobaixa.Text='SUBSTITUIÇÃO') then
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
        DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value := 'S';
        DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
        DM_Tabelas.ZQMovBancaria.Post;
        //      CalcSaldo('','',False);
      end;
    end;
    CDSParcelas.Next;
  end;
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
  While not ZqRecBai.Eof do Begin
    DM_tabelas.ZQReBxHi.Insert;
    percentbaixa := ZQRecBai.FieldByName('saldocorrig').Value * 100 / XNEMarcado.Value;
    DM_Tabelas.ZQReBxHi.FieldByName('descon').Value := XNEDesconto.Value * percentbaixa/100;
    if ZQRecBai.FieldByName('saldocorrig').Value<=baixatot Then
      varpago := ZQRecBai.FieldByName('saldocorrig').Value
    else
      varpago := baixatot+DM_Tabelas.ZQReBxHi.FieldByName('descon').Value;
    varjuros := XNEJuros.Value* percentbaixa/100;
    if varjuros <=varpago Then
      DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := varjuros
    else
      DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := varpago;
    varpago := varpago + DM_Tabelas.ZQReBxHi.FieldByName('juros').Value;
    DM_Tabelas.ZQReBxHi.FieldByName('valor').Value := varpago-DM_Tabelas.ZQReBxHi.FieldByName('descon').Value;
    DM_Tabelas.ZQReBxHi.FieldByName('data').Value := JDEntrada.Date;
    DM_Tabelas.ZQReBxHi.FieldByName('percent_usado').Value := percentbaixa;
    DM_tabelas.ZQReBxHi.FieldByName('refer').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
    DM_tabelas.ZQReBxHi.FieldByName('idrecib').Value := ZQRecBai.FieldByName('idrecebimento').Value;
    DM_tabelas.ZQReBxHi.FieldByName('sq').Value := contador;
    DM_Tabelas.ZQReBxHi.Post;


    ZQRecBai.Edit;
    varx := varpago-DM_Tabelas.ZQReBxHi.FieldByName('juros').Value;
//    varx := varx + DM_Tabelas.ZQReBxHidescon.Value;
    ZQRecBai.FieldByName('saldo').Value := truncar(ZQRecBai.FieldByName('saldo').Value - varx,2);
    ZQRecBai.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
    if (CBTipobaixa.Text='SUBSTITUIÇÃO') then
    begin
//      ZQRecBaisubstituicao.Value:='S';
      ZQRecBai.FieldByName('somar').Value:='N';
//      ZQRecBaisq.Value := contador;
    end;
    if pos('-',ZQRecBai.FieldByName('ordem').AsString)=0 Then
      Varordem := Varordem + ZQRecBai.FieldByName('idrecebimento').Text+' - ';
    if pos(ZQRecBai.FieldByName('nomeadversa').AsString,Memo1.Text)=0 Then
      Memo1.Lines.Add(ZQRecBai.FieldByName('nomeadversa').AsString);
    ZqRecBai.Next;
  end;
  CDS_dividido.Close;
  CDS_dividido.CreateDataSet;
  ZQRatear.SQL.Clear;
  ZQRatear.SQL.Add('select idrecebimento, h.valor, percent_usado, b.vr_rec, participante, percentual,planodeconta, dv.tipodocumento, re.ordem, dv.ordem from recbxhist as H');
  ZQRatear.SQL.Add(' join recebimento as re ON H.idrecib=RE.idrecebimento');
  ZQRatear.SQL.Add(' join dividido as dv ON dv.ordem=numordem');
  ZQRatear.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
  ZQRatear.SQL.Add(' join tipodocumento as D ON B.tipdoc=D.tipodoc');
  ZQRatear.SQL.Add(' where somapaga = '+quotedstr('S')+' and refer = '+QuotedStr(DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsString));
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
    DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := JDEntrada.date;
    DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := JDEntrada.date;
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
    DM_tabelas.ZQRecebimento.Post;
    CDS_dividido.Next;
  end;
  varreg := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
  DM_tabelas.ZQRecebimento.Refresh;
  DM_Tabelas.ZQRecebimento.Locate('idrecebimento',varreg,[]);
  FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
  FrmImpRecibo.Nomecli.Text := Enomes.Text;
  FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;

  ///
  ZQRecBai.Refresh;
  DBGBaixando.Refresh;
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
    FrmRelRecibodeQuita.ZQForma.SQL.Add('Select * from receb_baixa where refbaixa='+QuotedStr(DM_tabelas.ZQCobaRecotagem.Text));
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
  CBTipobaixa.SetFocus;
  DXBFechar.Enabled := true;  
end;

procedure TFrmtransferencia.DBCBTIPDOCExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if Empty(DBCBTIPDOC.Text) Then Begin
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipDoc.Text,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').Value='S' Then Begin
      if (datetostr(CDSParcelasVenci.Value)<>datetostr(JDEntrada.Date)) and ((CDSParcelasTipDoc.Value<>'DP')) Then BEgin
        showmessage('Este tipo de documento só aceita pagamento a vista!!!');
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

procedure TFrmtransferencia.DXBBaixaGravarEnter(Sender: TObject);
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

procedure TFrmtransferencia.DBCBTIPDOCEnter(Sender: TObject);
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

procedure TFrmtransferencia.DBEBaixaVrRecExit(Sender: TObject);
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
    VarDoc := VarDoc + quotedstr(CDSParcelasDocum.AsString)+',';
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

procedure TFrmtransferencia.EContabilExit(Sender: TObject);
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
    EContabil.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    if not (CDSParcelas.State in [DsInsert, DsEdit]) Then
      CDSParcelas.Edit;
    CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
    CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').Value;
  end;
//    CDSParcelasTipDoc.Value := FrmAchaPlanodeConta.ZQTempPlanoContasdoccomum.Value;
end;

procedure TFrmtransferencia.XDBNumEdit1Exit(Sender: TObject);
begin
  CDSParcelasVrParc.Value := XNERecebido.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmtransferencia.DBEBaixaDocumExit(Sender: TObject);
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
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (pos(quotedstr(CDSParcelasDocum.AsString),VarDoc)>0) Then Begin
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

procedure TFrmtransferencia.EContabilEnter(Sender: TObject);
begin
  if empty(Econtabil.Text) Then Begin
    Econtabil.Text := DM_Tabelas.ZQConfiguracoes.FieldByName('recebe_planodecontas').Text;
  end;
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;
  DBText1.Caption := CDSParcelasOrdem.AsString;
end;

procedure TFrmtransferencia.DBECpf_CnpjExit(Sender: TObject);
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

procedure TFrmtransferencia.EnomesExit(Sender: TObject);
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

procedure TFrmtransferencia.Label1Click(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
end;

procedure TFrmtransferencia.RBChequeClick(Sender: TObject);
begin
  Label1.Visible := true;
  DTPIni.Visible := true;
  Label23.Visible := true;
  DTPFim.Visible := true;
  Enomes.Visible := False;
end;

procedure TFrmtransferencia.GBBancoExit(Sender: TObject);
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

procedure TFrmtransferencia.GBChequeExit(Sender: TObject);
begin
  GBCheque.Visible := False;
end;

procedure TFrmtransferencia.EPlanocontabancoExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not AchaPlanoDeConta(450, 30, 'S', EPlanocontabanco.Text, '') then Begin
      EPlanocontabanco.SetFocus;
      exit;
    end;
    EPlanocontabanco.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    CDSParcelascontabanco.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
  end;
end;

procedure TFrmtransferencia.filtrabaixa;
var
  vartipo : string;
Begin
  ZQRecBai.SQL.Clear;
  if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    ZQRecBai.SQL.Add('select * from recebimento where (recpag='+quotedstr(varrecpag)+') and (saldo>0 and  not (saldo is null)) and (cliente ='+QuotedStr(Ecodnome.Text)+')');
    if not empty(EcodAdversa.text) Then Begin
      ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
    ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
  end
  else if RBCheque.Checked then Begin
    vartipo := '';
    DM_TAbelas.ZQTipodoc.First;
    while not DM_Tabelas.ZQTipodoc.Eof do begin
      if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
         vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
      DM_Tabelas.ZQTipodoc.Next;
    end;
    vartipo := copy(vartipo,1,length(vartipo)-1);
    ZQRecBai.SQL.Add('select * from recebimento where tipdoc in ('+vartipo+') and (recpag='+quotedstr(varrecpag)+') and (saldo>0 and  not (saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
    ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
    ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
    ZQRecBai.ParamByName('dt5').AsDate:=DTPIni.Date;
    ZQRecBai.ParamByName('dt6').AsDate:=DTPFim.Date;
  end;
  ZQRecBai.Open;
  XEEmpresend.Text := ZQRecBai.FieldByName('nomeempreend').AsString;
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  if RBParticipante.Checked Then
    while not ZQRecBai.Eof do begin
      ZQRecBai.Edit;
      if ZQRecBai.FieldByName('Dt_Vencimento').Value<=date Then
        ZQRecBai.FieldByName('marcar').Value := '0'
      else
        ZQRecBai.FieldByName('marcar').Value := '1';
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

procedure TFrmtransferencia.GBContraExit(Sender: TObject);
begin
  if not RBPonteiro.Checked Then
    Filtrabaixa;
end;

procedure TFrmtransferencia.RBParticipanteClick(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
  Enomes.SetFocus;
end;

procedure TFrmtransferencia.RBPonteiroClick(Sender: TObject);
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

procedure TFrmtransferencia.XDBNumEdit1Enter(Sender: TObject);
begin
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;

end;

procedure TFrmtransferencia.JDEntradaExit(Sender: TObject);
begin
  if not (DXBFechar.Focused or GBContra.Focused or CBSaldoCorrigido.Focused or RBPonteiro.Focused or RBCheque.Focused or RBParticipante.Focused ) Then Begin
    Atualiza_tela;
    DBGBaixando.SetFocus;
  end;
end;

procedure TFrmtransferencia.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  dm_tabelas.ZQLoteamento.close;
  JDEntrada.Enabled := False;
  JDBaixa.Enabled   := False; 
  ZQRecBai.Close;
  ZQRecebBxTemp.Close;
end;

procedure TFrmtransferencia.ZQRecBaiCalcFields(DataSet: TDataSet);
var
  vardI, vardT : Tdate;
  varsal : Double;
begin
    if (ZQRecBai.FieldByName('saldo').Value>0) and (ZQRecBai.FieldByName('recpag').Value='R') and (CBSaldoCorrigido.Checked)  Then Begin
    varsal := ZQRecBai.FieldByName('Valor').value;
    ZQRecebBxTemp.First;
    ZQRecebBxTemp.Refresh;
    vardI := ZQRecBai.FieldByName('Dt_Vencimento').Value;
    vardT := strtodate(Datetostr(FrmRecebBaixa.JDEntrada.Date));
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

procedure TFrmtransferencia.CDSParcelasCalcFields(DataSet: TDataSet);
begin
  CDSParcelasPercent.Value := CDSParcelasVrParc.Value*100/XNERecebido.Value;
//  CDSParcelasjuros.Value := XNEJuros.Value*CDSParcelasPercent.Value/100;
end;

procedure TFrmtransferencia.ZQRatearCalcFields(DataSet: TDataSet);
begin
    ZQRatear.FieldByName('vr_base').Value := ZQRatear.FieldByName('vr_rec').Value * ZQRatear.FieldByName('percent_usado').Value / 100;
    ZQRatear.FieldByName('vr_divi').Value := ZQRatear.FieldByName('vr_base').Value * ZQRatear.FieldByName('percentual').Value /100;

end;

procedure TFrmtransferencia.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').Value = 'R' Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrmtransferencia.CDS_divididoCalcFields(DataSet: TDataSet);
begin
  CDS_divididodescrcontabil2.Value := CDS_divididocontabil.Text+' - '+CDS_divididodescrcontabil.Value;
end;

procedure TFrmtransferencia.CBPeriodoEntExit(Sender: TObject);
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
  varvenci :=  JDEntrada.Date;
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
    if varvezes = 1 Then
       CDSParcelasVenci.Value := JDEntrada.Date
    else Begin
    if vartipo = 'S' then
        CDSParcelasVenci.Value := IncMonth(JDEntrada.Date,(varvezes*vardias)-(vardias))
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

procedure TFrmtransferencia.EadversaExit(Sender: TObject);
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
  end;
end;

procedure TFrmtransferencia.EAdversaNovaExit(Sender: TObject);
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

procedure TFrmtransferencia.CBTipobaixaExit(Sender: TObject);
begin
  JDEntrada.SetFocus;
end;

procedure TFrmtransferencia.DXBPesquisarClick(Sender: TObject);
begin
  FrmPesqRecebimento_bx.Top := FrmRecebBaixa.Top+100;
  FrmPesqRecebimento_bx.Left := FrmRecebBaixa.Left;
  FrmPesqRecebimento_bx.showmodal;


  JDEntrada.Date := Date;
  JDBaixa.Date   := Date;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Date := date;
  DTPFim.Date := date;
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


  CBPeriodoEnt.ItemIndex := 0;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True; 
  Atualiza_tela;


  DBGBaixando.SetFocus;
end;


procedure TFrmtransferencia.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmtransferencia, 'ZQRecBai', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaisaldocorrig', 'saldocorrig', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaidocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaicliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaisaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaimarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBairefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaicontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaicustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBairecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBainumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBainomeadversa', 'nomeadversa', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaivenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBainumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBainomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaiSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecBai', 'ZQRecBaisomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmtransferencia, 'ZQRecebBxTemp', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTemprefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTemppercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTemprefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTemporigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTemprecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRecebBxTemp', 'ZQRecebBxTempsubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmtransferencia, 'ZQRatear', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearvr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRateartipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearordem_1', 'ordem_1', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearvr_base', 'vr_base', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQRatear', 'ZQRatearvr_divi', 'vr_divi', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmtransferencia, 'ZQItensRecibo', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciborefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciborefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciborecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensRecibonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQItensRecibo', 'ZQItensReciboSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmtransferencia, 'ZQDiv_usado', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQDiv_usado', 'ZQDiv_usadoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQDiv_usado', 'ZQDiv_usadoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQDiv_usado', 'ZQDiv_usadopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQDiv_usado', 'ZQDiv_usadoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQDiv_usado', 'ZQDiv_usadotipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQDiv_usado', 'ZQDiv_usadoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmtransferencia, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmtransferencia, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.


