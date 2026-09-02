unit LimpaRemessa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, XEdit, Mask, XDate, ExtCtrls, XBanner,
  XLabel3D,   dxButton, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, JvExControls, JvSpecialProgress, Grids, DBGrids, DBClient,
  DBCtrls, TFlatCheckBoxUnit, ActnList, IniFiles, XNum;

type
  TFrmlimparemessa = class(TForm)
    RGOrdem: TRadioGroup;
    GroupBox1: TGroupBox;
    XDEEntradaInicio: TXDateEdit;
    XDEEntradaFinal: TXDateEdit;
    Label1: TLabel;
    Label2: TLabel;
    XLabel3D1: TXLabel3D;
    RGFormato: TRadioGroup;
    DXBFechar: TdxButton;
    DXBImprimir: TdxButton;
    DS_TempReceber: TDataSource;
    ZQTempReceber: TZQuery;
    GroupBox5: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    XDEVencimentoInicio: TXDateEdit;
    XDEVencimentoFinal: TXDateEdit;
    JSP1: TJvSpecialProgress;
    CBZebrar: TCheckBox;
    GroupBox6: TGroupBox;
    GroupBox7: TGroupBox;
    Label7: TLabel;
    Label10: TLabel;
    XDEBaixasinicio: TXDateEdit;
    XDEBaixasfinal: TXDateEdit;
    CBSaldo: TComboBox;
    DS_Baixa: TDataSource;
    ZQBaixa: TZQuery;
    GroupBox8: TGroupBox;
    CBVerbaixa: TComboBox;
    DS_ResumoReceb: TDataSource;
    ZQResumoReceb: TZQuery;
    DS_SomaBaixa: TDataSource;
    ZQSomaBaixa: TZQuery;

























    DS_AgruTipo: TDataSource;
    ZQAgruTipo: TZQuery;


















    RGContas: TRadioGroup;
    CBObsLanc: TCheckBox;
    CBOBSBaix: TCheckBox;
    DS_Temp2Cl: TDataSource;
    ZQTemp2Cl: TZQuery;
    DS_Cli: TDataSource;
    CDSCli: TClientDataSet;
    CDSCliCCli: TIntegerField;
    CDSCliRcli: TStringField;
    DS_Adversatemp: TDataSource;
    ZQAdvsersatemp: TZQuery;
    DS_Adversatemp2: TDataSource;
    CDSAdversaTemp2: TClientDataSet;
    CDSAdversaTemp2CodAdversa: TIntegerField;
    CDSAdversaTemp2nomeadversa: TStringField;














    DS_Plano: TDataSource;
    CDSPlano: TClientDataSet;
    CDSPlanomascara: TStringField;
    CDSPlanodescricao: TStringField;
    CDSPlanoCrDe: TStringField;
    CDSPlanosoma: TFloatField;
    CDSPlanoclassificacao: TStringField;
    CDSPlanocodigo: TIntegerField;
    DS_Detal: TDataSource;
    CDSDetal: TClientDataSet;
    CDSDetalmascara: TStringField;
    CDSDetaldata: TDateField;
    CDSDetaldetalhe: TMemoField;
    CDSDetalvalor: TFloatField;
    GBCliente: TGroupBox;
    XBanner2: TXBanner;
    ECliente: TEdit;
    DBGrid3: TDBGrid;
    GBAdversa: TGroupBox;
    XBanner10: TXBanner;
    DBGrid4: TDBGrid;
    DS_Rateio: TDataSource;
    ZQRateio: TZQuery;
    XBanner13: TXBanner;
    XBanner1: TXBanner;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    XBanner5: TXBanner;
    XBanner6: TXBanner;
    XBanner14: TXBanner;
    XBanner15: TXBanner;






























































































































































































































































    CBDesconsidera: TCheckBox;








































    GroupBox10: TGroupBox;
    XBanner16: TXBanner;
    CBEntrada: TFlatCheckBox;
    CBParcela: TFlatCheckBox;
    GroupBox11: TGroupBox;
    XBanner17: TXBanner;
    CLBEmpree: TCheckListBox;
    GroupBox4: TGroupBox;
    XBanner7: TXBanner;
    CLBDoc: TCheckListBox;
    DS_Empree: TDataSource;
    ZQEmpree: TZQuery;



















    DS_Venda: TDataSource;
    ZQVenda: TZQuery;

























    GroupBox2: TGroupBox;
    XBanner8: TXBanner;
    XBanner9: TXBanner;
    EQuadra: TEdit;
    DBGrid1: TDBGrid;
    Elote: TEdit;
    Label3: TLabel;
    DS_QuadraLote: TDataSource;
    CDSQuadraLote: TClientDataSet;
    CDSQuadraLoteQuadra: TStringField;
    CDSQuadraLotelote: TStringField;
    CDSQuadraLoteidlote: TIntegerField;
    CDSQuadraLoteidloteamen: TIntegerField;
    DS_Ven1: TDataSource;
    ZQVen1: TZQuery;

























    DS_ROReceb: TDataSource;
    ZROQReceb: TZReadOnlyQuery;














































    DS_Tipodoc: TDataSource;
    CDSTipoDoc: TClientDataSet;
    CDSTipoDocTipoDoc: TStringField;
    CDSTipoDocdescDoc: TStringField;
    CDSTipoDocdadosch: TStringField;
    DS_ROReceb1: TDataSource;
    ZROQReceb1: TZReadOnlyQuery;

















































    CDSTipoDocsomatit: TFloatField;
    CDSTipoDocsomapago: TFloatField;
    CDSTipoDocsomajuros: TFloatField;



    DS_Ch: TDataSource;
    ZROQCh: TZReadOnlyQuery;





















































    GroupBox3: TGroupBox;
    XBanner11: TXBanner;
    Label4: TLabel;
    Label8: TLabel;
    XDERefInicio: TXDateEdit;
    XDERefFinal: TXDateEdit;
    DataQReceb1: TDataSource;
    QReceb1: TClientDataSet;
    QReceb1idreceb_baixa: TLargeintField;
    QReceb1Dt_rec: TDateField;
    QReceb1Vr_rec: TFloatField;
    QReceb1TipDoc: TStringField;
    QReceb1Docum: TStringField;
    QReceb1Juros_Vr: TFloatField;
    QReceb1Desc_Vr: TFloatField;
    QReceb1RefBaixa: TIntegerField;
    QReceb1obsebx: TMemoField;
    QReceb1Codcontabil: TIntegerField;
    QReceb1dataref: TDateField;
    QReceb1vencimento: TDateField;
    QReceb1idrecbxhist: TLargeintField;
    QReceb1refer: TIntegerField;
    QReceb1idrecib: TIntegerField;
    QReceb1valor: TFloatField;
    QReceb1descon: TFloatField;
    QReceb1juros: TFloatField;
    QReceb1percent_usado: TFloatField;
    QReceb1data: TDateField;
    QReceb1idrecebimento: TLargeintField;
    QReceb1documento: TStringField;
    QReceb1cliente: TIntegerField;
    QReceb1usuario: TIntegerField;
    QReceb1Dt_Entrada: TDateField;
    QReceb1Dt_Vencimento: TDateField;
    QReceb1Valor_1: TFloatField;
    QReceb1Observ: TMemoField;
    QReceb1VrDoc: TFloatField;
    QReceb1ordem: TStringField;
    QReceb1TipDoc_1: TStringField;
    QReceb1saldo: TFloatField;
    QReceb1marcar: TStringField;
    QReceb1RefBaixa_1: TIntegerField;
    QReceb1refvinda: TIntegerField;
    QReceb1contabil: TIntegerField;
    QReceb1empresa: TIntegerField;
    QReceb1custodaparcela: TFloatField;
    QReceb1origem: TStringField;
    QReceb1adversa: TIntegerField;
    QReceb1recpag: TStringField;
    QReceb1numordem: TIntegerField;
    QReceb1idloteamento: TIntegerField;
    QReceb1venda_idvenda: TIntegerField;
    QReceb1quadralote: TStringField;
    QReceb1numboleto: TStringField;
    QReceb1nomeadversa: TStringField;
    QReceb1nomecliente: TStringField;
    QReceb1nomeempree: TStringField;

    dxbx: TRadioGroup;















    CDS_MarcaTit: TClientDataSet;
    CDS_MarcaTitmarca: TStringField;
    CDS_MarcaTitIdreceb: TIntegerField;
    CDS_MarcaTitidparti: TIntegerField;
    CDS_MarcaTitDT_Entrada: TDateField;
    CDS_MarcaTitDt_Vencimento: TDateField;
    CDS_MarcaTitValor: TFloatField;
    CDS_MarcaTitordem: TStringField;
    CDS_MarcaTitnome_parte: TStringField;
    CDS_MarcaTitende_cob: TStringField;
    CDS_MarcaTitcep_cob: TStringField;
    CDS_MarcaTitdoc1: TStringField;
    CDS_MarcaTitTipDoc: TStringField;
    CDS_MarcaTitnumordem: TIntegerField;
    CDS_MarcaTitcliente: TIntegerField;
    CDS_MarcaTitidloteam: TIntegerField;
    DS_MarcaTit: TDataSource;

    CDS_MarcaTitquadralote: TStringField;
    CDS_MarcaTitnumboleto: TStringField;
    dxbSelecionar: TdxButton;
    stotais: TCheckBox;

    DataZQTempReceber2: TDataSource;
    ZQTempReceber2: TZQuery;
































    CDS_MarcaTitdigito_dif: TStringField;
    CDS_MarcaTitbairro: TStringField;
    CDS_MarcaTitdoc2: TStringField;
    CDS_MarcaTitcidade: TLargeintField;
    CDS_MarcaTitnomecidade: TStringField;
    CDS_MarcaTituf: TStringField;















    Datamensal: TDataSource;
    ZQmensal: TZQuery;
    cbmensal: TCheckBox;












    ZQmensal2: TZQuery;
    Dzqmensal2: TDataSource;
    ActionList1: TActionList;














    XBanner12: TXBanner;
    GroupBox9: TGroupBox;
    XBanner18: TXBanner;
    XBanner19: TXBanner;
    Edit1: TEdit;
    DBGrid2: TDBGrid;
    DataZQRelatorio2: TDataSource;
    ZQRelatorio2: TZQuery;
    ZQRelatorio: TZQuery;














    DataZQRelatorio: TDataSource;














    ZqresumoE: TZReadOnlyQuery;
    DS_resumoE: TDataSource;
























































    ZqresumoP: TZReadOnlyQuery;
    DS_ZqresumoP: TDataSource;
























































    xnparcelas_dentro: TXNumEdit;
    xnentradas_dentro: TXNumEdit;
    xnentradas_fora: TXNumEdit;
    xnparcelas_fora: TXNumEdit;
    geral_recebido: TXNumEdit;
    xresumo: TCheckBox;
    entrada_atrazada: TXNumEdit;
    entrada_adiantada: TXNumEdit;
    parcela_atrazada: TXNumEdit;
    parcela_adiantada: TXNumEdit;
    CBObsBai: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DXBImprimirClick(Sender: TObject);
    procedure XDEEntradaInicioExit(Sender: TObject);
    procedure XDEVencimentoInicioExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GBEscClienteExit(Sender: TObject);
    procedure ZQBaixaCalcFields(DataSet: TDataSet);
    procedure EClienteExit(Sender: TObject);
    procedure DBGrid3KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid4KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RGContasClick(Sender: TObject);
    procedure RGFormatoClick(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure XDERefInicioExit(Sender: TObject);
    procedure dxbSelecionarClick(Sender: TObject);
    procedure CBSaldoExit(Sender: TObject);
    procedure Edit1Exit(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frmlimparemessa: TFrmlimparemessa;
  varschemata : string;
implementation

uses Tabelas, Funcoes, Principal, RelReceb02, Cad_Recebimento,
  RelRecebGerencial, CobrancaBancaria, Recebi03, RelReceb02_total,
  RelReceb02_mensal, RelReceb02_mensal_ab, uRuntimeFields;

{$R *.dfm}

procedure TFrmlimparemessa.FormShow(Sender: TObject);
Var
  VarPath : string;
  ArqIni2 : tIniFile;
begin
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni2 := tIniFile.Create(varpath+'siai.Ini');
  varschemata := ArqIni2.ReadString('BANCO_DE_DADOS','SCHEMATA', varschemata );


  if (RGFormato.ItemIndex = 4) or (RGFormato.ItemIndex = 2) Then
    CBDesconsidera.Visible := True
  else
    CBDesconsidera.Visible := False;
  ZQTemp2Cl.Open;
  CDSQuadraLote.Close;
  CDSQuadraLote.CreateDataSet;
  CDSCli.Close;
  CDSCli.CreateDataSet;
  CDSAdversaTemp2.Close;
  CDSAdversaTemp2.CreateDataSet;
  CLBDoc.Sorted := False;
  DM_tabelas.ZQTipodoc.First;
  CLBEmpree.Clear;
  DM_Tabelas.ZQincorp_loteame.open;
  DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQLoteamento.First;
  while not DM_TAbelas.ZQLoteamento.Eof do begin
    CLBEmpree.Items.Add(DM_Tabelas.ZQLoteamento.FieldByName('apelido').Value+' - '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text);
    DM_Tabelas.ZQLoteamento.Next;
  end;
  CLBDoc.Clear;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    CLBDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    CLBDoc.Checked[CLBDoc.Count-1] := True;
   DM_tabelas.ZQTipodoc.Next;
  end;
  CLBDoc.Sorted := True;
  RGContas.SetFocus;
end;

procedure TFrmlimparemessa.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmlimparemessa.DXBImprimirClick(Sender: TObject);
var
  VarEmpree, VarDoc, VarOrdem, vartitulo, Varaleatorio, vardevedor, Vrtexto,
  dentromes,forames,data,d,m,a,entrada,vcto,baixa,varmasc, vardocumento, varquadra, varsql,VarLogo : String;
  mes1,mes2,varcontador, x, y, varvezes : Integer;
const
  vartama : array[1..5] of integer = (1,3,6,10,15);

begin
  VarEmpree:= '';
  Varlogo:='';
  varcontador := 0;
  for x := 0 to CLBEmpree.Count-1 do begin
    if CLBEmpree.Checked[x] Then Begin
      inc(varcontador);
      for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
        if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
          break;
      VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
    end;
  end;
  if varcontador<>CLBEmpree.Count Then
    VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
  else
    VarEmpree := '';
  varlogo:=VarEmpree;
  if not empty(VarLogo) then
  begin
    if pos(',',varlogo)=0 then
    begin
      ZQEmpree.Open;
      ZQEmpree.Filtered:=false;
      ZQEmpree.Filter:='idloteamento='+varlogo;
      ZQEmpree.Filtered:=true;
    end
    else
    begin
      VarLogo:='';
      ZQEmpree.Close;
    end;
  end;
  varquadra := '';
  CDSQuadraLote.First;
  while not CDSQuadraLote.Eof do begin
    varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
    CDSQuadraLote.Next;
  end;
  varquadra := copy(varquadra,1,length(varquadra)-1);
  vardocumento := '';
  if CBEntrada.Checked <> CBParcela.Checked Then Begin
    if CBParcela.Checked Then
      vardocumento := ' and documento like '+quotedstr('%-P-%');
    if CBEntrada.Checked Then
      vardocumento := ' and documento like '+quotedstr('%-E-%');
  end;
  FrmRelReceb02.XNsomabaixa.Value:=0;
  FrmRelReceb02.XNNegociado.Value:=0;
  FrmRelReceb02.XNRenova.Value:=0;
  FrmRelReceb02.XNPrincipal.Value:=0;
  Varaleatorio := '';
  if CDSCli.Active Then Begin
    CDSCli.First;
    while not CDSCli.Eof do Begin
      Varaleatorio := Varaleatorio +quotedstr(CDSCliCcli.Text)+',';
      CDSCli.Next;
    end;
    Varaleatorio := copy(Varaleatorio,1,length(Varaleatorio)-1)
  end;
  vardevedor := '';
  if CDSAdversaTemp2.Active Then Begin
    CDSAdversaTemp2.First;
    while not CDSAdversaTemp2.Eof do Begin
      vardevedor := vardevedor +quotedstr(CDSAdversaTemp2CodAdversa.Text)+',';
      CDSAdversaTemp2.Next;
    end;
    vardevedor := copy(vardevedor,1,length(vardevedor)-1)
  end;
  if RGFormato.ItemIndex=5 then
  begin
    case RGOrdem.ItemIndex of
      0: VarOrdem := 'DT_Vencimento';
  //   1: VarOrdem := 'DT_Entrada';
      1: VArOrdem := 'nomeadversa';
      2: VArOrdem := 'quadralote';
      3: VArOrdem := 'TipDoc';
      4: VArOrdem := 'dataref';
    end;
  end
  else
  begin
    case RGOrdem.ItemIndex of
      0: VarOrdem := 'DT_Vencimento';
  //   1: VarOrdem := 'DT_Entrada';
      1: VArOrdem := 'nomeadversa';
      2: VArOrdem := 'quadralote';
      3: VArOrdem := 'TipDoc';
    end;
  end;
  JSP1.Visible := True;
  VarDoc := '';
  varcontador := 0;
  JSP1.Maximum := CLBEmpree.Count-1;
  for x := 0 to CLBEmpree.Count-1 do begin
    JSP1.Position := x;
//    VarEmpree :='';
    if CLBEmpree.Checked[x] Then Begin
      inc(varcontador);
      for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
        if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
          break;
      VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
    end;
  end;
  if varcontador<>CLBEmpree.Count Then
    VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
  else
    VarEmpree := '';
  varcontador := 0;
  JSP1.Maximum := CLBDoc.Count-1;
  for x := 0 to CLBDoc.Count-1 do begin
    JSP1.Position := x;
    if CLBDoc.Checked[x] Then Begin
      inc(varcontador);
      VarDoc := VarDoc +quotedstr(CLBDoc.Items.Strings[x])+',';
    end;
  end;
  VarDoc := copy(varDoc,1,length(varDoc)-1);

  if RGFormato.ItemIndex = 5 Then
  Begin

   if cbmensal.Checked=true then
   begin
    ZQmensal2.SQL.Clear;
//    ZQmensal2.SQL.Add('DROP VIEW IF EXISTS `'+varschemata+'`.`relmensal`;');
    ZQmensal2.SQL.Add('CREATE OR REPLACE VIEW `'+varschemata+'`.`relmensal` AS ');
    ZQmensal2.SQL.Add(' ( select month(rb.dataref) as Mes, case month(rb.dataref)');
    ZQmensal2.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Março''');
    ZQmensal2.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQmensal2.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQmensal2.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
    ZQmensal2.SQL.Add(' end AS MESDESC, CASE when (rb.Vr_rec > re.valor) then re.valor else rb.Vr_rec end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.valor, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr, ');
    ZQmensal2.SQL.Add(' year(rb.dataref) as Ano from receb_baixa as rb join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento where RB.TIPDOC in ('+varDoc+')');
    if not empty(VarEmpree) then
       ZQmensal2.SQL.Add(' and re.idloteamento in ('+VarEmpree+')');
    if RGContas.ItemIndex = 0 Then
      ZQmensal2.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZQmensal2.SQL.Add(' and re.recpag ='+quotedstr('P'));
    if not empty(vardocumento) Then
      ZQmensal2.SQL.Add(vardocumento);
    if not empty(varquadra) then
      ZQmensal2.SQL.Add(' and re.quadralote in ('+varquadra+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and re.DT_Entrada between :dt1 and :dt2');
      ZQmensal2.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQmensal2.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' até '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQmensal2.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQmensal2.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' até '+XDEVencimentoFinal.DateText;
    end;

    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZQmensal2.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZQmensal2.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' até '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZQmensal2.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZQmensal2.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' até '+XDERefFinal.DateText;
    end;
//    ZQmensal2.SQL.Add(' group by mes,ano order by ano,mes');
    ZQmensal2.SQL.Add(' order by ano,mes )');
//    ZQmensal2.Open;
  //  ZQmensal2.First;
    ZQmensal2.ExecSQL;
    ZQrelatorio.close;
    ZQrelatorio.SQL.clear;
    ZQrelatorio.SQL.Add('select *,sum(vr_rel) as valorpago, sum(juros_vr+vr_rel) as valorpagojuros from relmensal group by mes,ano order by ano,mes');
    ZQrelatorio.open;

    if FrmRelReceb02_mensal=nil then
      FrmRelReceb02_mensal:=TFrmRelReceb02_mensal.Create(Application);
    if not empty(vrtexto) Then
      FrmRelReceb02_mensal.RLLabel2.Caption := 'Período de : '+vrtexto;
    FrmRelReceb02_mensal.RLLabel1.Caption := 'Contas Recebidas';
    if RGContas.ItemIndex = 1 Then

    FrmRelReceb02_mensal.RLLabel1.Caption := 'Contas Pagas';
    FrmRelReceb02_mensal.RLLabel17.Caption:='Vr. Saldo Baixado  ';
    FrmRelReceb02_mensal.RLDBResult2.Text:='Saldo Geral Baixado:    ';
    FrmRelReceb02_mensal.RLDBResult1.Text:='Saldo Total Baixado:    ';
    FrmRelReceb02_mensal.RLLabel11.Caption:='Vr. Saldo Baixado/Variação';
//    FrmRelReceb02_mensal.RLDBResult3.Text:='Saldo Geral Baixado/Variação:    ';
//    FrmRelReceb02_mensal.RLDBResult4.Text:='Saldo Total Baixado/Variação:    ';

    FrmRelReceb02_mensal.RLReport1.DataSource:=DataZQrelatorio;
    FrmRelReceb02_mensal.RLReport1.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBText2.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBText8.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBText3.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBResult1.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBResult2.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBText4.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBResult3.DataSource:=FrmRelRecebimento.DataZQrelatorio;
    FrmRelReceb02_mensal.RLDBResult4.DataSource:=FrmRelRecebimento.DataZQrelatorio;

    FrmRelReceb02_mensal.RLReport1.Preview;
    ZQEmpree.Close;
    ZQmensal2.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;


    FRM_Recebi03.XNumEdit1.Value := 0;
    FRM_Recebi03.XNumEdit2.Value := 0;
    FRM_Recebi03.XNumEdit3.Value := 0;
    CDSTipoDoc.Close;
    CDSTipoDoc.CreateDataSet;
    ZROQReceb.SQL.Clear;
    Vrtexto:='';
    ZROQReceb.SQL.Add('SELECT *, sum(Re.valor) as somatit, sum(vr_rec) as somapago, sum(juros_vr) as somajuros FROM receb_baixa as RB');
    ZROQReceb.SQL.Add(' join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
    ZROQReceb.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');
    if not empty(vardevedor) then
      ZROQReceb.SQL.Add(' and adversa in ('+vardevedor+')');

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZROQReceb.SQL.Add(' and quadralote in ('+varquadra+')');
    if RGContas.ItemIndex = 0 Then
      ZROQReceb.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZROQReceb.SQL.Add(' and re.recpag ='+quotedstr('P'));
    if not empty(Varaleatorio) then
      ZROQReceb.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      ZROQReceb.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZROQReceb.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZROQReceb.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' até '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZROQReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZROQReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' até '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZROQReceb.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZROQReceb.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' até '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZROQReceb.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZROQReceb.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' até '+XDERefFinal.DateText;
    end;
    ZROQReceb.SQL.Add(' group by RB.tipdoc ');
    ZROQReceb.Open;

    if not empty(vrtexto) Then
      FRM_Recebi03.RLLabel2.Caption := 'Período de : '+vrtexto;
    FRM_Recebi03.XNumEdit1.Value := 0;
    while not ZROQReceb.Eof do begin
      CDSTipoDoc.Insert;
      CDSTipoDocTipoDoc.Value := ZROQReceb.FieldByName('TipDoc').Value;
      CDSTipoDocsomatit.Value := ZROQReceb.FieldByName('somatit').Value;
      CDSTipoDocsomapago.Value := ZROQReceb.FieldByName('somapago').Value;
      CDSTipoDocsomajuros.Value := ZROQReceb.FieldByName('somajuros').Value;
      CDSTipoDoc.Post;
      FRM_Recebi03.XNumEdit1.Value := FRM_Recebi03.XNumEdit1.Value+ZROQReceb.FieldByName('somatit').Value;
//      FRM_Recebi03. :=0;
      FRM_Recebi03.XNumEdit2.Value := FRM_Recebi03.XNumEdit2.Value+ZROQReceb.FieldByName('somapago').Value;
      FRM_Recebi03.XNumEdit3.Value := FRM_Recebi03.XNumEdit3.Value+ZROQReceb.FieldByName('somajuros').Value;
      ZROQReceb.Next;
    end;
    ZROQReceb1.SQL.Clear;
    ZROQReceb1.SQL.Add('SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
    ZROQReceb1.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');


    if CBSaldo.ItemIndex = 2 Then
      ZROQReceb1.SQL.Add(' and (vr_rec is null or vr_ab > 0)');
    if CBSaldo.ItemIndex = 1 Then
      ZROQReceb1.SQL.Add(' and (vr_ab is null or vr_rec > 0)');


    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZROQReceb1.SQL.Add(' and quadralote in ('+varquadra+')');

    if RGContas.ItemIndex = 0 Then
      ZROQReceb1.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZROQReceb1.SQL.Add(' and re.recpag ='+quotedstr('P'));

    if not empty(vardevedor) then
      ZROQReceb1.SQL.Add(' and adversa in ('+vardevedor+')');
    if not empty(Varaleatorio) then
      ZROQReceb1.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      ZROQReceb1.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZROQReceb1.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZROQReceb1.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZROQReceb1.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZROQReceb1.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZROQReceb1.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZROQReceb1.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZROQReceb1.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZROQReceb1.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    end;

    // coloquei aqui 24/01/2011   para nao sair duplicado as parcelas
//    ZROQReceb1.SQL.Add(' group by RE.ordem ');

    if VArOrdem='TipDoc' then
       ZROQReceb1.SQL.Add(' ORDER BY RB.'+VarOrdem)
    else
       ZROQReceb1.SQL.Add(' ORDER BY '+VarOrdem);
    ZROQReceb1.Open;

   //resumo

    ZqresumoE.SQL.Clear;
    ZqresumoE.SQL.Add('SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
    ZqresumoE.SQL.Add(' where RE.documento like ''%-E-%'' and RB.TIPDOC in ('+varDoc+')');


    if CBSaldo.ItemIndex = 2 Then
      ZqresumoE.SQL.Add(' and (vr_rec is null or vr_ab > 0)');
    if CBSaldo.ItemIndex = 1 Then
      ZqresumoE.SQL.Add(' and (vr_ab is null or vr_rec > 0)');

    if RGContas.ItemIndex = 0 Then
      ZqresumoE.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZqresumoE.SQL.Add(' and re.recpag ='+quotedstr('P'));

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZqresumoE.SQL.Add(' and quadralote in ('+varquadra+')');


    if not empty(vardevedor) then
      ZqresumoE.SQL.Add(' and adversa in ('+vardevedor+')');
    if not empty(Varaleatorio) then
      ZqresumoE.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      ZqresumoE.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZqresumoE.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZqresumoE.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZqresumoE.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZqresumoE.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoE.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoE.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZqresumoE.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZqresumoE.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZqresumoE.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZqresumoE.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZqresumoE.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZqresumoE.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    end;

    // coloquei aqui 24/01/2011   para nao sair duplicado as parcelas
//    ZROQReceb1.SQL.Add(' group by RE.ordem ');

    if VArOrdem='TipDoc' then
       ZqresumoE.SQL.Add(' ORDER BY RB.'+VarOrdem)
    else
       ZqresumoE.SQL.Add(' ORDER BY '+VarOrdem);
    ZqresumoE.Open;

    ZqresumoE.First;
    xnentradas_fora.Value:=0;
    xnentradas_dentro.Value:=0;
    entrada_adiantada.Value:=0;
    entrada_atrazada.Value:=0;
    JSP1.Visible := True;
    JSP1.Position:=0;
    JSP1.Maximum:=ZqresumoE.RecordCount;
    while not ZqresumoE.eof do begin
      JSP1.Position:=ZqresumoE.RecNo;
      if mes(ZqresumoE.FieldByName('dataref').Value)=mes(ZqresumoE.FieldByName('vencimento').Value) then
      begin
        if ZqresumoE.FieldByName('Vr_rec').Value>ZqresumoE.FieldByName('Valor_1').Value then
        begin
           xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoE.FieldByName('Valor_1').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
        end
        else
        begin
           if ZqresumoE.FieldByName('Vr_rec').Value>0 then
           begin
             xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoE.FieldByName('Vr_rec').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
           end
           else
           begin
             xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoE.FieldByName('VR_AB').Value;
           end;
        end;
      end
      else
      begin
        mes1:=strtoint(mes(ZqresumoE.FieldByName('dataref').Value));
        mes2:=strtoint(mes(ZqresumoE.FieldByName('vencimento').Value));
        if mes1>mes2 then
        begin
          if ZqresumoE.FieldByName('Vr_rec').Value>ZqresumoE.FieldByName('Valor_1').Value then
          begin
             xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('Valor_1').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
             entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoE.FieldByName('Valor_1').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
          end
          else
          begin
             if ZqresumoE.FieldByName('Vr_rec').Value>0 then
             begin
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('Vr_rec').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
               entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoE.FieldByName('Vr_rec').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
             end
             else
             begin
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('VR_AB').Value;
               entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoE.FieldByName('VR_AB').Value;
             end;
          end;
        end
        else
        begin
          if ZqresumoE.FieldByName('Vr_rec').Value>ZqresumoE.FieldByName('Valor_1').Value then
          begin
             xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('Valor_1').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
             entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoE.FieldByName('Valor_1').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
          end
          else
          begin
             if ZqresumoE.FieldByName('Vr_rec').Value>0 then
             begin
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('Vr_rec').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
               entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoE.FieldByName('Vr_rec').Value+ZqresumoE.FieldByName('Juros_Vr').Value;
             end
             else
             begin
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('VR_AB').Value;
               entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoE.FieldByName('VR_AB').Value;
             end;
          end;
        end;
      end;
      ZqresumoE.Next;
    end;


    ZqresumoP.SQL.Clear;
    ZqresumoP.SQL.Add('SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
    ZqresumoP.SQL.Add(' where RE.documento like ''%-P-%''  and RB.TIPDOC in ('+varDoc+')');


    if CBSaldo.ItemIndex = 2 Then
      ZqresumoP.SQL.Add(' and (vr_rec is null or vr_ab > 0)');
    if CBSaldo.ItemIndex = 1 Then
      ZqresumoP.SQL.Add(' and (vr_ab is null or vr_rec > 0)');

    if RGContas.ItemIndex = 0 Then
      ZqresumoP.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZqresumoP.SQL.Add(' and re.recpag ='+quotedstr('P'));


    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZqresumoP.SQL.Add(' and quadralote in ('+varquadra+')');


    if not empty(vardevedor) then
      ZqresumoP.SQL.Add(' and adversa in ('+vardevedor+')');
    if not empty(Varaleatorio) then
      ZqresumoP.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      ZqresumoP.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZqresumoP.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZqresumoP.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZqresumoP.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZqresumoP.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoP.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoP.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZqresumoP.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZqresumoP.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZqresumoP.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZqresumoP.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZqresumoP.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZqresumoP.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    end;

    // coloquei aqui 24/01/2011   para nao sair duplicado as parcelas
//    ZROQReceb1.SQL.Add(' group by RE.ordem ');

    if VArOrdem='TipDoc' then
       ZqresumoP.SQL.Add(' ORDER BY RB.'+VarOrdem)
    else
       ZqresumoP.SQL.Add(' ORDER BY '+VarOrdem);
    ZqresumoP.Open;

    JSP1.Position:=0;
    JSP1.Maximum:=ZqresumoP.RecordCount;
    ZqresumoP.First;
    xnparcelas_dentro.Value:=0;
    xnparcelas_fora.Value:=0;
    parcela_atrazada.Value:=0;
    parcela_adiantada.Value:=0;
    
    while not ZqresumoP.eof do begin
      JSP1.Position:=ZqresumoP.RecNo;
      if mes(ZqresumoP.FieldByName('dataref').Value)=mes(ZqresumoP.FieldByName('vencimento').Value) then
      begin
        if ZqresumoP.FieldByName('Vr_rec').Value>ZqresumoP.FieldByName('Valor_1').Value then
        begin
           xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoP.FieldByName('Valor_1').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
        end
        else
        begin
           if ZqresumoP.FieldByName('Vr_rec').Value>0 then
           begin
             xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoP.FieldByName('Vr_rec').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
           end
           else
           begin
             xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoP.FieldByName('VR_AB').Value;
           end;
        end;
      end
      else
      begin
        mes1:=strtoint(mes(ZqresumoP.FieldByName('dataref').Value));
        mes2:=strtoint(mes(ZqresumoP.FieldByName('vencimento').Value));
        if mes1>mes2 then
        begin
          if ZqresumoP.FieldByName('Vr_rec').Value>ZqresumoP.FieldByName('Valor_1').Value then
          begin
             xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('Valor_1').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
             parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoP.FieldByName('Valor_1').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
          end
          else
          begin
             if ZqresumoP.FieldByName('Vr_rec').Value>0 then
             begin
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('Vr_rec').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
               parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoP.FieldByName('Vr_rec').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
             end
             else
             begin
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('VR_AB').Value;
               parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoP.FieldByName('VR_AB').Value;
             end;
          end;
        end
        else
        begin
          if ZqresumoP.FieldByName('Vr_rec').Value>ZqresumoP.FieldByName('Valor_1').Value then
          begin
             xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('Valor_1').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
             parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoP.FieldByName('Valor_1').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
          end
          else
          begin
             if ZqresumoP.FieldByName('Vr_rec').Value>0 then
             begin
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('Vr_rec').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
               parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoP.FieldByName('Vr_rec').Value+ZqresumoP.FieldByName('Juros_Vr').Value;
             end
             else
             begin
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('VR_AB').Value;
               parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoP.FieldByName('VR_AB').Value;
             end;
          end;
        end;
      end;
      ZqresumoP.Next;
    end;
    JSP1.Position:=0;
    JSP1.Visible := false;
    ZqresumoP.close;
    ZqresumoE.close;
    ZROQCh.Open;
    if FRM_Recebi03=nil then
      FRM_Recebi03:=TFRM_Recebi03.Create(Application);

    FRM_Recebi03.RLLabel41.Caption:=Transform(xnentradas_dentro.value,'###,###,##0.00');
    FRM_Recebi03.RLLabel42.Caption:=Transform(xnentradas_fora.value,'###,###,##0.00');
    FRM_Recebi03.RLLabel47.Caption:=Transform(entrada_atrazada.Value,'###,###,##0.00');
    FRM_Recebi03.RLLabel49.Caption:=Transform(entrada_adiantada.Value,'###,###,##0.00');

    FRM_Recebi03.RLLabel44.Caption:=Transform(xnparcelas_dentro.value,'###,###,##0.00');
    FRM_Recebi03.RLLabel43.Caption:=Transform(xnparcelas_fora.value,'###,###,##0.00');
    FRM_Recebi03.RLLabel53.Caption:=Transform(parcela_atrazada.Value,'###,###,##0.00');
    FRM_Recebi03.RLLabel52.Caption:=Transform(parcela_adiantada.Value,'###,###,##0.00');
    if CBObsBai.Checked=true then
    begin
      FRM_Recebi03.RLLabel19.Caption:='Observações';
      FRM_Recebi03.RLDBText13.Visible:=false;
      FRM_Recebi03.RLDBMemo2.Visible:=true;
    end
    else
    begin
      FRM_Recebi03.RLLabel19.Caption:='Loteamento';
      FRM_Recebi03.RLDBText13.Visible:=true;
      FRM_Recebi03.RLDBMemo2.Visible:=false;
    end;
    geral_recebido.Value:=0;
    geral_recebido.Value:=xnparcelas_fora.value+xnparcelas_dentro.value+xnentradas_dentro.value+xnentradas_fora.value;
    xnparcelas_fora.value:=0;
    xnparcelas_dentro.value:=0;
    xnentradas_dentro.value:=0;
    xnentradas_fora.value:=0;
    entrada_adiantada.Value:=0;
    entrada_atrazada.Value:=0;
    parcela_atrazada.Value:=0;
    parcela_adiantada.Value:=0;

    if xresumo.Checked=true then
    begin
      FRM_Recebi03.RLLabel45.Visible:=true;
      FRM_Recebi03.RLLabel45.Caption:=Transform(geral_recebido.Value,'###,###,##0.00');
      FRM_Recebi03.RLBand2.Visible:=false;
      FRM_Recebi03.RLBand7.Visible:=false;
      FRM_Recebi03.RLSubDetail1.Visible:=false;
      FRM_Recebi03.RLImage1.Visible:=false;
      FRM_Recebi03.RLLabel30.Visible:=false;
    end
    else
    begin
      FRM_Recebi03.RLLabel45.Caption:='';
      FRM_Recebi03.RLLabel45.Visible:=false;
      FRM_Recebi03.RLBand2.Visible:=true;
      FRM_Recebi03.RLBand7.Visible:=true;
      FRM_Recebi03.RLSubDetail1.Visible:=true;
      FRM_Recebi03.RLImage1.Visible:=true;
      FRM_Recebi03.RLLabel30.Visible:=true;
    end;
    geral_recebido.Value:=0;

    if dxbx.ItemIndex=1 then
    begin
       FRM_Recebi03.RLLabel15.Visible:=false;
       FRM_Recebi03.RLDBText9.Visible:=false;
    end
    else
    begin
      FRM_Recebi03.RLLabel15.Visible:=true;
      FRM_Recebi03.RLDBText9.Visible:=true;
    end;
    FRM_Recebi03.RLReport1.Preview;
    ZQEmpree.Close;
    exit;
  end;
  if RGFormato.ItemIndex = 4 Then
  Begin
    DM_Tabelas.CDSEmpreendimento.Close;
    DM_Tabelas.CDSEmpreendimento.CreateDataSet;
    DM_Tabelas.ZQLoteamento.First;
    While not DM_Tabelas.ZQLoteamento.Eof do Begin
      while not DM_Tabelas.ZQincorp_loteame.Eof do begin
        if (pos(DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').Text,Varaleatorio)>0) or empty(Varaleatorio) Then Begin
          DM_Tabelas.CDSEmpreendimento.Insert;
          DM_TAbelas.CDSEmpreendimentoidordem.Value := DM_Tabelas.CDSEmpreendimento.RecordCount;
          DM_Tabelas.CDSEmpreendimentoidloteamento.Value := DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
          DM_Tabelas.CDSEmpreendimentocodcontabancaria.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('codcontabancaria').Value;
          DM_Tabelas.CDSEmpreendimentoidparticipante.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').Value;
          DM_Tabelas.CDSEmpreendimentodigito_dif.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('digito_dif').Value;
          DM_Tabelas.CDSEmpreendimentomarca.Value := '0';
          DM_Tabelas.CDSEmpreendimento.Post;
        end;
        DM_Tabelas.ZQincorp_loteame.Next;
      end;
      DM_TAbelas.ZQLoteamento.Next;
    end;

    DM_Tabelas.ZQRecebParticp.SQL.Clear;
    DM_Tabelas.ZQRecebParticp.SQL.Add('select * from recebimento join participante on adversa=idpaticipante');
    DM_Tabelas.ZQRecebParticp.SQL.Add(' where recpag = '+quotedstr('R'));
    if CBEntrada.Checked <> CBParcela.Checked Then Begin
      if CBParcela.Checked Then
        DM_Tabelas.ZQRecebParticp.SQL.Add(' and documento like '+quotedstr('%-P-%'));
    end;
    if (not empty(vardevedor)) and (RGContas.ItemIndex=0) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and adversa in ('+vardevedor+')');


    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and quadralote in ('+varquadra+')');



    if not empty(Varaleatorio) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if CBSaldo.ItemIndex = 1 Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and (saldo <valor or saldo is null)');
    if CBSaldo.ItemIndex = 2 Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and saldo > 0');
    if not empty(vardocumento) Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(vardocumento);
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and BX.DT_Rec between :dt5 and :dt6');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and BX.vencimento between :dt7 and :dt8');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    end;
    DM_Tabelas.ZQRecebParticp.SQL.Add(' and TipDoc in ('+varDoc+')');

    if CBDesconsidera.Checked Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and not exists (select * from remessa_receb where idrec=idrecebimento)');
      
    DM_Tabelas.ZQRecebParticp.Open;
//    showmessage(DM_Tabelas.ZQRecebParticp.SQL.Text);
    if DM_Tabelas.ZQRecebParticp.RecordCount>0 then
       JSP1.Maximum := DM_Tabelas.ZQRecebParticp.RecordCount;

{    DM_Tabelas.CDS_MarcaTit.Close;
    DM_TAbelas.CDS_MarcaTit.CreateDataSet;

    while not DM_Tabelas.ZQRecebParticp.Eof do begin
      JSP1.Position := DM_Tabelas.ZQRecebParticp.RecNo;
      DM_Tabelas.CDS_MarcaTit.Insert;
      DM_Tabelas.CDS_MarcaTitIdreceb.Value := DM_Tabelas.ZQRecebParticpidrecebimento.Value;
      DM_TAbelas.CDS_MarcaTitidparti.Value := DM_Tabelas.ZQRecebParticpadversa.Value;
      DM_TAbelas.CDS_MarcaTitcliente.Value := DM_Tabelas.ZQRecebParticpcliente.Value;
      DM_Tabelas.CDS_MarcaTitidloteam.Value := DM_Tabelas.ZQRecebParticpidloteamento.Value;
      DM_Tabelas.CDS_MarcaTitmarca.Value := '0';
      DM_Tabelas.CDS_MarcaTit.Post;
      DM_Tabelas.ZQRecebParticp.Next;
    end;}
    if CDS_MarcaTit.Active=false then
    begin
      CDS_MarcaTit.Close;
      CDS_MarcaTit.CreateDataSet;
    end;

    while not DM_Tabelas.ZQRecebParticp.Eof do
    begin
      JSP1.Position := DM_Tabelas.ZQRecebParticp.RecNo;
      CDS_MarcaTit.Insert;
      CDS_MarcaTitIdreceb.Value := DM_Tabelas.ZQRecebParticp.FieldByName('idrecebimento').Value;
      CDS_MarcaTitidparti.Value := DM_Tabelas.ZQRecebParticp.FieldByName('adversa').Value;
      CDS_MarcaTitcliente.Value := DM_Tabelas.ZQRecebParticp.FieldByName('cliente').Value;
      CDS_MarcaTitidloteam.Value := DM_Tabelas.ZQRecebParticp.FieldByName('idloteamento').Value;
      CDS_MarcaTitquadralote.Value:=DM_Tabelas.ZQRecebParticp.FieldByName('quadralote').Value;
      CDS_MarcaTitmarca.Value := '0';

      CDS_MarcaTitDT_Entrada.value:= DM_Tabelas.ZQRecebParticp.FieldByName('Dt_Entrada').Value;
      CDS_MarcaTitDt_Vencimento.value:= DM_Tabelas.ZQRecebParticp.FieldByName('Dt_Vencimento').Value;
      CDS_MarcaTitValor.value:= DM_Tabelas.ZQRecebParticp.FieldByName('Valor').Value;
      CDS_MarcaTitordem.value:= DM_Tabelas.ZQRecebParticp.FieldByName('ordem').Value;
      CDS_MarcaTitnome_parte.value:= DM_Tabelas.ZQRecebParticp.FieldByName('nome_parte').Value;
      CDS_MarcaTitende_cob.value:= DM_Tabelas.ZQRecebParticp.FieldByName('ende_cob').Value;
      CDS_MarcaTitcep_cob.value:= DM_Tabelas.ZQRecebParticp.FieldByName('cep_cob').Value;
      CDS_MarcaTitdoc1.value:= DM_Tabelas.ZQRecebParticp.FieldByName('doc1').Value;
      CDS_MarcaTitTipDoc.value:= DM_Tabelas.ZQRecebParticp.FieldByName('TipDoc').Value;
      CDS_MarcaTitnumordem.value:= DM_Tabelas.ZQRecebParticp.FieldByName('numordem').Value;
      CDS_MarcaTitbairro.value:=DM_Tabelas.ZQRecebParticp.FieldByName('bairro_cob').Value;
      CDS_MarcaTitcidade.value:=DM_Tabelas.ZQRecebParticp.FieldByName('cidade_cob').Value;
      CDS_MarcaTit.Post;
      DM_Tabelas.ZQRecebParticp.Next;
    end;
    JSP1.Visible := False;
{    FrmCobrancaBancaria.Top := FrmRelRecebimento.Top + 40;
    FrmCobrancaBancaria.Left :=FrmRelRecebimento.Left+30;
    FrmCobrancaBancaria.showmodal;}
    dxbSelecionar.Enabled:=true;
    dxbselecionar.SetFocus;
    Exit;
  end;
  vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
  if not empty(XDEVencimentoInicio.Text) Then
    vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' até '+XDEVencimentoFinal.Text;
  if not empty(XDEEntradaInicio.Text) Then
    vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' até '+XDEEntradaFinal.Text;
  if not empty(XDEBaixasInicio.Text) Then
    vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' até '+XDEBaixasFinal.Text;
  vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;
  ZQTempReceber.SQL.Clear;
  ZQTempReceber.SQL.Add('select * from (recebimento as RE join participante as CL on cl.idpaticipante=re.cliente )');
  if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then
    ZQTempReceber.SQL.Add('join recbxhist as H  ON RE.idrecebimento=H.idrecib join receb_baixa as BX ON BX.refbaixa=H.refer');
  ZQTempReceber.SQL.Add(' where 1 = 1 ');
  if not empty(vardocumento) Then
      ZQTempReceber.SQL.Add(vardocumento);
  if RGFormato.ItemIndex <> 3 Then Begin
    if RGContas.ItemIndex = 0 Then
      ZQTempReceber.SQL.Add(' and recpag ='+quotedstr('R'))
    else
      ZQTempReceber.SQL.Add(' and recpag ='+quotedstr('P'));
  end;
  // 13/12/2010
  if CBDesconsidera.Checked Then
     ZQTempReceber.SQL.Add(' and ( Numboleto is null or Numboleto='+quotedstr('')+') ');

  if not empty(Varaleatorio) then
    ZQTempReceber.SQL.Add(' and idpaticipante in ('+Varaleatorio+')');

  if not empty(VarEmpree) then
     ZQTempReceber.SQL.Add(' and idloteamento in ('+VarEmpree+')');

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZQTempReceber.SQL.Add(' and quadralote in ('+varquadra+')');

  if (not empty(vardevedor)) and (RGContas.ItemIndex=0) then
    ZQTempReceber.SQL.Add(' and adversa in ('+vardevedor+')');
  if not empty(varDoc) Then
    ZQTempReceber.SQL.Add(' and TIPDOC in ('+varDoc+')');
  if CBSaldo.ItemIndex = 1 Then
     ZQTempReceber.SQL.Add(' and (RE.saldo <re.valor or RE.saldo is null)');
  if CBSaldo.ItemIndex = 2 Then
     ZQTempReceber.SQL.Add(' and RE.saldo > 0');
  if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
    ZQTempReceber.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
    ZQTempReceber.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
  end;
  if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
    ZQTempReceber.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
    ZQTempReceber.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
  end;
  if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and BX.DT_Rec between :dt5 and :dt6');
    ZQTempReceber.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
    ZQTempReceber.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
  end;
  if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and BX.dataref between :dt7 and :dt8');
    ZQTempReceber.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
    ZQTempReceber.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
  end;
  // 24/11/2010
  if CBDesconsidera.Checked Then
     ZQTempReceber.SQL.Add(' and  not exists (select * from remessa_receb where idrec=idrecebimento)');
  //
  ZQTempReceber.SQL.Add(' group by re.idrecebimento');
  ZQTempReceber.SQL.Add(' ORDER BY '+VarOrdem);
  ZQTempReceber.Open;
  //2
  if RGFormato.ItemIndex = 2 Then
  Begin
   Vrtexto := '';

   if cbmensal.Checked=true then
   begin
    ZQmensal.SQL.Clear;
//    ZQmensal.SQL.Add('DROP VIEW IF EXISTS `'+varschemata+'`.`relmensal`;');
    ZQmensal.SQL.Add('CREATE OR REPLACE VIEW `'+varschemata+'`.`relmensal2` AS ');
    ZQmensal.SQL.Add(' ( select month(re.dt_vencimento) as Mes, case month(re.dt_vencimento)');
    ZQmensal.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Março''');
    ZQmensal.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQmensal.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQmensal.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
    ZQmensal.SQL.Add(' end AS MESDESC, CASE when  re.valor then re.valor end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.valor, re.DT_Entrada, re.DT_Vencimento, ');
    ZQmensal.SQL.Add(' year(re.dt_vencimento) as Ano from recebimento as RE where re.saldo>0 and RE.TIPDOC in ('+varDoc+')');
    if not empty(VarEmpree) then
       ZQmensal.SQL.Add(' and re.idloteamento in ('+VarEmpree+')');
    if RGContas.ItemIndex = 0 Then
      ZQmensal.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZQmensal.SQL.Add(' and re.recpag ='+quotedstr('P'));
    if not empty(vardocumento) Then
      ZQmensal.SQL.Add(vardocumento);
    if not empty(varquadra) then
      ZQmensal.SQL.Add(' and re.quadralote in ('+varquadra+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQmensal.SQL.Add(' and re.DT_Entrada between :dt1 and :dt2');
      ZQmensal.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQmensal.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' até '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQmensal.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQmensal.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQmensal.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' até '+XDEVencimentoFinal.DateText;
    end;

    ZQmensal.SQL.Add(' order by ano,mes )');
    ZQmensal.ExecSQL;
    ZQrelatorio2.close;
    ZQrelatorio2.SQL.clear;
    ZQrelatorio2.SQL.Add('select *,sum(vr_rel) as valorpago, sum(valor) as valorpagojuros from relmensal2 group by mes,ano order by ano,mes');
    ZQrelatorio2.open;


    if FrmRelReceb02_mensal_ab=nil then
      FrmRelReceb02_mensal_ab:=TFrmRelReceb02_mensal_ab.Create(Application);
    if not empty(vrtexto) Then
      FrmRelReceb02_mensal_ab.RLLabel2.Caption := 'Período de : '+vrtexto;
    FrmRelReceb02_mensal_ab.RLLabel1.Caption := 'Contas a Receber';
    if RGContas.ItemIndex = 1 Then
    FrmRelReceb02_mensal_ab.RLLabel1.Caption := 'Contas a Pagar';
    FrmRelReceb02_mensal_ab.RLLabel17.Caption:='Vr. Saldo em Aberto';
    FrmRelReceb02_mensal_ab.RLDBResult2.Text:='Saldo Geral em Aberto:  ';
    FrmRelReceb02_mensal_ab.RLDBResult1.Text:='Saldo Total em Aberto:  ';
    FrmRelReceb02_mensal_ab.RLReport1.DataSource:=DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLReport1.DataSource:=FrmRelRecebimento.DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLDBText2.DataSource:=FrmRelRecebimento.DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLDBText8.DataSource:=FrmRelRecebimento.DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLDBText3.DataSource:=FrmRelRecebimento.DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLDBResult1.DataSource:=FrmRelRecebimento.DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLDBResult2.DataSource:=FrmRelRecebimento.DataZQrelatorio2;
    FrmRelReceb02_mensal_ab.RLReport1.Preview;
    ZQEmpree.Close;
    ZQmensal.close;
    ZQmensal2.close;    
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;

   if stotais.Checked=false then
   begin
    Vrtexto := '';
    ZqresumoReceb.SQL.Clear;
    ZqresumoReceb.SQL.Add('select * from recebimento as RE join participante as CL on cl.idpaticipante=re.cliente ');
    if RGContas.ItemIndex = 0 Then
      ZqresumoReceb.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZqresumoReceb.SQL.Add(' where recpag ='+quotedstr('P'));
    if not empty(vardocumento) Then
      ZqresumoReceb.SQL.Add(vardocumento);
    if not empty(varDoc) Then
      ZqresumoReceb.SQL.Add(' and TIPDOC in ('+varDoc+')');

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZqresumoReceb.SQL.Add(' and quadralote in ('+varquadra+')');


    if CBSaldo.ItemIndex = 1 Then
       ZqresumoReceb.SQL.Add(' and RE.saldo > 0');
    if CBSaldo.ItemIndex = 2 Then
       ZqresumoReceb.SQL.Add(' and (RE.saldo = 0 or RE.saldo is null)');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZqresumoReceb.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZqresumoReceb.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZqresumoReceb.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' até '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZqresumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' até '+XDEVencimentoFinal.DateText;
    end;
    ZqresumoReceb.SQL.Add(' GROUP BY RE.tipdoc ORDER BY '+VarOrdem);
    ZqresumoReceb.Open;
    ZQSomaBaixa.SQL.Clear;
    ZQSomaBaixa.SQL.Add('select *, sum(vr_rec) as somabaixa FROM (recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib)');
    ZQSomaBaixa.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
    if RGContas.ItemIndex = 0 Then
      ZQSomaBaixa.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQSomaBaixa.SQL.Add(' where recpag ='+quotedstr('P'));
    ZQSomaBaixa.SQL.Add(' group by b.refbaixa');
    ZQSomaBaixa.Open;
    ZQBaixa.SQL.Clear;
    ZQBaixa.SQL.Add(' SELECT * FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib');
    ZQBaixa.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
    ZQBaixa.SQL.Add(' join tipodocumento ON tipodoc= b.tipdoc');
    if RGContas.ItemIndex = 0 Then
      ZQBaixa.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQBaixa.SQL.Add(' where recpag ='+quotedstr('P'));
    ZQBaixa.Open;
    ZQAgruTipo.SQL.Clear;
    ZQAgruTipo.SQL.Add('Select *, sum(saldo) as somasaldo from recebimento');
    if RGContas.ItemIndex = 0 Then
      ZQAgruTipo.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQAgruTipo.SQL.Add(' where recpag ='+quotedstr('P'));
    ZQAgruTipo.SQL.Add(' group by TipDoc');
    ZQAgruTipo.Open;
    FrmRelReceb02.RLDBResult1.DataField := 'saldo';
    FrmRelReceb02.RLDBText9.DataField := 'saldo';
    FrmRelReceb02.RLLabel1.Caption := 'Contas a Receber';
    if RGContas.ItemIndex = 1 Then
    FrmRelReceb02.RLLabel1.Caption := 'Contas a Pagar';
    if CBSaldo.ItemIndex=2 then
    begin
      FrmRelReceb02.RLBand4.Visible:=false;
      FrmRelReceb02.RLBand7.Visible:=false;
    end
    else
    begin
      FrmRelReceb02.RLBand4.Visible:=true;
      FrmRelReceb02.RLBand7.Visible:=true;
    end;
    if FrmRelReceb02=nil then
      FrmRelReceb02:=TFrmRelReceb02.Create(Application);
    if not empty(vrtexto) Then
      FrmRelReceb02.RLLabel2.Caption := 'Período de : '+vrtexto;
    FrmRelReceb02.RLReport1.Preview;
    ZQEmpree.Close;
    vrtexto:='';
   end
   else
   begin
     vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
     if not empty(XDEVencimentoInicio.Text) Then
        vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' até '+XDEVencimentoFinal.Text;
     if not empty(XDEEntradaInicio.Text) Then
        vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' até '+XDEEntradaFinal.Text;
     if not empty(XDEBaixasInicio.Text) Then
        vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' até '+XDEBaixasFinal.Text;
     vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;
     ZQTempReceber2.SQL.Clear;
     ZQTempReceber2.SQL.Add('select *,sum(valor) as tvr, sum(saldo) as tsld from recebimento ');
     if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then
         ZQTempReceber2.SQL.Add('join recbxhist as H  ON idrecebimento=H.idrecib join receb_baixa as BX ON BX.refbaixa=H.refer');
//     ZQTempReceber2.SQL.Add(' where 1 = 1 ');
     if not empty(vardocumento) Then
        ZQTempReceber2.SQL.Add(vardocumento);
     if RGFormato.ItemIndex <> 3 Then
     Begin
        if RGContas.ItemIndex = 0 Then
           ZQTempReceber2.SQL.Add(' where recpag ='+quotedstr('R'))
        else
           ZQTempReceber2.SQL.Add(' and recpag ='+quotedstr('P'));
     end;
    // 13/12/2010
     if CBDesconsidera.Checked Then
       ZQTempReceber2.SQL.Add(' and ( Numboleto is null or Numboleto='+quotedstr('')+') ');

     if not empty(VarEmpree) then
        ZQTempReceber2.SQL.Add(' and idloteamento in ('+VarEmpree+')');
       // coloquei aqui no dia 15/06/2010
     if not empty(varquadra) then
        ZQTempReceber2.SQL.Add(' and quadralote in ('+varquadra+')');
     if not empty(varDoc) Then
        ZQTempReceber2.SQL.Add(' and TIPDOC in ('+varDoc+')');
     if CBSaldo.ItemIndex = 1 Then
        ZQTempReceber2.SQL.Add(' and (saldo < valor or saldo is null)');
     if CBSaldo.ItemIndex = 2 Then
        ZQTempReceber2.SQL.Add(' and saldo > 0');
     if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then
     Begin
       ZQTempReceber2.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
       ZQTempReceber2.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
       ZQTempReceber2.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
     end;
     if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
     Begin
       ZQTempReceber2.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
       ZQTempReceber2.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
       ZQTempReceber2.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
     end;
     if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then
     Begin
        ZQTempReceber2.SQL.Add(' and BX.DT_Rec between :dt5 and :dt6');
        ZQTempReceber2.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
        ZQTempReceber2.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
     end;
     if  not empty(copy(XDERefInicio.Text,1,2)) Then
     Begin
       ZQTempReceber2.SQL.Add(' and BX.Dataref between :dt7 and :dt8');
       ZQTempReceber2.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
       ZQTempReceber2.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
     end;

     // 24/11/2010
     if CBDesconsidera.Checked Then
        ZQTempReceber2.SQL.Add(' and not exists (select * from remessa_receb where idrec=idrecebimento)');
     //

     ZQTempReceber2.SQL.Add(' group by idloteamento');

     ZQTempReceber2.SQL.Add(' ORDER BY '+VarOrdem);
     ZQTempReceber2.Open;


     Vrtexto := '';
{     ZqresumoReceb.SQL.Clear;
     ZqresumoReceb.SQL.Add('select * from recebimento ');
     if RGContas.ItemIndex = 0 Then
        ZqresumoReceb.SQL.Add(' where recpag ='+quotedstr('R'))
     else
        ZqresumoReceb.SQL.Add(' where recpag ='+quotedstr('P'));
     if not empty(vardocumento) Then
        ZqresumoReceb.SQL.Add(vardocumento);
     if not empty(varDoc) Then
        ZqresumoReceb.SQL.Add(' and TIPDOC in ('+varDoc+')');

     // coloquei aqui no dia 15/06/2010
     if not empty(varquadra) then
        ZqresumoReceb.SQL.Add(' and quadralote in ('+varquadra+')');
     if CBSaldo.ItemIndex = 1 Then
        ZqresumoReceb.SQL.Add(' and saldo > 0');
     if CBSaldo.ItemIndex = 2 Then
        ZqresumoReceb.SQL.Add(' and (saldo = 0 or saldo is null)');}
     if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then
     Begin
{       ZqresumoReceb.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
       ZqresumoReceb.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
       ZqresumoReceb.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);}
       Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' até '+XDEEntradaFinal.DateText;
     end;
     if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
     Begin
{      ZqresumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);}
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' até '+XDEVencimentoFinal.DateText;
     end;
{     ZqresumoReceb.SQL.Add(' GROUP BY idloteamento ORDER BY '+VarOrdem);
     ZqresumoReceb.SQL.Text;
     ZqresumoReceb.Open;

     ZQSomaBaixa.SQL.Clear;
     ZQSomaBaixa.SQL.Add('select *, sum(vr_rec) as somabaixa FROM (recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib)');
     ZQSomaBaixa.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
     if RGContas.ItemIndex = 0 Then
        ZQSomaBaixa.SQL.Add(' where recpag ='+quotedstr('R'))
     else
       ZQSomaBaixa.SQL.Add(' where recpag ='+quotedstr('P'));
     ZQSomaBaixa.SQL.Add(' group by b.refbaixa');
     ZQSomaBaixa.Open;
     ZQBaixa.SQL.Clear;
     ZQBaixa.SQL.Add(' SELECT * FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib');
     ZQBaixa.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
     ZQBaixa.SQL.Add(' join tipodocumento ON tipodoc= b.tipdoc');
     if RGContas.ItemIndex = 0 Then
       ZQBaixa.SQL.Add(' where recpag ='+quotedstr('R'))
     else
       ZQBaixa.SQL.Add(' where recpag ='+quotedstr('P'));
     ZQBaixa.Open;
     ZQAgruTipo.SQL.Clear;
     ZQAgruTipo.SQL.Add('Select *, sum(saldo) as somasaldo from recebimento');
     if RGContas.ItemIndex = 0 Then
       ZQAgruTipo.SQL.Add(' where recpag ='+quotedstr('R'))
     else
       ZQAgruTipo.SQL.Add(' where recpag ='+quotedstr('P'));
     ZQAgruTipo.SQL.Add(' group by idloteamento');
     ZQAgruTipo.Open;}
//     FrmRelReceb02_total.RLDBResult1.DataField := 'saldo';
//     FrmRelReceb02_total.RLDBText9.DataField := 'saldo';
     FrmRelReceb02_total.RLLabel1.Caption := 'Contas a Receber';
     if RGContas.ItemIndex = 1 Then
        FrmRelReceb02_total.RLLabel1.Caption := 'Contas a Pagar';
     if FrmRelReceb02_total=nil then
        FrmRelReceb02_total:=TFrmRelReceb02_total.Create(Application);
     if not empty(vrtexto) Then
        FrmRelReceb02_total.RLLabel2.Caption := 'Período de : '+vrtexto;
     FrmRelReceb02_total.RLReport1.Preview;
     ZQEmpree.Close;
     vrtexto:='';

   end;
  end
  else if RGFormato.ItemIndex = 3 Then Begin
    CDSPlano.Close;
    CDSPlano.CreateDataSet;
    CDSPlano.IndexFieldNames := 'mascara';
    CDSDetal.Close;
    CDSDetal.CreateDataSet;
    CDSDetal.IndexFieldNames := 'data';
    CDSDetal.Filtered := True;
    JSP1.Visible := True;
    JSP1.Maximum := ZQTempReceber.RecordCount;
    while not ZQTempReceber.Eof do begin
      JSP1.Position := ZQTempReceber.RecNo;
      for varvezes:=1 to 5 do begin
        varmasc := copy(ZQTempReceber.FieldByName('planomascara').Value,1,vartama[varvezes]);
        if length(varmasc)<>vartama[varvezes] Then
          break;
        if CDSPlano.Locate('mascara',varmasc,[]) Then
          CDSPlano.Edit
        else
          CDSPlano.Insert;
        CDSPlanomascara.Value := varmasc;
        CDSPlanosoma.Value := CDSPlanosoma.Value + ZQTempReceber.FieldByName('Valor').Value;
        CDSPlano.Post;
      end;
      CDSDetal.Insert;
      CDSDetalmascara.Value := ZQTempReceber.FieldByName('planomascara').Value;
      CDSDetalvalor.Value := ZQTempReceber.FieldByName('Valor').Value;
      CDSDetaldata.Value := ZQTempReceber.FieldByName('Dt_Vencimento').Value;
      CDSDetaldetalhe.Value := 'Lanç: '+ZQTempReceber.FieldByName('idrecebimento').Text+'  Participante: '+ZQTempReceber.FieldByName('nome_parte').Value+'      Adversa: '+ZQTempReceber.FieldByName('adversanome').Value+chr(13);
      CDSDetaldetalhe.Value := CDSDetaldetalhe.Value + ZQTempReceber.FieldByName('Observ').Value;
      ZQTempReceber.Next;
    end;
    if CBObsLanc.Checked Then
      FrmRelRecebGerencial.RLBand4.Visible := True
    else
      FrmRelRecebGerencial.RLBand4.Visible := False;
    if FrmRelRecebGerencial=nil then
      FrmRelRecebGerencial:=TFrmRelRecebGerencial.Create(Application);
    FrmRelRecebGerencial.RLReport1.Preview;
    ZQEmpree.Close;    
  end;
  JSP1.Visible := False;
end;

procedure TFrmlimparemessa.XDEEntradaInicioExit(Sender: TObject);
begin
  if XDEBaixasfinal.DateValue<XDEBaixasinicio.DateValue Then
     XDEBaixasfinal.DateValue := XDEBaixasinicio.DateValue;
end;

procedure TFrmlimparemessa.XDEVencimentoInicioExit(Sender: TObject);
begin
  if XDEVencimentoFinal.DateValue<XDEVencimentoInicio.DateValue Then
     XDEVencimentoFinal.DateValue := XDEVencimentoInicio.DateValue;
end;

procedure TFrmlimparemessa.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    if not (ActiveControl is TDBGrid) then begin
      Key := #0;
      Perform(WM_NEXTDLGCTL, 0, 0);
    end;
end;

procedure TFrmlimparemessa.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQBaixa.Close;
  ZQTempReceber.Close;
  ZQTempReceber2.close;
  DM_Tabelas.ZQincorp_loteame.close;
  DM_Tabelas.ZQLoteamento.close;

end;

procedure TFrmlimparemessa.GBEscClienteExit(Sender: TObject);
begin
  FrmRelRecebimento.KeyPreview := True
end;

procedure TFrmlimparemessa.ZQBaixaCalcFields(DataSet: TDataSet);
begin
  if ZQBaixa.FieldByName('percent_usado').Value >0 Then
    ZQBaixa.FieldByName('vrusado').Value := ZQBaixa.FieldByName('Vr_rec').Value*ZQBaixa.FieldByName('percent_usado').Value/100
  else
    ZQBaixa.FieldByName('vrusado').Value := ZQBaixa.FieldByName('Vr_rec').Value;
end;

procedure TFrmlimparemessa.EClienteExit(Sender: TObject);
begin
  if not empty(ECliente.Text) Then Begin
    if AchaCLiente(200, 10,ECliente.Text) Then Begin
      CDSCli.Insert;
      CDSCliCCli.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').Value;
      CDSCli.Post;
    end;
    ECliente.Text := '';
    ECliente.SetFocus;
  end;
end;

procedure TFrmlimparemessa.DBGrid3KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSCli.RecordCount>0) Then Begin
    CDSCli.Delete;
  end;
end;

procedure TFrmlimparemessa.DBGrid4KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;
end;

procedure TFrmlimparemessa.RGContasClick(Sender: TObject);
begin
  if RGContas.ItemIndex = 0 then
    GBAdversa.Enabled := True
  else
    GBAdversa.Enabled := False;
end;

procedure TFrmlimparemessa.RGFormatoClick(Sender: TObject);
begin
{  if RGFormato.ItemIndex = 5 Then
     dxbx.Visible:=true
  else
     dxbx.Visible:=false;}
  if ( RGFormato.ItemIndex = 0 ) or ( RGFormato.ItemIndex = 1 ) or (RGFormato.ItemIndex = 4) or ( RGFormato.ItemIndex = 3) or ( RGFormato.ItemIndex = 2 ) Then
  begin
    CBDesconsidera.Visible := True;
    XDERefInicio.Clear;
    XDERefFinal.Clear;
    XDEBaixasinicio.Clear;
    XDEBaixasfinal.Clear;
    GroupBox3.Visible:=false;
    GroupBox7.Visible:=false;
    cbmensal.Visible:=false;
    xresumo.Visible:=false;
  end
  else
  begin
    CBDesconsidera.Visible := False;
    GroupBox3.Visible:=true;
    GroupBox7.Visible:=true;
    xresumo.Visible:=false;
  end;
  if RGFormato.ItemIndex = 4 Then
  Begin
    dxbSelecionar.visible:=true;
    cbmensal.Visible:=false;
    CBDesconsidera.Checked:=true;
    xresumo.Visible:=false;
  end
  else
  begin
    dxbSelecionar.visible:=false;
    xresumo.Visible:=false;
  end;
  if RGFormato.ItemIndex = 2 Then
  Begin
    CBSaldo.ItemIndex:=2;
    CBVerbaixa.ItemIndex:=1;
    GroupBox6.Visible:=false;
    GroupBox8.Visible:=false;
    stotais.Visible:=true;
    cbmensal.Visible:=true;
    CBDesconsidera.Checked:=false;
    xresumo.Visible:=false;
  end
  else
  begin
    stotais.Visible:=false;
    xresumo.Visible:=false;
  end;
  if RGFormato.ItemIndex = 5 Then
  Begin
    CBSaldo.ItemIndex:=1;
    CBVerbaixa.ItemIndex:=0;
    GroupBox6.Visible:=false;
    GroupBox8.Visible:=false;
    cbmensal.Visible:=true;
    xresumo.Visible:=true;
  end;
  if RGFormato.ItemIndex = 4 Then
  Begin
    CBSaldo.ItemIndex:=2;
    CBVerbaixa.ItemIndex:=1;
    GroupBox6.Visible:=false;
    GroupBox8.Visible:=false;
    cbmensal.Visible:=false;
    xresumo.Visible:=false;
  end;
  if RGFormato.ItemIndex=5 then
  begin
    RGOrdem.Items.Clear;
    RGOrdem.Items.Add('Vencimento');
    RGOrdem.Items.Add('Comprador');
    RGOrdem.Items.Add('Quadra/Lote');
    RGOrdem.Items.Add('Tipo');
    RGOrdem.Items.Add('Baixa');
    RGOrdem.ItemIndex:=4;
  end
  else
  begin
    RGOrdem.Items.Clear;
    RGOrdem.Items.Add('Vencimento');
    RGOrdem.Items.Add('Comprador');
    RGOrdem.Items.Add('Quadra/Lote');
    RGOrdem.Items.Add('Tipo');
    RGOrdem.ItemIndex:=0;
  end;
end;


procedure TFrmlimparemessa.EloteExit(Sender: TObject);
begin
  if not empty(Elote.Text+EQuadra.Text) Then Begin
    ZQVen1.SQL.Clear;
    ZQVen1.SQL.Add('Select * from venda join imovel on idimovel=imovel');
    ZQVen1.SQL.Add(' where quadra='+quotedstr(Equadra.Text)+' and lote='+quotedstr(Elote.Text));
    ZQVen1.Open;
    if ZQVen1.RecordCount>0 Then Begin
      CDSQuadraLote.Insert;
      CDSQuadraLoteidlote.Value := ZQVen1.FieldByName('idimovel').Value;
      CDSQuadraLoteQuadra.Value := Equadra.Text+'-'+Elote.Text;
      CDSQuadraLote.Post;
      if ZQVen1.FieldByName('datavenda').Value<XDEEntradaInicio.DateValue Then
        XDEEntradaInicio.DateValue := ZQVen1.FieldByName('datavenda').Value;
//      if ZQVen1datavenda.Value>XDEEntradaFinal.DateValue Then
//        XDEEntradaFinal.DateValue := ZQVen1datavenda.Value;
    end
    else
      Showmessage('Esta Quadra e lote não foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end
  else
    if CDSQuadraLote.RecordCount>0 Then
      DXBImprimir.SetFocus;
end;

procedure TFrmlimparemessa.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;
end;

procedure TFrmlimparemessa.XDERefInicioExit(Sender: TObject);
begin
  if XDERefFinal.DateValue<XDERefInicio.DateValue Then
     XDERefFinal.DateValue := XDERefInicio.DateValue;
end;

procedure TFrmlimparemessa.dxbSelecionarClick(Sender: TObject);
begin
    dxbSelecionar.enabled:=false;
    CDS_MarcaTit.IndexName:='nome_venc';
    FrmCobrancaBancaria.Top := FrmRelRecebimento.Top + 40;
    FrmCobrancaBancaria.Left :=FrmRelRecebimento.Left+30;
    FrmCobrancaBancaria.showmodal;
    CDS_MarcaTit.IndexName:='';    
end;

procedure TFrmlimparemessa.CBSaldoExit(Sender: TObject);
begin
  if CBSaldo.ItemIndex=2 then
  begin
    CBVerbaixa.ItemIndex:=1;
  end
  else
  begin
    CBVerbaixa.ItemIndex:=0;
  end;

end;

procedure TFrmlimparemessa.Edit1Exit(Sender: TObject);
begin
  IF NOT EMPTY(Edit1.Text) Then Begin
    if AchaCliente(400, 300, Edit1.Text) then Begin
      CDSAdversaTemp2.Insert;
      CDSAdversaTemp2CodAdversa.Value := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Value;
      CDSAdversaTemp2.Post;
    end;
    Edit1.Text := '';
    Edit1.SetFocus;
  end;

end;


procedure TFrmlimparemessa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQBaixa', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaTipdoc_1', 'Tipdoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixadescricao', 'descricao', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixatipodoc', 'tipodoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixavend_receb', 'vend_receb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixavend_caixa', 'vend_caixa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixareceb_caixa', 'receb_caixa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixareceb_receb', 'receb_receb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixadados_chequ', 'dados_chequ', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaso_avista', 'so_avista', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixalancabanco', 'lancabanco', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixasomapaga', 'somapaga', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixalancinclus', 'lancinclus', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixavrusado', 'vrusado', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaidtipodocumento', 'idtipodocumento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixanumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQBaixa', 'ZQBaixasubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQResumoReceb', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebtipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoReceblocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQResumoReceb', 'ZQResumoRecebnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQSomaBaixa', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixasomabaixa', 'somabaixa', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixanumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQSomaBaixa', 'ZQSomaBaixasubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQAgruTipo', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoidRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTiposaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTiposomasaldo', 'somasaldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTiporecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTiponumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTiporefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTiponumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAgruTipo', 'ZQAgruTipoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQTemp2Cl', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Cldoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Cldoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2ClFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Cltipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Claniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Cllocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2ClcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2ClcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTemp2Cl', 'ZQTemp2Clcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQRateio', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiotipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateiovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateionumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRateio', 'ZQRateioSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreeidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreematriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreetestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreetestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreetestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreetestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreeObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreebairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreepasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreenomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreetipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQEmpree', 'ZQEmpreeapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQVenda', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaquadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendalote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendalado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendamedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendamedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendamedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendamedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendaObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendacornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVenda', 'ZQVendadisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQVen1', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1quadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1lote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1lado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZROQReceb', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsomatit', 'somatit', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsomapago', 'somapago', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsomajuros', 'somajuros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebnomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebnomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsubstituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb', 'ZROQRecebnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZROQReceb1', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1idreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Dt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Vr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Docum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Juros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Desc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1obsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Codcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1vencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1idrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1refer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1idrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1valor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1descon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1juros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1percent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1data', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Valor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1RefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1numboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1nomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1nomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1substituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1Substituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1vr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1sq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1sq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1TipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1sq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1somar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQReceb1', 'ZROQReceb1nomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZROQCh', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChbanco', 'banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChagencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChdono', 'dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChCPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChalias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChconta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChnumero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZROQCh', 'ZROQChsubstituicao_2', 'substituicao_2', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQTempReceber2', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2planomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2adversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2nomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2numboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2Substituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2tvr', 'tvr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2tsld', 'tsld', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQTempReceber2', 'ZQTempReceber2nomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQmensal', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalTIPDOC', 'TIPDOC', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalvr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalnomelote', 'nomelote', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalMESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal', 'ZQmensalvr_rel', 'vr_rel', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQmensal2', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2MESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2vr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2tipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2DT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2DT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2vr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQmensal2', 'ZQmensal2juros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQRelatorio2', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2MESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2vr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2tipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2DT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2DT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio2', 'ZQRelatorio2valorpago', 'valorpago', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZQRelatorio', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioMESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriovr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriotipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriovr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatorioAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriovalorpago', 'valorpago', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriodataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZQRelatorio', 'ZQRelatoriovalorpagojuros', 'valorpagojuros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZqresumoE', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoETipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoERefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoECodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoESubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoErefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoETipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoERefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoErefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoErecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoESubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoE', 'ZqresumoEData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmlimparemessa, 'ZqresumoP', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmlimparemessa, 'ZqresumoP', 'ZqresumoPData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
