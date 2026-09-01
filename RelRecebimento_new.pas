unit RelRecebimento_new;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, XEdit, Mask, XDate, ExtCtrls, XBanner,
  XLabel3D,   dxButton, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, JvExControls, JvSpecialProgress, Grids, DBGrids, DBClient,
  DBCtrls, TFlatCheckBoxUnit, ActnList, IniFiles, XNum, Provider;

type
  TFrmRelRecebimento = class(TForm)
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













    CbCarteira: TCheckBox;
    DS_zqcarteira: TDataSource;
    ZQcarteira: TZQuery;














    ZQRelCarteira: TZQuery;
    DS_ZQRelCarteira: TDataSource;























    ZRdiario: TZReadOnlyQuery;
    DS_ZRdiario: TDataSource;































































    CBdiario: TCheckBox;

    DSincorp_loteame: TDataSource;
    ZQincorp_loteame: TZQuery;






























    DS_Empreendimento: TDataSource;
    CDSEmpreendimento: TClientDataSet;
    CDSEmpreendimentoidloteamento: TIntegerField;
    CDSEmpreendimentomarca: TStringField;
    CDSEmpreendimentonomeloteamento: TStringField;
    CDSEmpreendimentocodcontabancaria: TIntegerField;
    CDSEmpreendimentoconta: TStringField;
    CDSEmpreendimentoconta_v: TStringField;
    CDSEmpreendimentonomebanco: TStringField;
    CDSEmpreendimentoidparticipante: TIntegerField;
    CDSEmpreendimentonomeparticipante: TStringField;
    CDSEmpreendimentoidordem: TIntegerField;
    CDSEmpreendimentolayoutBoleto: TStringField;
    CDSEmpreendimentonomearquivo: TStringField;
    CDSEmpreendimentocod_no_banco: TStringField;
    CDSEmpreendimentonometitular: TStringField;
    CDSEmpreendimenton_agencia: TStringField;
    CDSEmpreendimenton_agencia_v: TStringField;
    CDSEmpreendimentojurosemboleto: TFloatField;
    CDSEmpreendimentocarteira: TStringField;
    CDSEmpreendimentoapelido: TStringField;
    CDSEmpreendimentodigito_dif: TStringField;
    CDSEmpreendimentoconvenio: TStringField;
    ZQRecebParticp: TZReadOnlyQuery;




















    DS_RecebParticp: TDataSource;
    ZQLoteamento: TZQuery;



















    DS_Loteamento: TDataSource;
    ZQCidade: TZQuery;






    DS_Cidade: TDataSource;













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
  FrmRelRecebimento: TFrmRelRecebimento;
  varschemata : string;
implementation

uses Tabelas, Funcoes, Principal, RelReceb02, Cad_Recebimento,
  RelRecebGerencial, CobrancaBancaria, Recebi03, RelReceb02_total,
  RelReceb02_mensal, RelReceb02_mensal_ab, RelReceb02_carteira,
  rel_diario_lote, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelRecebimento.FormShow(Sender: TObject);
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

procedure TFrmRelRecebimento.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRelRecebimento.DXBImprimirClick(Sender: TObject);
var
  VarEmpree, VarDoc, VarOrdem, vartitulo, Varaleatorio, vardevedor, Vrtexto,
  dentromes,forames,data,d,m,a,entrada,vcto,baixa,varmasc, vardocumento, varquadra, varsql,VarLogo : String;
  mes1,mes2,varcontador, x, x2, y, varvezes, posi : Integer;
const
  vartama : array[1..5] of integer = (1,3,6,10,15);

begin
  x2:=0;
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
      1: VArOrdem := 'nomeadversa';
      2: VArOrdem := 'quadralote';
      3: VArOrdem := 'TipDoc';
      4: VArOrdem := 'dataref';
      5: VarOrdem := 'DT_Entrada';
    end;
  end
  else
  begin
    case RGOrdem.ItemIndex of
      0: VarOrdem := 'DT_Vencimento';
      1: VArOrdem := 'nomeadversa';
      2: VArOrdem := 'quadralote';
      3: VArOrdem := 'TipDoc';
      4: VArOrdem := 'dataref';
      5: VarOrdem := 'DT_Entrada';      
    end;
  end;
  JSP1.Visible := True;
  VarDoc := '';
  varcontador := 0;
  JSP1.Maximum := CLBEmpree.Count-1;
  VarEmpree :='';
  for x := 0 to CLBEmpree.Count-1 do begin
    JSP1.Position := x;
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
    ZQmensal2.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Mar�o''');
    ZQmensal2.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQmensal2.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQmensal2.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
//    ZQmensal2.SQL.Add(' end AS MESDESC, CASE when (re.valor <> rb.vr_rec) then re.valor else rb.Vr_rec end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
    ZQmensal2.SQL.Add(' end AS MESDESC, rh.valor as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQmensal2.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQmensal2.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
    end;

    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZQmensal2.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZQmensal2.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' � '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZQmensal2.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZQmensal2.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Refer�ncia de '+XDERefInicio.DateText+' � '+XDERefFinal.DateText;
    end;
//    ZQmensal2.SQL.Add(' group by mes,ano order by ano,mes');
    ZQmensal2.SQL.Add(' order by ano,mes )');
//    ZQmensal2.Open;
  //  ZQmensal2.First;
    ZQmensal2.ExecSQL;
    ZQrelatorio.close;
    ZQrelatorio.SQL.clear;
    ZQrelatorio.SQL.Add('select *,sum(vr_rel) as valorpagojuros, sum(valor) as valorpago from relmensal group by mes,ano order by ano,mes');
    ZQrelatorio.open;

    if FrmRelReceb02_mensal=nil then
      FrmRelReceb02_mensal:=TFrmRelReceb02_mensal.Create(Application);
    if not empty(vrtexto) Then
      FrmRelReceb02_mensal.RLLabel2.Caption := 'Per�odo de : '+vrtexto;
    FrmRelReceb02_mensal.RLLabel1.Caption := 'Contas Recebidas';
    if RGContas.ItemIndex = 1 Then

    FrmRelReceb02_mensal.RLLabel1.Caption := 'Contas Pagas';
    FrmRelReceb02_mensal.RLLabel17.Caption:='Vr. Saldo Baixado  ';
    FrmRelReceb02_mensal.RLDBResult2.Text:='Saldo Geral Baixado:    ';
    FrmRelReceb02_mensal.RLDBResult1.Text:='Saldo Total Baixado:    ';
    FrmRelReceb02_mensal.RLLabel11.Caption:='Vr. Saldo Baixado/Varia��o';
//    FrmRelReceb02_mensal.RLDBResult3.Text:='Saldo Geral Baixado/Varia��o:    ';
//    FrmRelReceb02_mensal.RLDBResult4.Text:='Saldo Total Baixado/Varia��o:    ';

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



   if cbdiario.Checked=true then
   begin
    ZRdiario.SQL.Clear;
    ZRdiario.SQL.Add('SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa ');
    ZRdiario.SQL.Add(' join recebimento as RE on RH.idrecib=RE.idrecebimento');
    ZRdiario.SQL.Add(' join loteamento as lt on lt.idloteamento=RE.idloteamento');
    ZRdiario.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');
    if not empty(vardevedor) then
      ZRdiario.SQL.Add(' and adversa in ('+vardevedor+')');

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZRdiario.SQL.Add(' and quadralote in ('+varquadra+')');
    if RGContas.ItemIndex = 0 Then
      ZRdiario.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZRdiario.SQL.Add(' and re.recpag ='+quotedstr('P'));
    if not empty(Varaleatorio) then
      ZRdiario.SQL.Add(' and cliente in ('+Varaleatorio+')');
    x2:=length(VarEmpree);
    if not empty(VarEmpree) then
      ZRdiario.SQL.Add(' and RE.idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZRdiario.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZRdiario.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZRdiario.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZRdiario.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZRdiario.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZRdiario.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' � '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZRdiario.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZRdiario.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Refer�ncia de '+XDERefInicio.DateText+' � '+XDERefFinal.DateText;
    end;
    ZRdiario.SQL.Add(' order by '+VArOrdem);
    ZRdiario.Open;



    if frm_rel_diario=nil then
      frm_rel_diario:=Tfrm_rel_diario.Create(Application);
    if not empty(vrtexto) Then
      frm_rel_diario.RLLabel2.Caption := 'Per�odo de : '+vrtexto;

    frm_rel_diario.RLGroup1.DataFields:='idloteamento;'+VArOrdem;
    if (x2=6) or (x2=8) or (x2=10) then
    begin
      frm_rel_diario.RLLabel10.Visible:=false;
      frm_rel_diario.RLDBText2.Visible:=false;
    end
    else
    begin
      frm_rel_diario.RLLabel10.Visible:=true;
      frm_rel_diario.RLDBText2.Visible:=true;
    end;

    frm_rel_diario.RLReport1.Preview;
    ZQEmpree.Close;
    ZRdiario.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;



   if cbcarteira.Checked=true then
   begin
    ZQcarteira.SQL.Clear;
//    ZQcarteira.SQL.Add('DROP VIEW IF EXISTS `'+varschemata+'`.`relmensal`;');
    ZQcarteira.SQL.Add('CREATE OR REPLACE VIEW `'+varschemata+'`.`relcarteira` AS ');
    ZQcarteira.SQL.Add(' ( select month(rb.dataref) as Mes, lt.apelido as ap, case month(rb.dataref)');
    ZQcarteira.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Mar�o''');
    ZQcarteira.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQcarteira.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQcarteira.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
//    ZQcarteira.SQL.Add(' end AS MESDESC, CASE when (re.valor <> rb.vr_rec) then re.valor else rb.Vr_rec end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
    ZQcarteira.SQL.Add(' end AS MESDESC, rh.valor as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
    ZQcarteira.SQL.Add(' year(rb.dataref) as Ano from receb_baixa as rb join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento join loteamento as lt on RE.idloteamento=lt.idloteamento where RB.TIPDOC in ('+varDoc+')');
    if not empty(VarEmpree) then
       ZQcarteira.SQL.Add(' and re.idloteamento in ('+VarEmpree+')');
    if RGContas.ItemIndex = 0 Then
      ZQcarteira.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      ZQcarteira.SQL.Add(' and re.recpag ='+quotedstr('P'));
    if not empty(vardocumento) Then
      ZQcarteira.SQL.Add(vardocumento);
    if not empty(varquadra) then
      ZQcarteira.SQL.Add(' and re.quadralote in ('+varquadra+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and re.DT_Entrada between :dt1 and :dt2');
      ZQcarteira.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQcarteira.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQcarteira.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQcarteira.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
    end;

    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZQcarteira.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZQcarteira.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' � '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZQcarteira.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZQcarteira.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Refer�ncia de '+XDERefInicio.DateText+' � '+XDERefFinal.DateText;
    end;
//    ZQcarteira.SQL.Add(' group by mes,ano order by ano,mes');
    ZQcarteira.SQL.Add(' order by ano,mes )');
//    ZQcarteira.Open;
  //  ZQcarteira.First;
    ZQcarteira.ExecSQL;

    ZQRelCarteira.close;
    ZQRelCarteira.SQL.clear;
    ZQRelCarteira.SQL.Add('select *,sum(vr_rel) as valorpagojuros, sum(valor) as valorpago from relcarteira group by idloteamento,mes,ano order by idloteamento,ano,mes');
    ZQRelCarteira.open;


    if FrmRelReceb02_carteira=nil then
      FrmRelReceb02_carteira:=TFrmRelReceb02_carteira.Create(Application);
    if not empty(vrtexto) Then
      FrmRelReceb02_carteira.RLLabel2.Caption := 'Per�odo de : '+vrtexto;
    FrmRelReceb02_carteira.RLLabel1.Caption := 'Administra��o da Carteira de Receb�veis';
    if RGContas.ItemIndex = 1 Then

    FrmRelReceb02_carteira.RLLabel1.Caption := 'Administra��o da Carteira de Receb�veis';
    FrmRelReceb02_carteira.RLLabel17.Caption:='Vr. Saldo Baixado  ';
    FrmRelReceb02_carteira.RLDBResult2.Text:='Saldo Geral Baixado:    ';
    FrmRelReceb02_carteira.RLDBResult1.Text:='Saldo Total Baixado:    ';
    FrmRelReceb02_carteira.RLLabel11.Caption:='Vr. Saldo Baixado/Varia��o';
//    FrmRelReceb02_carteira.RLDBResult3.Text:='Saldo Geral Baixado/Varia��o:    ';
//    FrmRelReceb02_carteira.RLDBResult4.Text:='Saldo Total Baixado/Varia��o:    ';

    FrmRelReceb02_carteira.RLReport1.DataSource:=DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLReport1.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBText2.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBText8.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBText3.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBResult1.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBResult2.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBText4.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBResult3.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;
    FrmRelReceb02_carteira.RLDBResult4.DataSource:=FrmRelRecebimento.DS_ZQRelCarteira;

    FrmRelReceb02_carteira.RLReport1.Preview;
    ZQEmpree.Close;
    ZQcarteira.close;
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
//    ZROQReceb.SQL.Add('SELECT *, sum(Re.valor) as somatit, sum(vr_rec) as somapago, sum(juros_vr) as somajuros FROM receb_baixa as RB');
    ZROQReceb.SQL.Add('SELECT *, sum(Re.valor) as somatit, sum(RH.valor) as somapago, sum(rb.juros_vr) as somajuros FROM receb_baixa as RB');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZROQReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZROQReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZROQReceb.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZROQReceb.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' � '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZROQReceb.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZROQReceb.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Refer�ncia de '+XDERefInicio.DateText+' � '+XDERefFinal.DateText;
    end;
    ZROQReceb.SQL.Add(' group by RB.tipdoc ');
    ZROQReceb.Open;

    if not empty(vrtexto) Then
      FRM_Recebi03.RLLabel2.Caption := 'Per�odo de : '+vrtexto;
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
      if mes(ZqresumoE.FieldByName('dataref').Value)=mes(ZqresumoE.FieldByName('Dt_Vencimento').Value) then //mes(ZqresumoEvencimento.Value) then
      begin
{        if ZqresumoEvr_rec.Value>ZqresumoEvalor_1.Value then
        begin
           xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoEvalor_1.Value+ZqresumoEJuros_Vr.Value;
        end
        else
        begin
           if ZqresumoEVr_rec.Value>0 then
           begin}
             xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoE.FieldByName('valor').Value;//+ZqresumoEJuros_Vr.Value;
{           end
           else
           begin
             xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoEvr_ab.Value;
           end;
        end;}
      end
      else
      begin
        mes1:=strtoint(mes(ZqresumoE.FieldByName('dataref').Value));
        mes2:=strtoint(mes(ZqresumoE.FieldByName('Dt_Vencimento').Value)); //strtoint(mes(ZqresumoEvencimento.Value));        
        if mes1>mes2 then
        begin
          {if ZqresumoEvr_rec.Value>ZqresumoEvalor_1.Value then
          begin
             xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoEvalor_1.Value+ZqresumoEJuros_Vr.Value;
             entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoEvalor_1.Value+ZqresumoEJuros_Vr.Value;
          end
          else
          begin
             if ZqresumoEVr_rec.Value>0 then
             begin}
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('valor').Value;//+ZqresumoEJuros_Vr.Value;
               entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoE.FieldByName('valor').Value;//+ZqresumoEJuros_Vr.Value;
{             end
             else
             begin
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoEvr_ab.Value;
               entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoEvr_ab.Value;
             end;
          end;}
        end
        else
        begin
          {if ZqresumoEvr_rec.Value>ZqresumoEvalor_1.Value then
          begin
             xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoEvalor_1.Value+ZqresumoEJuros_Vr.Value;
             entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoEvalor_1.Value+ZqresumoEJuros_Vr.Value;
          end
          else
          begin
             if ZqresumoEVr_rec.Value>0 then
             begin}
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('valor').Value;//+ZqresumoEJuros_Vr.Value;
               entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoE.FieldByName('valor').Value;//+ZqresumoEJuros_Vr.Value;
{             end
             else
             begin
               xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoEvr_ab.Value;
               entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoEvr_ab.Value;
             end;
          end;}
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
      if mes(ZqresumoP.FieldByName('dataref').Value)=mes(ZqresumoP.FieldByName('Dt_Vencimento').Value) then // mes(ZqresumoPvencimento.Value) then
      begin
       { if ZqresumoPvr_rec.Value>ZqresumoPvalor_1.Value then
        begin
           xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoPvalor_1.Value+ZqresumoPJuros_Vr.Value;
        end
        else
        begin
           if ZqresumoPVr_rec.Value>0 then
           begin }
             xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoP.FieldByName('valor').Value;//+ZqresumoPJuros_Vr.Value;
{           end
           else
           begin
             xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoPvr_ab.Value;
           end;
        end;}
      end
      else
      begin
        mes1:=strtoint(mes(ZqresumoP.FieldByName('dataref').Value));
        mes2:=strtoint(mes(ZqresumoP.FieldByName('Dt_Vencimento').Value)); //strtoint(mes(ZqresumoPvencimento.Value));
        if mes1>mes2 then
        begin
          {if ZqresumoPvr_rec.Value>ZqresumoPvalor_1.Value then
          begin
             xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoPvalor_1.Value+ZqresumoPJuros_Vr.Value;
             parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoPvalor_1.Value+ZqresumoPJuros_Vr.Value;
          end
          else
          begin
             if ZqresumoPVr_rec.Value>0 then
             begin}
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('valor').Value;//+ZqresumoPJuros_Vr.Value;
               parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoP.FieldByName('valor').Value;//+ZqresumoPJuros_Vr.Value;
{             end
             else
             begin
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoPvr_ab.Value;
               parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoPvr_ab.Value;
             end;
          end;}
        end
        else
        begin
          {if ZqresumoPvr_rec.Value>ZqresumoPvalor_1.Value then
          begin
             xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoPvalor_1.Value+ZqresumoPJuros_Vr.Value;
             parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoPvalor_1.Value+ZqresumoPJuros_Vr.Value;
          end
          else
          begin
             if ZqresumoPVr_rec.Value>0 then
             begin}
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('valor').Value;//+ZqresumoPJuros_Vr.Value;
               parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoP.FieldByName('valor').Value;//+ZqresumoPJuros_Vr.Value;
{             end
             else
             begin
               xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoPvr_ab.Value;
               parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoPvr_ab.Value;
             end;
          end;}
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
      FRM_Recebi03.RLLabel19.Caption:='Observa��es';
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
{    DM_Tabelas.CDSEmpreendimento.Close;
    DM_Tabelas.CDSEmpreendimento.CreateDataSet;
    DM_Tabelas.ZQLoteamento.First;
    JSP1.visible:=true;
    JSP1.maximum:=DM_Tabelas.ZQLoteamento.recordcount;
    JSP1.position:=0;
    While not DM_Tabelas.ZQLoteamento.Eof do
    Begin
      JSP1.position:=DM_Tabelas.ZQLoteamento.recno;
      application.ProcessMessages;
      while not DM_Tabelas.ZQincorp_loteame.Eof do
      begin
        if (pos(DM_Tabelas.ZQincorp_loteameincorporador_idincorporador.Text,Varaleatorio)>0) or empty(Varaleatorio) Then Begin
          DM_Tabelas.CDSEmpreendimento.Insert;
          DM_TAbelas.CDSEmpreendimentoidordem.Value := DM_Tabelas.CDSEmpreendimento.RecordCount;
          DM_Tabelas.CDSEmpreendimentoidloteamento.Value := DM_Tabelas.ZQLoteamentoidloteamento.Value;
          DM_Tabelas.CDSEmpreendimentocodcontabancaria.Value := DM_Tabelas.ZQincorp_loteamecodcontabancaria.Value;
          DM_Tabelas.CDSEmpreendimentoidparticipante.Value := DM_Tabelas.ZQincorp_loteameincorporador_idincorporador.Value;
          DM_Tabelas.CDSEmpreendimentodigito_dif.Value := DM_Tabelas.ZQincorp_loteamedigito_dif.Value;
          DM_Tabelas.CDSEmpreendimentomarca.Value := '0';
          DM_Tabelas.CDSEmpreendimento.Post;
        end;
        DM_Tabelas.ZQincorp_loteame.Next;
      end;
      DM_TAbelas.ZQLoteamento.Next;
    end;}


    CDSEmpreendimento.Close;
    CDSEmpreendimento.CreateDataSet;
    CDSEmpreendimento.open;

    ZQincorp_loteame.close;
    ZQincorp_loteame.SQL.Clear;
    ZQincorp_loteame.SQL.Add('Select * from incorporador_loteamento as i join loteamento as l on i.loteamento_idloteamento=l.idloteamento where 1=1 ');
    if not empty(VarEmpree) then
       ZQincorp_loteame.SQL.Add(' and i.loteamento_idloteamento in ('+VarEmpree+')');
    ZQincorp_loteame.open;
    
    ZQincorp_loteame.First;
    JSP1.visible:=true;
    JSP1.maximum:=ZQincorp_loteame.recordcount;

    JSP1.position:=0;
    While not ZQincorp_loteame.Eof do
    Begin
      JSP1.position:=ZQincorp_loteame.recno;
      application.ProcessMessages;
      if (pos(ZQincorp_loteame.FieldByName('incorporador_idincorporador').Text,Varaleatorio)>0) or empty(Varaleatorio) Then
      Begin
          CDSEmpreendimento.Insert;
          CDSEmpreendimentoidordem.Value := CDSEmpreendimento.RecordCount;
          CDSEmpreendimentoidloteamento.Value := ZQincorp_loteame.FieldByName('idloteamento').Value;
          CDSEmpreendimentocodcontabancaria.Value := ZQincorp_loteame.FieldByName('codcontabancaria').Value;
          CDSEmpreendimentoidparticipante.Value := ZQincorp_loteame.FieldByName('incorporador_idincorporador').Value;
          CDSEmpreendimentodigito_dif.Value := ZQincorp_loteame.FieldByName('digito_dif').Value;
          CDSEmpreendimentomarca.Value := '0';
          CDSEmpreendimento.Post;
      end;
      ZQincorp_loteame.Next;
    end;
    FrmCobrancaBancaria.CDSEmpreendimento.Close;
    FrmCobrancaBancaria.CDSEmpreendimento.CreateDataSet;
    FrmCobrancaBancaria.CDSEmpreendimento.open;
    FrmCobrancaBancaria.CDSEmpreendimento.CloneCursor(CDSEmpreendimento, true, true);
    FrmCobrancaBancaria.CDSEmpreendimento.Data:=CDSEmpreendimento.Data;

    JSP1.position:=0;
    JSP1.visible:=false;
    application.ProcessMessages;

    ZQRecebParticp.close;
    ZQRecebParticp.SQL.Clear;
    ZQRecebParticp.SQL.Add('select r.idrecebimento,r.adversa,r.cliente,r.documento,r.idloteamento,r.quadralote,');
    ZQRecebParticp.SQL.Add('r.marcar,r.numboleto,r.Dt_Entrada,r.Dt_Vencimento,r.Valor,r.ordem,r.TipDoc,r.numordem,');
    ZQRecebParticp.SQL.Add('r.saldo,r.origem,r.recpag,r.venda_idvenda,r.substituicao,r.sq,p.idpaticipante,p.nome_parte,');
    ZQRecebParticp.SQL.Add('p.doc1,p.endereco,p.bairro,p.cidade,p.cep,p.ende_cob,p.bairro_cob,p.cidade_cob,p.cep_cob,');
    ZQRecebParticp.SQL.Add('p.fone1,p.naturalidade');
    ZQRecebParticp.SQL.Add(' from recebimento  as r join participante as p on r.adversa=p.idpaticipante where r.recpag = '+quotedstr('R'));
    if CBEntrada.Checked <> CBParcela.Checked Then Begin
      if CBParcela.Checked Then
        ZQRecebParticp.SQL.Add(' and documento like '+quotedstr('%-P-%'));
    end;
    if (not empty(vardevedor)) and (RGContas.ItemIndex=0) then
      ZQRecebParticp.SQL.Add(' and r.adversa in ('+vardevedor+')');


    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      ZQRecebParticp.SQL.Add(' and r.quadralote in ('+varquadra+')');


    if not empty(VarEmpree) then
       ZQRecebParticp.SQL.Add(' and r.idloteamento in ('+VarEmpree+')');

    if not empty(Varaleatorio) then
      ZQRecebParticp.SQL.Add(' and r.cliente in ('+Varaleatorio+')');
    if CBSaldo.ItemIndex = 1 Then
      ZQRecebParticp.SQL.Add(' and (r.saldo < r.valor or r.saldo is null)');
    if CBSaldo.ItemIndex = 2 Then
      ZQRecebParticp.SQL.Add(' and r.saldo > 0');
    if not empty(vardocumento) Then
      ZQRecebParticp.SQL.Add(vardocumento);
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQRecebParticp.SQL.Add(' and r.DT_Entrada between :dt1 and :dt2');
      ZQRecebParticp.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQRecebParticp.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQRecebParticp.SQL.Add(' and r.DT_Vencimento between :dt3 and :dt4');
      ZQRecebParticp.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQRecebParticp.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    ZQRecebParticp.SQL.Add(' and r.TipDoc in ('+varDoc+')');

    if CBDesconsidera.Checked Then
      ZQRecebParticp.SQL.Add(' and not exists (select * from remessa_receb where idrec=idrecebimento)');

    ZQRecebParticp.Open;


    if ZQRecebParticp.RecordCount>0 then
       JSP1.Maximum := ZQRecebParticp.RecordCount;

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
      CDS_MarcaTit.open;
    end;

    JSP1.visible:=true;
    JSP1.maximum:=ZQRecebParticp.recordcount;
    JSP1.position:=0;


    while not ZQRecebParticp.Eof do
    begin
      JSP1.Position := ZQRecebParticp.RecNo;
      application.ProcessMessages;
      CDS_MarcaTit.Insert;
      CDS_MarcaTitIdreceb.Value := ZQRecebParticp.FieldByName('idrecebimento').Value;
      CDS_MarcaTitidparti.Value := ZQRecebParticp.FieldByName('adversa').Value;
      CDS_MarcaTitcliente.Value := ZQRecebParticp.FieldByName('cliente').Value;
      CDS_MarcaTitidloteam.Value := ZQRecebParticp.FieldByName('idloteamento').Value;
      CDS_MarcaTitquadralote.Value:=ZQRecebParticp.FieldByName('quadralote').Value;
      CDS_MarcaTitmarca.Value := '0';
      CDS_MarcaTitnumboleto.Value:= ZQRecebParticp.FieldByName('numboleto').Value;
      CDS_MarcaTitDT_Entrada.value:= ZQRecebParticp.FieldByName('Dt_Entrada').Value;
      CDS_MarcaTitDt_Vencimento.value:= ZQRecebParticp.FieldByName('Dt_Vencimento').Value;
      CDS_MarcaTitValor.value:= ZQRecebParticp.FieldByName('Valor').Value;
      CDS_MarcaTitordem.value:= ZQRecebParticp.FieldByName('ordem').Value;
      CDS_MarcaTitnome_parte.value:= ZQRecebParticp.FieldByName('nome_parte').Value;
      CDS_MarcaTitende_cob.value:= ZQRecebParticp.FieldByName('ende_cob').Value;
      CDS_MarcaTitcep_cob.value:= ZQRecebParticp.FieldByName('cep_cob').Value;
      CDS_MarcaTitdoc1.value:= ZQRecebParticp.FieldByName('doc1').Value;
      CDS_MarcaTitTipDoc.value:= ZQRecebParticp.FieldByName('TipDoc').Value;
      CDS_MarcaTitnumordem.value:= ZQRecebParticp.FieldByName('numordem').Value;
      CDS_MarcaTitbairro.value:=ZQRecebParticp.FieldByName('bairro_cob').Value;
      CDS_MarcaTitcidade.value:=ZQRecebParticp.FieldByName('cidade_cob').Value;
      posi:=pos('-',ZQRecebParticp.FieldByName('ordem').Value);
      inc(posi);
      CDS_MarcaTitordem.Value:= copy(ZQRecebParticp.FieldByName('ordem').Value,posi,15);
      CDS_MarcaTit.Post;
      ZQRecebParticp.Next;
    end;
    JSP1.position:=0;
    JSP1.visible:=false;
    application.ProcessMessages;



    dxbSelecionar.enabled:=false;
    CDS_MarcaTit.IndexName:='nome_venc';
    FrmCobrancaBancaria.Top := FrmRelRecebimento.Top + 40;
    FrmCobrancaBancaria.Left :=FrmRelRecebimento.Left+30;
    FrmCobrancaBancaria.showmodal;
    CDS_MarcaTit.IndexName:='';


//    dxbSelecionar.Enabled:=true;
//    dxbselecionar.SetFocus;
    Exit;
  end;
  vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
  if not empty(XDEVencimentoInicio.Text) Then
    vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' � '+XDEVencimentoFinal.Text;
  if not empty(XDEEntradaInicio.Text) Then
    vartitulo := vartitulo + ', com Emiss�o de '+XDEEntradaInicio.Text+' � '+XDEEntradaFinal.Text;
  if not empty(XDEBaixasInicio.Text) Then
    vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' � '+XDEBaixasFinal.Text;
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
    ZQmensal.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Mar�o''');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQmensal.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQmensal.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQmensal.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
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
      FrmRelReceb02_mensal_ab.RLLabel2.Caption := 'Per�odo de : '+vrtexto;
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;      
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZqresumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
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
      FrmRelReceb02.RLLabel2.Caption := 'Per�odo de : '+vrtexto;
    FrmRelReceb02.RLReport1.Preview;
    ZQEmpree.Close;
    vrtexto:='';
   end
   else
   begin
     vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
     if not empty(XDEVencimentoInicio.Text) Then
        vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' � '+XDEVencimentoFinal.Text;
     if not empty(XDEEntradaInicio.Text) Then
        vartitulo := vartitulo + ', com Emiss�o de '+XDEEntradaInicio.Text+' � '+XDEEntradaFinal.Text;
     if not empty(XDEBaixasInicio.Text) Then
        vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' � '+XDEBaixasFinal.Text;
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
       Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' � '+XDEEntradaFinal.DateText;
     end;
     if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
     Begin
{      ZqresumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);}
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' � '+XDEVencimentoFinal.DateText;
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
        FrmRelReceb02_total.RLLabel2.Caption := 'Per�odo de : '+vrtexto;
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
      CDSDetaldetalhe.Value := 'Lan�: '+ZQTempReceber.FieldByName('idrecebimento').Text+'  Participante: '+ZQTempReceber.FieldByName('nome_parte').Value+'      Adversa: '+ZQTempReceber.FieldByName('adversanome').Value+chr(13);
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

procedure TFrmRelRecebimento.XDEEntradaInicioExit(Sender: TObject);
begin
  if XDEBaixasfinal.DateValue<XDEBaixasinicio.DateValue Then
     XDEBaixasfinal.DateValue := XDEBaixasinicio.DateValue;
end;

procedure TFrmRelRecebimento.XDEVencimentoInicioExit(Sender: TObject);
begin
  if XDEVencimentoFinal.DateValue<XDEVencimentoInicio.DateValue Then
     XDEVencimentoFinal.DateValue := XDEVencimentoInicio.DateValue;
end;

procedure TFrmRelRecebimento.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    if not (ActiveControl is TDBGrid) then begin
      Key := #0;
      Perform(WM_NEXTDLGCTL, 0, 0);
    end;
end;

procedure TFrmRelRecebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQBaixa.Close;
  ZQTempReceber.Close;
  ZQTempReceber2.close;
  DM_Tabelas.ZQincorp_loteame.close;
  DM_Tabelas.ZQLoteamento.close;
  ZQincorp_loteame.close;
  ZQLoteamento.close;
  ZQRecebParticp.close;

end;

procedure TFrmRelRecebimento.GBEscClienteExit(Sender: TObject);
begin
  FrmRelRecebimento.KeyPreview := True
end;

procedure TFrmRelRecebimento.ZQBaixaCalcFields(DataSet: TDataSet);
begin
  if ZQBaixa.FieldByName('percent_usado').Value >0 Then
    ZQBaixa.FieldByName('vrusado').Value := ZQBaixa.FieldByName('Vr_rec').Value*ZQBaixa.FieldByName('percent_usado').Value/100
  else
    ZQBaixa.FieldByName('vrusado').Value := ZQBaixa.FieldByName('Vr_rec').Value;
end;

procedure TFrmRelRecebimento.EClienteExit(Sender: TObject);
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

procedure TFrmRelRecebimento.DBGrid3KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSCli.RecordCount>0) Then Begin
    CDSCli.Delete;
  end;
end;

procedure TFrmRelRecebimento.DBGrid4KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;
end;

procedure TFrmRelRecebimento.RGContasClick(Sender: TObject);
begin
  if RGContas.ItemIndex = 0 then
    GBAdversa.Enabled := True
  else
    GBAdversa.Enabled := False;
end;

procedure TFrmRelRecebimento.RGFormatoClick(Sender: TObject);
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
    CbCarteira.Visible:=false;
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
    CbCarteira.Visible:=false;
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
    CbCarteira.Visible:=false;
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
    CbCarteira.Visible:=true;
    xresumo.Visible:=true;
  end;
  if RGFormato.ItemIndex = 4 Then
  Begin
    CBSaldo.ItemIndex:=2;
    CBVerbaixa.ItemIndex:=1;
    GroupBox6.Visible:=false;
    GroupBox8.Visible:=false;
    cbmensal.Visible:=false;
    CbCarteira.Visible:=false;    
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


procedure TFrmRelRecebimento.EloteExit(Sender: TObject);
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
      Showmessage('Esta Quadra e lote n�o foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end
  else
    if CDSQuadraLote.RecordCount>0 Then
      DXBImprimir.SetFocus;
end;

procedure TFrmRelRecebimento.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;
end;

procedure TFrmRelRecebimento.XDERefInicioExit(Sender: TObject);
begin
  if XDERefFinal.DateValue<XDERefInicio.DateValue Then
     XDERefFinal.DateValue := XDERefInicio.DateValue;
end;

procedure TFrmRelRecebimento.dxbSelecionarClick(Sender: TObject);
begin
{    dxbSelecionar.enabled:=false;
    CDS_MarcaTit.IndexName:='nome_venc';
    FrmCobrancaBancaria.Top := FrmRelRecebimento.Top + 40;
    FrmCobrancaBancaria.Left :=FrmRelRecebimento.Left+30;
    FrmCobrancaBancaria.showmodal;
    CDS_MarcaTit.IndexName:='';}    
end;

procedure TFrmRelRecebimento.CBSaldoExit(Sender: TObject);
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

procedure TFrmRelRecebimento.Edit1Exit(Sender: TObject);
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


procedure TFrmRelRecebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQBaixa', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaTipdoc_1', 'Tipdoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadescricao', 'descricao', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixatipodoc', 'tipodoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavend_receb', 'vend_receb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavend_caixa', 'vend_caixa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixareceb_caixa', 'receb_caixa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixareceb_receb', 'receb_receb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadados_chequ', 'dados_chequ', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaso_avista', 'so_avista', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixalancabanco', 'lancabanco', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasomapaga', 'somapaga', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixalancinclus', 'lancinclus', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavrusado', 'vrusado', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidtipodocumento', 'idtipodocumento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixanumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQResumoReceb', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebtipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoReceblocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQSomaBaixa', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasomabaixa', 'somabaixa', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixanumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQAgruTipo', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoidRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiposaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiposomasaldo', 'somasaldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiporecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiponumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiporefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiponumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTemp2Cl', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cldoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cldoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cltipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Claniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cllocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRateio', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiotipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateionumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreematriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreebairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreepasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreenomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQVenda', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaquadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendalote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendalado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendacornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendadisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQVen1', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1quadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1lote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1lado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomatit', 'somatit', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomapago', 'somapago', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomajuros', 'somajuros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsubstituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb1', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Vr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Docum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Juros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Desc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1obsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Codcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1vencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1refer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1valor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1descon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1juros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1percent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1data', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Valor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1RefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1numboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1substituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Substituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1vr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1sq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1sq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1TipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1sq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1somar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1valor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1valor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Reajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Data_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Proximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Parcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQCh', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChbanco', 'banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChagencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdono', 'dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChCPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChalias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChconta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChsubstituicao_2', 'substituicao_2', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTempReceber2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2planomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2adversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2nomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2numboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Substituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2tvr', 'tvr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2tsld', 'tsld', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2nomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQmensal', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalTIPDOC', 'TIPDOC', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalvr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalnomelote', 'nomelote', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalMESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalvr_rel', 'vr_rel', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQmensal2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2MESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2vr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2tipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2DT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2DT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2vr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2juros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRelatorio2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2MESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2vr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2tipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2DT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2DT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2valorpago', 'valorpago', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRelatorio', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioMESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriovr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriotipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriovr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriovalorpago', 'valorpago', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriodataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriovalorpagojuros', 'valorpagojuros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriojuros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZqresumoE', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoETipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoERefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoECodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoESubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoErefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoETipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoERefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoErefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoErecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoESubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZqresumoP', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPnumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPnomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQcarteira', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraMESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteirarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiratipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiradataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteirajuros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraap', 'ap', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRelCarteira', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraMESDESC', 'MESDESC', TStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiravr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteirarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiratipdoc', 'tipdoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiradataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiravr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteirajuros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiravalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiradescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteirajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiravalorpago', 'valorpago', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiravalorpagojuros', 'valorpagojuros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraap', 'ap', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZRdiario', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariorefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarionumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarionomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariojuros_1', 'juros_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioApelido', 'Apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQincorp_loteame', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedigito_dif', 'digito_dif', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamenconta', 'nconta', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameconta_v', 'conta_v', TStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta_v', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameagencia', 'agencia', TStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameagencia_v', 'agencia_v', TStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia_v', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamen_banco', 'n_banco', TStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_banco', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamematriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamebairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamepasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamenomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria_1', 'codcontabancaria_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedigito_dif_1', 'digito_dif_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRecebParticp', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpnumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticporigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticprecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebParticp', 'ZQRecebParticpnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQLoteamento', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentoidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentomatriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentopasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentotestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentotestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentotestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentotestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentoObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentobairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentotipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentonomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQLoteamento', 'ZQLoteamentoapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQCidade', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQCidade', 'ZQCidadeidcidade', 'idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQCidade', 'ZQCidadenomecid', 'nomecid', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQCidade', 'ZQCidadeddd', 'ddd', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQCidade', 'ZQCidadeemancipacao', 'emancipacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQCidade', 'ZQCidadeestado', 'estado', TStringField, fkData, 2, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQCidade', 'ZQCidadecepgeral', 'cepgeral', TStringField, fkData, 10, 0, False, '', '', '', '99\.999\-999;1;_', 0, '', '', '', '', False);

end.
