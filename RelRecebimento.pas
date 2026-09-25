unit RelRecebimento;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, XEdit, Mask, XDate, ExtCtrls, XBanner,
  XLabel3D, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, JvExControls, JvSpecialProgress, Grids, DBGrids, DBClient,
  DBCtrls, TFlatCheckBoxUnit, ActnList, IniFiles, XNum, Provider,
  System.Actions,system.Threading;

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
    DXBFechar: TdxButtonArround;
    DXBImprimir: TdxButtonArround;
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
    CDSCliRcli: TWideStringField;
    DS_Adversatemp: TDataSource;
    ZQAdvsersatemp: TZQuery;
    DS_Adversatemp2: TDataSource;
    CDSAdversaTemp2: TClientDataSet;
    CDSAdversaTemp2CodAdversa: TIntegerField;
    CDSAdversaTemp2nomeadversa: TWideStringField;














    DS_Plano: TDataSource;
    CDSPlano: TClientDataSet;
    CDSPlanomascara: TWideStringField;
    CDSPlanodescricao: TWideStringField;
    CDSPlanoCrDe: TWideStringField;
    CDSPlanosoma: TFloatField;
    CDSPlanoclassificacao: TWideStringField;
    CDSPlanocodigo: TIntegerField;
    DS_Detal: TDataSource;
    CDSDetal: TClientDataSet;
    CDSDetalmascara: TWideStringField;
    CDSDetaldata: TDateField;
    CDSDetaldetalhe: TWideMemoField;
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
    gbfinanc: TGroupBox;
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
    CDSQuadraLoteQuadra: TWideStringField;
    CDSQuadraLotelote: TWideStringField;
    CDSQuadraLoteidlote: TIntegerField;
    CDSQuadraLoteidloteamen: TIntegerField;
    DS_Ven1: TDataSource;
    ZQVen1: TZQuery;

























    DS_ROReceb: TDataSource;
    ZROQReceb: TZReadOnlyQuery;














































    DS_Tipodoc: TDataSource;
    CDSTipoDoc: TClientDataSet;
    CDSTipoDocTipoDoc: TWideStringField;
    CDSTipoDocdescDoc: TWideStringField;
    CDSTipoDocdadosch: TWideStringField;
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
    QReceb1TipDoc: TWideStringField;
    QReceb1Docum: TWideStringField;
    QReceb1Juros_Vr: TFloatField;
    QReceb1Desc_Vr: TFloatField;
    QReceb1RefBaixa: TIntegerField;
    QReceb1obsebx: TWideMemoField;
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
    QReceb1documento: TWideStringField;
    QReceb1cliente: TIntegerField;
    QReceb1usuario: TIntegerField;
    QReceb1Dt_Entrada: TDateField;
    QReceb1Dt_Vencimento: TDateField;
    QReceb1Valor_1: TFloatField;
    QReceb1Observ: TWideMemoField;
    QReceb1VrDoc: TFloatField;
    QReceb1ordem: TWideStringField;
    QReceb1TipDoc_1: TWideStringField;
    QReceb1saldo: TFloatField;
    QReceb1marcar: TWideStringField;
    QReceb1RefBaixa_1: TIntegerField;
    QReceb1refvinda: TIntegerField;
    QReceb1contabil: TIntegerField;
    QReceb1empresa: TIntegerField;
    QReceb1custodaparcela: TFloatField;
    QReceb1origem: TWideStringField;
    QReceb1adversa: TIntegerField;
    QReceb1recpag: TWideStringField;
    QReceb1numordem: TIntegerField;
    QReceb1idloteamento: TIntegerField;
    QReceb1venda_idvenda: TIntegerField;
    QReceb1quadralote: TWideStringField;
    QReceb1numboleto: TWideStringField;
    QReceb1nomeadversa: TWideStringField;
    QReceb1nomecliente: TWideStringField;
    QReceb1nomeempree: TWideStringField;

    dxbx: TRadioGroup;















    DS_MarcaTit: TDataSource;

    dxbSelecionar: TdxButtonArround;
    stotais: TCheckBox;

    DataZQTempReceber2: TDataSource;
    ZQTempReceber2: TZQuery;















































    Datamensal: TDataSource;
    ZQmensal: TZQuery;
    cbmensal: TCheckBox;












    ZQmensal2: TZQuery;
    Dzqmensal2: TDataSource;
    ActionList1: TActionList;














    XBanner12: TXBanner;
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
    GBParcelas: TGroupBox;
    XBanner20: TXBanner;
    CBFinanciamento: TComboBox;
    CBfinan: TCheckBox;
    zqfinanciamento: TZReadOnlyQuery;
    DS_financiamento: TDataSource;

























































    DSfinan: TDataSource;
    CDSEmpreendimentocod_transmissao: TStringField;
    CDSEmpreendimentocomplemento: TStringField;


    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;








































    CBAgrupado: TCheckBox;
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
    CDS_MarcaTitquadralote: TStringField;
    CDS_MarcaTitdigito_dif: TStringField;
    CDS_MarcaTitbairro: TStringField;
    CDS_MarcaTitdoc2: TStringField;
    CDS_MarcaTitcidade: TLargeintField;
    CDS_MarcaTitnomecidade: TStringField;
    CDS_MarcaTituf: TStringField;
    CDS_MarcaTitnumboleto: TStringField;
    CDS_finan: TClientDataSet;
    CDS_finanidrecebimento: TLargeintField;
    CDS_finandocumento: TStringField;
    CDS_financliente: TIntegerField;
    CDS_finanDt_Entrada: TDateField;
    CDS_finanDt_Vencimento: TDateField;
    CDS_finanordem: TStringField;
    CDS_finansaldo: TFloatField;
    CDS_finanidloteamento: TIntegerField;
    CDS_finanquadralote: TStringField;
    CDS_finannomeadversa: TStringField;
    CDS_finanapelido: TStringField;




    CBQuadra: TCheckBox;
    cbtotais: TCheckBox;
    ZRTotPago: TZReadOnlyQuery;
    DS_Zrtotpago: TDataSource;

































































    DXBMarcDoc: TdxButtonArround;
    dbxmarcempre: TdxButtonArround;
    Eadversa: TEdit;

    CDSTipoDocsomadescontos: TFloatField;

    CBEmpre: TCheckBox;
    xquadraLote: TMemo;
    RGQuadraLote: TRadioGroup;
    procedure FormShow(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DXBImprimirClick(Sender: TObject);
    procedure XDEEntradaInicioExit(Sender: TObject);
    procedure XDEVencimentoInicioExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GBEscClienteExit(Sender: TObject);
    procedure ZQBaixaCalcFields(DataSet: TDataSet);
    procedure ZQRecebimentoCalcFields(DataSet: TDataSet);
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
    procedure CBfinanClick(Sender: TObject);
    procedure cbtotaisClick(Sender: TObject);
    procedure DXBMarcDocClick(Sender: TObject);
    procedure dbxmarcempreClick(Sender: TObject);
    procedure EadversaExit(Sender: TObject);
    procedure CBQuadraClick(Sender: TObject);
    procedure RGQuadraLoteClick(Sender: TObject);
    procedure xquadraLoteExit(Sender: TObject);
  private
    { Private declarations }

    procedure PrepararLookupsReajuste;
    procedure PrepararNomesTempReceber;

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
  rel_diario_lote, rel_financiamento, Relajuste2, RelReceb04, RelReceb05,
  RelReceb02_total_ab, Recebitotpago, RelReceb02_total_ab_lt, uRuntimeFields, uSiaiPerformance;

{$R *.dfm}

function ListaDoc: TStrings;
begin
  Result := TStringList.Create;
  Result.Clear;
  Result.BeginUpdate;
  DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQLoteamento.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
            Result.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
            DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Result.EndUpdate;
end;

function ListaEmpree: TStrings;
begin
  Result := TStringList.Create;
  Result.Clear;
  Result.BeginUpdate;
  DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQLoteamento.First;
  DM_TAbelas.ZQLoteamento.DisableControls;
  while not DM_TAbelas.ZQLoteamento.Eof do begin
        Result.Add(DM_Tabelas.ZQLoteamento.FieldByName('apelido').AsString+' - '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text);
        DM_Tabelas.ZQLoteamento.Next;
  end;
  DM_TAbelas.ZQLoteamento.EnableControls;
  Result.EndUpdate;
end;

procedure TFrmRelRecebimento.PrepararNomesTempReceber;
var
  LCampo: TWideStringField;
begin
  { Campos SQL independentes dos buffers calculados dos lookups. As subconsultas
    preservam a quantidade de titulos e os filtros da consulta principal. }
  ZQTempReceber.SQL.Add(' ,(select L.apelido from loteamento L');
  ZQTempReceber.SQL.Add('   where L.idloteamento = RE.idloteamento) as nomeempreend_relatorio');
  ZQTempReceber.SQL.Add(' ,COALESCE((select P.nome_parte from participante P');
  ZQTempReceber.SQL.Add('   where P.idpaticipante = RE.adversa), RE.nomeadversa) as adversanome_relatorio');

  if ZQTempReceber.FindField('nomeempreend_relatorio') = nil then
  begin
    LCampo := TWideStringField.Create(ZQTempReceber);
    LCampo.FieldName := 'nomeempreend_relatorio';
    LCampo.Size := 100;
    LCampo.DataSet := ZQTempReceber;
  end;
  if ZQTempReceber.FindField('adversanome_relatorio') = nil then
  begin
    LCampo := TWideStringField.Create(ZQTempReceber);
    LCampo.FieldName := 'adversanome_relatorio';
    LCampo.Size := 140;
    LCampo.DataSet := ZQTempReceber;
  end;
end;

procedure TFrmRelRecebimento.PrepararLookupsReajuste;
begin
  if not DM_Tabelas.ZqParticipante.Active then
    DM_Tabelas.ZqParticipante.Open;
  if not DM_Tabelas.ZQLoteamento.Active then
    DM_Tabelas.ZQLoteamento.Open;
end;

procedure TFrmRelRecebimento.ZQRecebimentoCalcFields(DataSet: TDataSet);
var
  LChave: TField;
  LDestino: TField;
  LOrigem: TField;
  LValor: Variant;
begin
  if (DataSet = nil) or (DM_Tabelas = nil) then
    Exit;

  LChave := DataSet.FindField('adversa');
  LDestino := DataSet.FindField('adversanome');
  if LDestino <> nil then
  begin
    LDestino.Clear;
    if (LChave <> nil) and (not LChave.IsNull) and
       (LChave.AsLargeInt <> 0) and DM_Tabelas.ZqParticipante.Active then
    begin
      LValor := DM_Tabelas.ZqParticipante.Lookup(
        'idpaticipante', LChave.AsLargeInt, 'nome_parte');
      if not VarIsNull(LValor) and not VarIsEmpty(LValor) then
        LDestino.AsString := VarToStr(LValor);
    end;
    if LDestino.AsString = '' then
    begin
      LOrigem := DataSet.FindField('nomeadversa');
      if (LOrigem <> nil) and (not LOrigem.IsNull) then
        LDestino.AsString := LOrigem.AsString;
    end;
  end;

  LChave := DataSet.FindField('idloteamento');
  LDestino := DataSet.FindField('nome_loteamento');
  if (LDestino <> nil) and (LChave <> nil) and (not LChave.IsNull) and
     (LChave.AsLargeInt <> 0) and DM_Tabelas.ZQLoteamento.Active then
  begin
    LDestino.Clear;
    LValor := DM_Tabelas.ZQLoteamento.Lookup(
      'idloteamento', LChave.AsLargeInt, 'apelido');
    if not VarIsNull(LValor) and not VarIsEmpty(LValor) then
      LDestino.AsString := VarToStr(LValor);
  end;
end;

procedure TFrmRelRecebimento.FormShow(Sender: TObject);
Var
  VarPath : string;
  ArqIni2 : tIniFile;
  ListaTemporaria: TStrings;
  Task : Itask;
  I:integer;
  LInicio, LStep: UInt64;
begin
  LInicio := PerformanceStart;
  LStep := PerformanceStart;
  PrepararLookupsReajuste;
  PerformanceElapsed('RelRecebimento: preparar lookups', LStep);
  xquadraLote.Lines.Clear;
  VarPath := ExtractFilePath( Application.ExeName );
  ArqIni2 := tIniFile.Create(varpath+'siai.Ini');
  try
    varschemata := ArqIni2.ReadString('BANCO_DE_DADOS','SCHEMATA', varschemata );
  finally
    ArqIni2.Free;
  end;


  if (RGFormato.ItemIndex = 4) or (RGFormato.ItemIndex = 2) Then
    CBDesconsidera.Visible := True
  else
    CBDesconsidera.Visible := False;
  LStep := PerformanceStart;
  ZQTemp2Cl.Open;
  PerformanceElapsed('RelRecebimento: abrir filtro temporario', LStep);
  CDSQuadraLote.Close;
  CDSQuadraLote.CreateDataSet;
  CDSCli.Close;
  CDSCli.CreateDataSet;
  CDSAdversaTemp2.Close;
  CDSAdversaTemp2.CreateDataSet;
  CLBDoc.Sorted := False;
  DM_tabelas.ZQTipodoc.First;
  CLBEmpree.Clear;
  ListaTemporaria := ListaEmpree;
  try
    CLBEmpree.Items.Assign(ListaTemporaria);
  finally
    ListaTemporaria.Free;
  end;

  //CLBEmpree.Sorted:=true;

  //  i:=0;
//  for i := 0 to CLBEmpree.Items.Count -1 do
//  begin
//    CLBEmpree.Checked[I] := True;
//  end;

  LStep := PerformanceStart;
  DM_Tabelas.ZQincorp_loteame.open;
  PerformanceElapsed('RelRecebimento: abrir empreendimentos', LStep);
  CLBDoc.Clear;
  ListaTemporaria := ListaDoc;
  try
    CLBDoc.Items.Assign(ListaTemporaria);
  finally
    ListaTemporaria.Free;
  end;
  CLBDoc.Sorted := True;
  i:=0;
  for i := 0 to CLBDoc.Items.Count -1 do
  begin
    CLBDoc.Checked[I] := True;
  end;

  RGContas.SetFocus;
  PerformanceElapsed('RelRecebimento: preparar formulario', LInicio);
  FlushPerformanceLog;
end;

procedure TFrmRelRecebimento.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRelRecebimento.DXBImprimirClick(Sender: TObject);
var
  VarEmpree, VarDoc, VarOrdem, vartitulo, Varaleatorio, vardevedor, Vrtexto,
  p,dentromes,forames,data,d,m,a,entrada,vcto,baixa,varmasc, vardocumento, varquadra, varsql,VarLogo : String;
  ano1,ano2,mes1,mes2,varcontador, x, x2, y, varvezes, posi : Integer;
  QTeste: TZQuery;
  TemDuplicidade: Boolean;
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
  CDSQuadraLote.DisableControls;
  while not CDSQuadraLote.Eof do begin
    varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
    CDSQuadraLote.Next;
  end;
  CDSQuadraLote.EnableControls;
  varquadra := copy(varquadra,1,length(varquadra)-1);
  vardocumento := '';
  if CBEntrada.Checked <> CBParcela.Checked Then Begin
    if CBParcela.Checked Then
      vardocumento := ' and documento like '+quotedstr('%-P-%');
    if CBEntrada.Checked Then
      vardocumento := ' and documento like '+quotedstr('%-E-%');
  end;
  Varaleatorio := '';
  if CDSCli.Active Then Begin
    CDSCli.First;
    CDSCli.DisableControls;
    while not CDSCli.Eof do Begin
      Varaleatorio := Varaleatorio +quotedstr(CDSCliCcli.Text)+',';
      CDSCli.Next;
    end;
    CDSCli.EnableControls;
    Varaleatorio := copy(Varaleatorio,1,length(Varaleatorio)-1)
  end;
  vardevedor := '';
  if CDSAdversaTemp2.Active Then Begin
    CDSAdversaTemp2.First;
    CDSAdversaTemp2.DisableControls;
    while not CDSAdversaTemp2.Eof do Begin
      vardevedor := vardevedor +quotedstr(CDSAdversaTemp2CodAdversa.Text)+',';
      CDSAdversaTemp2.Next;
    end;
    CDSAdversaTemp2.EnableControls;
    vardevedor := copy(vardevedor,1,length(vardevedor)-1)
  end;
  if (RGFormato.ItemIndex=5)  then
  begin
    case RGOrdem.ItemIndex of
      0: VarOrdem := 'DT_Vencimento';
      1: VArOrdem := 'nomeadversa';
      2: VArOrdem := 'quadralote';
      3: VArOrdem := 'TipDoc';
      4: VArOrdem := 'dataref';
      5: VarOrdem := 'DT_Entrada';
      6: VarOrdem := 'idloteamento';
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
      6: VarOrdem := 'idloteamento';
    end;
  end;

  if RGFormato.ItemIndex=6 then
  begin
    case RGOrdem.ItemIndex of
      0: VarOrdem := 'DT_Vencimento';
      1: VArOrdem := 'nomeadversa';
      2: VArOrdem := 'quadralote';
      3: VArOrdem := 'TipDoc';
      4: VarOrdem := 'idloteamento';
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



  QTeste := TZQuery.Create(nil);
  try
    QTeste.Connection := DM_Tabelas.zconeccao;

    // verifica duplicidade
    QTeste.SQL.Clear;
    QTeste.SQL.Add('SELECT RH.idrecib');
    QTeste.SQL.Add('FROM receb_baixa RB');
    QTeste.SQL.Add('JOIN recbxhist RH ON RH.refer = RB.refbaixa');
    QTeste.SQL.Add('JOIN recebimento RE ON RH.idrecib = RE.idrecebimento');

    QTeste.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');
    if not empty(vardevedor) then
      QTeste.SQL.Add(' and adversa in ('+vardevedor+')');

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      QTeste.SQL.Add(' and quadralote in ('+varquadra+')');
    if RGContas.ItemIndex = 0 Then
      QTeste.SQL.Add(' and re.recpag ='+quotedstr('R'))
    else
      QTeste.SQL.Add(' and re.recpag ='+quotedstr('P'));
    if not empty(Varaleatorio) then
      QTeste.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      QTeste.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      QTeste.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      QTeste.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      QTeste.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    End;

    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      QTeste.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      QTeste.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      QTeste.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      QTeste.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      QTeste.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      QTeste.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      QTeste.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      QTeste.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      QTeste.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    End;
    QTeste.SQL.Add('GROUP BY RH.idrecib');
    QTeste.SQL.Add('HAVING COUNT(*) > 1');
    QTeste.Open;
    TemDuplicidade := not QTeste.IsEmpty;

  finally
    QTeste.Close;
    QTeste.Free;
  end;


  // total de parcelas do financiamento
  if (CBfinan.Checked=true)  then
  begin
    p:='%';
    zqfinanciamento.close;
    zqfinanciamento.SQL.Clear;
    zqfinanciamento.SQL.Add('Select RE.idrecebimento,RE.documento,RE.cliente,RE.usuario,RE.Dt_Entrada,RE.Dt_Vencimento,RE.Valor,RE.Observ,RE.VrDoc,RE.ordem,RE.TipDoc,RE.saldo,RE.marcar,');
    zqfinanciamento.SQL.Add('       RE.RefBaixa,RE.refvinda,RE.contabil,RE.empresa,RE.custodaparcela,RE.origem,RE.adversa,RE.recpag,RE.numordem,RE.idloteamento,RE.venda_idvenda,RE.quadralote,');
    zqfinanciamento.SQL.Add('       RE.numboleto,RE.Substituicao,RE.sq,RE.somar,RE.nomeadversa,RE.Reajustado,RE.Data_reajuste,RE.Proximo_Reajuste,RE.Parcelas_fixas,RE.juros,RE.descontos,RE.Data_Quitacao,');
    zqfinanciamento.SQL.Add('       lt.idloteamento,lt.cidade_idcidade,lt.matriculaloteamento,lt.datacadastro,lt.datainalguracao,lt.dataconclusao,lt.logotipo,lt.mapa,lt.testemunha1,lt.testemunha2,');
    zqfinanciamento.SQL.Add('       lt.testemunha3,lt.testemunha4,lt.Observincorp,lt.bairro,lt.pasta_mapa,lt.nomeloteamento,lt.tipo,lt.codcontabancaria,lt.apelido,lt.digito_dif ');
    zqfinanciamento.SQL.Add(' FROM recebimento as RE join loteamento as lt on lt.idloteamento=RE.idloteamento ');
  //  zqfinanciamento.SQL.Add(' SELECT * FROM recebimento as RE join recbxhist as RH on RH.idrecib=RE.idrecebimento ');
//    zqfinanciamento.SQL.Add(' join receb_baixa as RB on RH.refer=RB.refbaixa');
//    zqfinanciamento.SQL.Add(' join loteamento as lt on lt.idloteamento=RE.idloteamento');

    if not empty(CBFinanciamento.text) then
       zqfinanciamento.SQL.Add(' where RE.ordem like '+quotedstr(p+trim(CBFinanciamento.Text))+' and RE.documento like ''%-P-%'' and RE.TIPDOC in ('+varDoc+')')
    else
       zqfinanciamento.SQL.Add(' where RE.documento like ''%-P-%'' and RE.TIPDOC in ('+varDoc+')');
    if not empty(vardevedor) then
      zqfinanciamento.SQL.Add(' and RE.adversa in ('+vardevedor+')');

    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      zqfinanciamento.SQL.Add(' and RE.quadralote in ('+varquadra+')');
    if RGContas.ItemIndex = 0 Then
      zqfinanciamento.SQL.Add(' and RE.recpag ='+quotedstr('R'))
    else
      zqfinanciamento.SQL.Add(' and RE.recpag ='+quotedstr('P'));
    if not empty(Varaleatorio) then
      zqfinanciamento.SQL.Add(' and RE.cliente in ('+Varaleatorio+')');
    x2:=length(VarEmpree);
    if not empty(VarEmpree) then
      zqfinanciamento.SQL.Add(' and RE.idloteamento in ('+VarEmpree+')');

{    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      zqfinanciamento.SQL.Add(' and RE.DT_Entrada between :dt1 and :dt2');
      zqfinanciamento.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      zqfinanciamento.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      zqfinanciamento.SQL.Add(' and RE.DT_Vencimento between :dt3 and :dt4');
      zqfinanciamento.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      zqfinanciamento.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      zqfinanciamento.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      zqfinanciamento.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      zqfinanciamento.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' à '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      zqfinanciamento.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      zqfinanciamento.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      zqfinanciamento.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;  }
    zqfinanciamento.SQL.Add(' group by RE.quadralote order by lt.apelido,re.nomeadversa,re.quadralote');
//    zqfinanciamento.SQL.Add(' group by RE.quadralote,re.saldo  order by re.quadralote,re.saldo desc,RE.'+VArOrdem);
//    zqfinanciamento.SQL.Add(' group by RE.quadralote order by RE.nomeadversa');
//    Mensagem(zqfinanciamento.SQL.Text);
    zqfinanciamento.Open;
    zqfinanciamento.First;
    if zqfinanciamento.recordcount>0 then
    begin
      varquadra:=zqfinanciamento.FieldByName('quadralote').AsString;
      JSP1.Visible:=true;
      JSP1.Maximum:=zqfinanciamento.RecordCount;
      JSP1.Position:=0;

      CDS_finan.Close;
      CDS_finan.CreateDataSet;
      CDS_finan.Open;
      CDS_finan.insert;
      CDS_finanidrecebimento.Value:=zqfinanciamento.FieldByName('idrecebimento').AsLargeInt;
      CDS_finandocumento.Value:=zqfinanciamento.FieldByName('documento').AsString;
      CDS_financliente.Value:=zqfinanciamento.FieldByName('cliente').AsInteger;
      CDS_finanDt_Entrada.Value:=zqfinanciamento.FieldByName('Dt_Entrada').AsDateTime;
      CDS_finanDt_Vencimento.Value:=zqfinanciamento.FieldByName('Dt_Vencimento').AsDateTime;
      posi:=pos('/',zqfinanciamento.FieldByName('ordem').AsString);
      inc(posi);
      CDS_finanordem.Value:=copy(zqfinanciamento.FieldByName('ordem').AsString,posi,5);
      CDS_finansaldo.Value:=zqfinanciamento.FieldByName('saldo').AsFloat;
      CDS_finanidloteamento.Value:=zqfinanciamento.FieldByName('idloteamento').AsInteger;
      CDS_finanquadralote.Value:=zqfinanciamento.FieldByName('quadralote').AsString;
      CDS_finannomeadversa.Value:=zqfinanciamento.FieldByName('nomeadversa').AsString;
      CDS_finanapelido.Value:=zqfinanciamento.FieldByName('apelido').AsString;
      CDS_finan.post;
      zqfinanciamento.Next;
      zqfinanciamento.DisableControls;
      while not zqfinanciamento.Eof do
      begin
        JSP1.Position:=zqfinanciamento.RecNo;
        if varquadra<>zqfinanciamento.FieldByName('quadralote').AsString then
        begin
          CDS_finan.insert;
          CDS_finanidrecebimento.Value:=zqfinanciamento.FieldByName('idrecebimento').AsLargeInt;
          CDS_finandocumento.Value:=zqfinanciamento.FieldByName('documento').AsString;
          CDS_financliente.Value:=zqfinanciamento.FieldByName('cliente').AsInteger;
          CDS_finanDt_Entrada.Value:=zqfinanciamento.FieldByName('Dt_Entrada').AsDateTime;
          CDS_finanDt_Vencimento.Value:=zqfinanciamento.FieldByName('Dt_Vencimento').AsDateTime;
          posi:=pos('/',zqfinanciamento.FieldByName('ordem').AsString);
          inc(posi);
          CDS_finanordem.Value:=copy(zqfinanciamento.FieldByName('ordem').AsString,posi,5);
          CDS_finansaldo.Value:=zqfinanciamento.FieldByName('saldo').AsFloat;
          CDS_finanidloteamento.Value:=zqfinanciamento.FieldByName('idloteamento').AsInteger;
          CDS_finanquadralote.Value:=zqfinanciamento.FieldByName('quadralote').AsString;
          CDS_finannomeadversa.Value:=zqfinanciamento.FieldByName('nomeadversa').AsString;
          CDS_finanapelido.Value:=zqfinanciamento.FieldByName('apelido').AsString;
          CDS_finan.post;
          varquadra:=zqfinanciamento.FieldByName('quadralote').AsString;
        end;
        zqfinanciamento.Next;
      end;
      zqfinanciamento.EnableControls;
    end;
    JSP1.Position:=0;
    JSP1.Visible:=false;


    if frm_financiamento=nil then
      frm_financiamento:=Tfrm_financiamento.Create(Application);
//    if not empty(vrtexto) Then
//      frm_financiamento.RLLabel2.Caption := 'Período de : '+vrtexto;

//    frm_financiamento.RLGroup1.DataFields:='idloteamento;'+VArOrdem;


    frm_financiamento.RLReport1.Previewmodal;
    FreeAndNil(frm_financiamento);
    ZQEmpree.Close;
    zqfinanciamento.close;
    CDS_finan.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
  end;

  if RGFormato.ItemIndex = 7 Then
  Begin
    if FrmRelReceb04=nil then
      FrmRelReceb04:=TFrmRelReceb04.Create(Application);

    vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];

    if not empty(XDEVencimentoInicio.Text) Then
      vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' à '+XDEVencimentoFinal.Text;
    if not empty(XDEEntradaInicio.Text) Then
      vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' à '+XDEEntradaFinal.Text;
    if not empty(XDEBaixasInicio.Text) Then
      vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' à '+XDEBaixasFinal.Text;
    if not empty(XDERefInicio.Text) Then
      vartitulo := vartitulo + ', com Ref/Baixas de '+XDERefInicio.Text+' à '+XDERefFinal.Text;

    vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;

    ZQTempReceber.Close;
    ZQTempReceber.SQL.Clear;
    ZQTempReceber.SQL.Add('select re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
    ZQTempReceber.SQL.Add('       re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
    ZQTempReceber.SQL.Add('       re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
    ZQTempReceber.SQL.Add('       pa.idpaticipante,pa.nome_parte,pa.doc1,pa.doc2,pa.endereco,pa.bairro,pa.cidade,pa.cep,pa.ende_cob,pa.bairro_cob,pa.cidade_cob,pa.cep_cob,pa.Fone1,pa.fone2,');
    ZQTempReceber.SQL.Add('       pa.fone3,pa.tipopessoa,pa.aniversario,pa.nacionalidade,pa.email,pa.naturalidade,pa.localdetrab,pa.profissao,pa.estadocivil,pa.renda,pa.observacao,pa.codpaginc,');
    ZQTempReceber.SQL.Add('       pa.codrecinc,pa.codpagBx,pa.codrecBx,pa.fone4,pa.complemento,pa.complemento_cob,pa.cadastrado ');
    PrepararNomesTempReceber;
    if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
    begin
        ZQTempReceber.SQL.Add('   ,H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,H.sq,H.valor_parcela,');
        ZQTempReceber.SQL.Add('   BX.idreceb_baixa,BX.Dt_rec,BX.Vr_rec,BX.TipDoc,BX.Docum,BX.Juros_Vr,BX.Desc_Vr,BX.RefBaixa,BX.obsebx,BX.Codcontabil,BX.dataref,BX.vencimento,BX.substituicao,BX.vr_ab,BX.sq,BX.valor_parcela ');
    end;
    ZQTempReceber.SQL.Add('     from (recebimento as RE join participante as pa on pa.idpaticipante=re.cliente )');
    if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
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
//    if CBDesconsidera.Checked Then
  //     ZQTempReceber.SQL.Add(' and ( Numboleto is null or Numboleto='+quotedstr('')+') ');

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
      ZQTempReceber.SQL.Add(' and RE.TIPDOC in ('+varDoc+')');
    if CBSaldo.ItemIndex = 1 Then
       ZQTempReceber.SQL.Add(' and (RE.saldo <re.valor or RE.saldo is null)');
    if CBSaldo.ItemIndex = 2 Then
       ZQTempReceber.SQL.Add(' and RE.saldo > 0');

    ZQTempReceber.SQL.Add(' and RE.dt_nao_pagou_no_mes > 0');

    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQTempReceber.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZQTempReceber.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQTempReceber.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQTempReceber.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZQTempReceber.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQTempReceber.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
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
      Vrtexto := 'Baixado de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;
    // 24/11/2010
//    if CBDesconsidera.Checked Then
//       ZQTempReceber.SQL.Add(' and  not exists (select * from remessa_receb where idrec=idrecebimento)');
  //
    ZQTempReceber.SQL.Add(' group by re.idrecebimento');
    ZQTempReceber.SQL.Add(' ORDER BY '+VarOrdem);
    ZQTempReceber.Open;

    if not empty(vrtexto) Then
      FrmRelReceb04.RLLabel2.Caption := 'Período de : '+vrtexto+' - Ordem '+VArOrdem;

    FrmRelReceb04.RLLabel1.Caption := 'Contas Receber - Parcelas Fora do Mês de Vecto.';

    FrmRelReceb04.RLReport1.PreviewModal;
    FreeAndNil(FrmRelReceb04);
    ZQTempReceber.Close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
  end;

  if RGFormato.ItemIndex = 8 Then
  Begin
    if FrmRelReceb05=nil then
      FrmRelReceb05:=TFrmRelReceb05.Create(Application);

    vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];

    if not empty(XDEVencimentoInicio.Text) Then
      vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' à '+XDEVencimentoFinal.Text;
    if not empty(XDEEntradaInicio.Text) Then
      vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' à '+XDEEntradaFinal.Text;
    if not empty(XDEBaixasInicio.Text) Then
      vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' à '+XDEBaixasFinal.Text;
    if not empty(XDERefInicio.Text) Then
      vartitulo := vartitulo + ', com Ref/Baixas de '+XDERefInicio.Text+' à '+XDERefFinal.Text;
    vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;

    ZQTempReceber.Close;
    ZQTempReceber.SQL.Clear;
    ZQTempReceber.SQL.Add('select re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
    ZQTempReceber.SQL.Add('       re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
    ZQTempReceber.SQL.Add('       re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
    ZQTempReceber.SQL.Add('       pa.idpaticipante,pa.nome_parte,pa.doc1,pa.doc2,pa.endereco,pa.bairro,pa.cidade,pa.cep,pa.ende_cob,pa.bairro_cob,pa.cidade_cob,pa.cep_cob,pa.Fone1,pa.fone2,');
    ZQTempReceber.SQL.Add('       pa.fone3,pa.tipopessoa,pa.aniversario,pa.nacionalidade,pa.email,pa.naturalidade,pa.localdetrab,pa.profissao,pa.estadocivil,pa.renda,pa.observacao,pa.codpaginc,');
    ZQTempReceber.SQL.Add('       pa.codrecinc,pa.codpagBx,pa.codrecBx,pa.fone4,pa.complemento,pa.complemento_cob,pa.cadastrado ');
    PrepararNomesTempReceber;
    if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
    begin
        ZQTempReceber.SQL.Add('   ,H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,H.sq,H.valor_parcela,');
        ZQTempReceber.SQL.Add('   BX.idreceb_baixa,BX.Dt_rec,BX.Vr_rec,BX.TipDoc,BX.Docum,BX.Juros_Vr,BX.Desc_Vr,BX.RefBaixa,BX.obsebx,BX.Codcontabil,BX.dataref,BX.vencimento,BX.substituicao,BX.vr_ab,BX.sq,BX.valor_parcela ');
    end;
    ZQTempReceber.SQL.Add(' from (recebimento as RE join participante as pa on pa.idpaticipante=re.cliente )');
    if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
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
   // if CBDesconsidera.Checked Then
      // ZQTempReceber.SQL.Add(' and ( Numboleto is null or Numboleto='+quotedstr('')+') ');

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
      ZQTempReceber.SQL.Add(' and RE.TIPDOC in ('+varDoc+')');
    if CBSaldo.ItemIndex = 1 Then
       ZQTempReceber.SQL.Add(' and (RE.saldo <re.valor or RE.saldo is null)');
    if CBSaldo.ItemIndex = 2 Then
       ZQTempReceber.SQL.Add(' and RE.saldo > 0');

    ZQTempReceber.SQL.Add(' and RE.juridico=''SIM''');

    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQTempReceber.SQL.Add(' and Data_juridico between :dt1 and :dt2');
      ZQTempReceber.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQTempReceber.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQTempReceber.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZQTempReceber.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQTempReceber.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
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
      Vrtexto := 'Baixado de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;
    // 24/11/2010
   // if CBDesconsidera.Checked Then
     //  ZQTempReceber.SQL.Add(' and  not exists (select * from remessa_receb where idrec=idrecebimento)');
  //
    ZQTempReceber.SQL.Add(' group by re.idrecebimento');
    ZQTempReceber.SQL.Add(' ORDER BY '+VarOrdem);
    ZQTempReceber.Open;

    if not empty(vrtexto) Then
      FrmRelReceb05.RLLabel2.Caption := 'Período de : '+vrtexto+' - Ordem '+VArOrdem;

    FrmRelReceb05.RLLabel1.Caption := 'Contas Receber - Parcelas Em Jurídico';

    FrmRelReceb05.RLReport1.PreviewModal;
    FreeAndNil(FrmRelReceb05);
    ZQTempReceber.Close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;

  end;


  if RGFormato.ItemIndex = 6 Then
  Begin
    FrmRelReajuste2:=nil;
    if FrmRelReajuste2=nil then
       FrmRelReajuste2:=TFrmRelReajuste2.Create(Application);
    FrmRelReajuste2.RLLabel2.Caption := 'Contratos Reajustados no Período de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
//    FrmRelReajuste2.rlpercent.Caption:=transform(percent.Value,'##0.00');

    ZQRecebimento.Close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
    if CBAgrupado.Checked then
    begin
       ZQRecebimento.SQL.Add(' from Recebimento where (TipDoc=''BO'') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and (Data_reajuste between :dt1 and  :dt2) and (saldo>0) ');
       if not empty(VarEmpree) then
          ZQRecebimento.SQL.Add(' and (idloteamento in ('+VarEmpree+'))');
       ZQRecebimento.SQL.Add('group by quadralote order by '+VArOrdem);

    end
    else
    begin
       ZQRecebimento.SQL.Add(' from Recebimento where (TipDoc=''BO'') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and (Data_reajuste between :dt1 and  :dt2) and (saldo>0) ');
       if not empty(VarEmpree) then
          ZQRecebimento.SQL.Add(' and (idloteamento in ('+VarEmpree+'))');
       ZQRecebimento.SQL.Add(' order by '+VArOrdem);
    end;
    ZQRecebimento.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
    ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    ZQRecebimento.open;
    PrepararLookupsReajuste;
    ZQRecebimento.First;
    if RGFormato.ItemIndex=6 then
    begin
      case RGOrdem.ItemIndex of
        0:FrmRelReajuste2.RLGroup1.DataFields:='DT_Vencimento';
        1:FrmRelReajuste2.RLGroup1.DataFields:='nomeadversa';
        2:FrmRelReajuste2.RLGroup1.DataFields:='quadralote';
        3:FrmRelReajuste2.RLGroup1.DataFields:='TipDoc';
        4:FrmRelReajuste2.RLGroup1.DataFields:='idloteamento';

      end;
    end;


    FrmRelReajuste2.RLReport1.PreviewModal;
    FreeAndNil(FrmRelReajuste2);
    ZQRecebimento.Close;

    JSP1.Visible := False;
    exit;
  end;


  if RGFormato.ItemIndex = 5 Then
  Begin
    if cbtotais.Checked=true then
    begin
  //    ZQEmpree.Open;
      ZRTotPago.SQL.Clear;
      ZRTotPago.SQL.Add('SELECT RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.substituicao,RB.vr_ab,RB.sq,RB.valor_parcela,');
      ZRTotPago.SQL.Add('	   RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.valor_parcela,');
      ZRTotPago.SQL.Add('	   re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
      ZRTotPago.SQL.Add('    re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
      ZRTotPago.SQL.Add('    re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
      ZRTotPago.SQL.Add('	   re.sq,re.somar ');
      ZRTotPago.SQL.Add('    , sum(Re.valor) as somatit, sum(Rb.vr_rec) as somapago, sum(rb.juros_vr) as somajuros FROM receb_baixa as RB');
      ZRTotPago.SQL.Add(' join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
      ZRTotPago.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');
      if not empty(vardevedor) then
        ZRTotPago.SQL.Add(' and adversa in ('+vardevedor+')');

      // coloquei aqui no dia 15/06/2010
      if not empty(varquadra) then
        ZRTotPago.SQL.Add(' and quadralote in ('+varquadra+')');
      if RGContas.ItemIndex = 0 Then
        ZRTotPago.SQL.Add(' and re.recpag ='+quotedstr('R'))
      else
        ZRTotPago.SQL.Add(' and re.recpag ='+quotedstr('P'));
      if not empty(Varaleatorio) then
        ZRTotPago.SQL.Add(' and cliente in ('+Varaleatorio+')');
      if not empty(VarEmpree) then
        ZRTotPago.SQL.Add(' and idloteamento in ('+VarEmpree+')');
      if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
        ZRTotPago.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
        ZRTotPago.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
        ZRTotPago.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
        Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
      end;
      if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
        ZRTotPago.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
        ZRTotPago.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
        ZRTotPago.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
        if not empty(vrtexto) Then
          vrtexto := vrtexto+' e ';
        Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
      end;
      if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
        ZRTotPago.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
        ZRTotPago.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
        ZRTotPago.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
        if not empty(vrtexto) Then
          vrtexto := vrtexto+' e ';
        Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' à '+XDEBaixasFinal.DateText;
      end;
      if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
        ZRTotPago.SQL.Add(' and RB.dataref between :dt7 and :dt8');
        ZRTotPago.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
        ZRTotPago.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
        if not empty(vrtexto) Then
          vrtexto := vrtexto+' e ';
        Vrtexto := 'Referência de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
      end;
      ZRTotPago.SQL.Add(' group by Re.quadralote ');
      ZRTotPago.Open;
      try
        if FRM_Recebitotpago=nil then
           FRM_Recebitotpago:=TFRM_Recebitotpago.Create(Application);
        FRM_Recebitotpago.RLReport1.PreviewModal;
      finally
        ZRTotPago.Close;
//        ZQEmpree.close;
        freeAndNil(FRM_Recebitotpago);
      end;

      JSP1.Visible := False;
      exit;
    end;


   if cbmensal.Checked=true then
   begin
    ZQmensal2.SQL.Clear;
    ZQmensal2.SQL.Add('select month(rb.dataref) as Mes, case month(rb.dataref)');
    ZQmensal2.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Março''');
    ZQmensal2.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQmensal2.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQmensal2.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
//    ZQmensal2.SQL.Add(' end AS MESDESC, CASE when (re.valor <> rb.vr_rec) then re.valor else rb.Vr_rec end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
    ZQmensal2.SQL.Add(' end AS MESDESC, rh.valor as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, RH.valor_parcela,');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQmensal2.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQmensal2.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
    end;

    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZQmensal2.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZQmensal2.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' à '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZQmensal2.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZQmensal2.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZQmensal2.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;
    // Keep the detail selection local; relmensal is a table in some databases.
    ZQrelatorio.close;
    ZQrelatorio.SQL.clear;
    ZQrelatorio.SQL.Add('select M.*,sum(M.vr_rel) as valorpagojuros, sum(M.valor_parcela) as valorpago from (');
    ZQrelatorio.SQL.Add(ZQmensal2.SQL.Text);
    ZQrelatorio.SQL.Add(') as M group by M.mes,M.ano order by M.ano,M.mes');
    ZQrelatorio.Params.Assign(ZQmensal2.Params);
    ZQrelatorio.open;
    FrmRelReceb02_mensal:=nil;
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

    FrmRelReceb02_mensal.RLReport1.PreviewModal;
    FreeAndNil(FrmRelReceb02_mensal);
    ZQEmpree.Close;
    ZQmensal2.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;



   if cbdiario.Checked=true then
   begin
    ZRdiario.close;
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZRdiario.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZRdiario.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZRdiario.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZRdiario.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' à '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZRdiario.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZRdiario.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZRdiario.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;
    ZRdiario.SQL.Add(' order by '+VArOrdem);
    ZRdiario.Open;


    frm_rel_diario:=nil;
    if frm_rel_diario=nil then
       frm_rel_diario:=Tfrm_rel_diario.Create(Application);
    if not empty(vrtexto) Then
      frm_rel_diario.RLLabel2.Caption := 'Período de : '+vrtexto;

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

    frm_rel_diario.RLReport1.PreviewModal;
    FreeAndNil(frm_rel_diario);

    ZQEmpree.Close;
    ZRdiario.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;

   if cbcarteira.Checked=true then
   begin
    ZQcarteira.SQL.Clear;
    ZQcarteira.SQL.Add('select month(rb.dataref) as Mes, lt.apelido as ap, case month(rb.dataref)');
    ZQcarteira.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Março''');
    ZQcarteira.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQcarteira.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQcarteira.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
//    ZQcarteira.SQL.Add(' end AS MESDESC, CASE when (re.valor <> rb.vr_rec) then re.valor else rb.Vr_rec end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros, ');
    ZQcarteira.SQL.Add(' end AS MESDESC, rh.valor as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.DT_Entrada, re.DT_Vencimento, rb.dataref, rb.vr_rec,rb.juros_vr,RE.valor,RH.descon,RH.juros,RH.valor_parcela, ');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQcarteira.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQcarteira.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
    end;

    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZQcarteira.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZQcarteira.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' à '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZQcarteira.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZQcarteira.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZQcarteira.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;
    // Keep the detail selection local; relcarteira is a table in some databases.
    ZQRelCarteira.close;
    ZQRelCarteira.SQL.clear;
    ZQRelCarteira.SQL.Add('select C.*,sum(C.vr_rel) as valorpagojuros, sum(C.valor_parcela) as valorpago from (');
    ZQRelCarteira.SQL.Add(ZQcarteira.SQL.Text);
    ZQRelCarteira.SQL.Add(') as C group by C.idloteamento,C.mes,C.ano order by C.idloteamento,C.ano,C.mes');
    ZQRelCarteira.Params.Assign(ZQcarteira.Params);
    ZQRelCarteira.open;

    FrmRelReceb02_carteira:=nil;
    if FrmRelReceb02_carteira=nil then
       FrmRelReceb02_carteira:=TFrmRelReceb02_carteira.Create(Application);
    if not empty(vrtexto) Then
      FrmRelReceb02_carteira.RLLabel2.Caption := 'Período de : '+vrtexto;
    FrmRelReceb02_carteira.RLLabel1.Caption := 'Administração da Carteira de Recebíveis';
    if RGContas.ItemIndex = 1 Then

    FrmRelReceb02_carteira.RLLabel1.Caption := 'Administração da Carteira de Recebíveis';
    FrmRelReceb02_carteira.RLLabel17.Caption:='Vr. Saldo Baixado  ';
    FrmRelReceb02_carteira.RLDBResult2.Text:='Saldo Geral Baixado:    ';
    FrmRelReceb02_carteira.RLDBResult1.Text:='Saldo Total Baixado:    ';
    FrmRelReceb02_carteira.RLLabel11.Caption:='Vr. Saldo Baixado/Variação';
//    FrmRelReceb02_carteira.RLDBResult3.Text:='Saldo Geral Baixado/Variação:    ';
//    FrmRelReceb02_carteira.RLDBResult4.Text:='Saldo Total Baixado/Variação:    ';

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

    FrmRelReceb02_carteira.RLReport1.Previewmodal;
    FreeAndNil(FrmRelReceb02_carteira);
    ZQEmpree.Close;
    ZQcarteira.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;

    FRM_Recebi03:=nil;
    if FRM_Recebi03=nil then
       FRM_Recebi03:=TFRM_Recebi03.Create(Self);
    FRM_Recebi03.XNumEdit1.Value := 0;
    FRM_Recebi03.XNumEdit2.Value := 0;
    FRM_Recebi03.XNumEdit3.Value := 0;
    CDSTipoDoc.Close;
    CDSTipoDoc.CreateDataSet;
    ZROQReceb.SQL.Clear;
    Vrtexto:='';
//    ZROQReceb.SQL.Add('SELECT *, sum(Re.valor) as somatit, sum(vr_rec) as somapago, sum(juros_vr) as somajuros FROM receb_baixa as RB');
    ZROQReceb.SQL.Add('SELECT 	   RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.substituicao,RB.vr_ab,RB.sq,RB.valor_parcela,');
    ZROQReceb.SQL.Add('	   RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.valor_parcela,');
    ZROQReceb.SQL.Add('	   re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
    ZROQReceb.SQL.Add('    re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
    ZROQReceb.SQL.Add('    re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
    ZROQReceb.SQL.Add('	   re.sq,re.somar ');
    ZROQReceb.SQL.Add('    , sum(Re.valor) as somatit, sum(RH.valor) as somapago, sum(rb.juros_vr) as somajuros,sum(RB.Desc_Vr) as somadescontos FROM receb_baixa as RB');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZROQReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZROQReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZROQReceb.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZROQReceb.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' à '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.dataref between :dt7 and :dt8');
      ZROQReceb.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZROQReceb.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' à '+XDERefFinal.DateText;
    end;
//    if TemDuplicidade then
//      ZROQReceb.SQL.Add('GROUP BY RH.idrecib,RB.tipdoc')
//    else
      ZROQReceb.SQL.Add(' group by RB.tipdoc ');
    // coloquei aqui 24/01/2011   para nao sair duplicado as parcelas
 //   ZROQReceb.SQL.Add(' group by rb.idreceb_baixa ');
    ZROQReceb.Open;
    ZROQReceb.RecordCount;
    if not empty(vrtexto) Then
      FRM_Recebi03.RLLabel2.Caption := 'Período de : '+vrtexto;
    FRM_Recebi03.XNumEdit1.Value := 0;
    ZROQReceb.DisableControls;
    while not ZROQReceb.Eof do begin
      CDSTipoDoc.Insert;
      CDSTipoDocTipoDoc.Value      := ZROQReceb.FieldByName('TipDoc').AsString;
      CDSTipoDocsomatit.Value      := ZROQReceb.FieldByName('somatit').AsFloat;
      CDSTipoDocsomapago.Value     := ZROQReceb.FieldByName('somapago').AsFloat;
      CDSTipoDocsomajuros.Value    := ZROQReceb.FieldByName('somajuros').AsFloat;
      CDSTipoDocsomadescontos.Value:= ZROQReceb.FieldByName('somadescontos').AsFloat;
      CDSTipoDoc.Post;
      FRM_Recebi03.XNumEdit1.Value := FRM_Recebi03.XNumEdit1.Value+ZROQReceb.FieldByName('somatit').AsFloat;
//      FRM_Recebi03.              :=0;
      FRM_Recebi03.XNumEdit2.Value := FRM_Recebi03.XNumEdit2.Value+ZROQReceb.FieldByName('somapago').AsFloat;
      FRM_Recebi03.XNumEdit3.Value := FRM_Recebi03.XNumEdit3.Value+ZROQReceb.FieldByName('somajuros').AsFloat;
      FRM_Recebi03.XNumEdit8.Value := FRM_Recebi03.XNumEdit8.Value+ZROQReceb.FieldByName('somadescontos').AsFloat;
      ZROQReceb.Next;
    end;
    ZROQReceb.EnableControls;
    ZROQReceb1.SQL.Clear;
    ZROQReceb1.SQL.Add('SELECT RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.substituicao,RB.vr_ab,RB.sq,RB.valor_parcela,');
    ZROQReceb1.SQL.Add('	     RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.valor_parcela,');
    ZROQReceb1.SQL.Add('	     re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
    ZROQReceb1.SQL.Add('       re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
    ZROQReceb1.SQL.Add('       re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
    ZROQReceb1.SQL.Add('	     re.sq,re.somar,re.Reajustado,re.Data_reajuste,re.Proximo_Reajuste,re.Parcelas_fixas ');
    ZROQReceb1.SQL.Add('  FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
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
 //   ZROQReceb1.SQL.Add(' group by rb.idreceb_baixa ');

     if TemDuplicidade then
      ZROQReceb1.SQL.Add('GROUP BY RH.refer,RH.idrecib');

    if VArOrdem='TipDoc' then
       ZROQReceb1.SQL.Add('  ORDER BY RB.'+VarOrdem)
    else
       ZROQReceb1.SQL.Add('  ORDER BY '+VarOrdem);
    ZROQReceb1.Open;
    ZROQReceb1.recordcount;
   //resumo

    ZqresumoE.SQL.Clear;
    ZqresumoE.SQL.Add('SELECT  RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.substituicao,RB.vr_ab,RB.sq,RB.valor_parcela,');
    ZqresumoE.SQL.Add('   	   RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.valor_parcela,');
    ZqresumoE.SQL.Add('	       re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
    ZqresumoE.SQL.Add('        re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
    ZqresumoE.SQL.Add('        re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
    ZqresumoE.SQL.Add('	       re.sq,re.somar,re.Reajustado,re.Data_reajuste,re.Proximo_Reajuste,re.Parcelas_fixas ');
    ZqresumoE.SQL.Add('  FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
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
//    ZROQReceb1.SQL.Add(' group by rb.idreceb_baixa  ');
    if TemDuplicidade then
      ZqresumoE.SQL.Add('GROUP BY RH.idrecib');

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
    ZqresumoE.DisableControls;
    while not ZqresumoE.eof do begin
      JSP1.Position:=ZqresumoE.RecNo;
      if mesano(ZqresumoE.FieldByName('dataref').AsDateTime)=mesano(ZqresumoE.FieldByName('Dt_Vencimento').AsDateTime) then //mes(ZqresumoEvencimento.Value) then
      begin
{        if ZqresumoEvr_rec.Value>ZqresumoEvalor_1.Value then
        begin
           xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoEvalor_1.Value+ZqresumoEJuros_Vr.Value;
        end
        else
        begin
           if ZqresumoEVr_rec.Value>0 then
           begin}
             xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoE.FieldByName('valor').AsFloat;//+ZqresumoEJuros_Vr.Value;
{           end
           else
           begin
             xnentradas_dentro.Value:=xnentradas_dentro.Value+ZqresumoEvr_ab.Value;
           end;
        end;}
      end
      else
      begin
        mes1:=strtoint(mes(ZqresumoE.FieldByName('dataref').AsDateTime));
        mes2:=strtoint(mes(ZqresumoE.FieldByName('Dt_Vencimento').AsDateTime)); //strtoint(mes(ZqresumoEvencimento.Value));
        ano1:=strtoint(ano(ZqresumoE.FieldByName('dataref').AsDateTime));
        ano2:=strtoint(ano(ZqresumoE.FieldByName('Dt_Vencimento').AsDateTime)); //strtoint(mes(ZqresumoEvencimento.Value));
        if ((mes1>mes2) or (ano1>ano2)) or ((mes1>mes2) and (ano1=ano2))then
        begin
          xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('valor').AsFloat;//+ZqresumoEJuros_Vr.Value;
          entrada_atrazada.Value:=entrada_atrazada.Value+ZqresumoE.FieldByName('valor').AsFloat;//+ZqresumoEJuros_Vr.Value;
        end
        else
        begin
          xnentradas_fora.Value:=xnentradas_fora.Value+ZqresumoE.FieldByName('valor').AsFloat;//+ZqresumoEJuros_Vr.Value;
          entrada_adiantada.Value:=entrada_adiantada.Value+ZqresumoE.FieldByName('valor').AsFloat;//+ZqresumoEJuros_Vr.Value;
        end;
      end;
      if ZqresumoE.Active then
         ZqresumoE.Next;
    end;
    ZqresumoE.EnableControls;

    ZqresumoP.SQL.Clear;
    ZqresumoP.SQL.Add('SELECT RB.idreceb_baixa,RB.Dt_rec,RB.Vr_rec,RB.TipDoc,RB.Docum,RB.Juros_Vr,RB.Desc_Vr,RB.RefBaixa,RB.obsebx,RB.Codcontabil,RB.dataref,RB.vencimento,RB.substituicao,RB.vr_ab,RB.sq,RB.valor_parcela,');
    ZqresumoP.SQL.Add('	      RH.idrecbxhist,RH.refer,RH.idrecib,RH.valor,RH.descon,RH.juros,RH.percent_usado,RH.data,RH.sq,RH.valor_parcela,');
    ZqresumoP.SQL.Add('	      re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
    ZqresumoP.SQL.Add('       re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
    ZqresumoP.SQL.Add('       re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
    ZqresumoP.SQL.Add('	      re.sq,re.somar,re.Reajustado,re.Data_reajuste,re.Proximo_Reajuste,re.Parcelas_fixas ');
    ZqresumoP.SQL.Add('  FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento ');
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
//    ZROQReceb1.SQL.Add(' group by rb.idreceb_baixa  ');

    if TemDuplicidade then
       ZqresumoP.SQL.Add('GROUP BY RH.idrecib');

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
    ZqresumoP.DisableControls;
    while not ZqresumoP.eof do begin
      JSP1.Position:=ZqresumoP.RecNo;
      if mesano(ZqresumoP.FieldByName('dataref').AsDateTime)=mesano(ZqresumoP.FieldByName('Dt_Vencimento').AsDateTime) then // mes(ZqresumoPvencimento.Value) then
      begin
       { if ZqresumoPvr_rec.Value>ZqresumoPvalor_1.Value then
        begin
           xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoPvalor_1.Value+ZqresumoPJuros_Vr.Value;
        end
        else
        begin
           if ZqresumoPVr_rec.Value>0 then
           begin }
             xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoP.FieldByName('valor').AsFloat;//+ZqresumoPJuros_Vr.Value;
{           end
           else
           begin
             xnparcelas_dentro.Value:=xnparcelas_dentro.Value+ZqresumoPvr_ab.Value;
           end;
        end;}
      end
      else
      begin
        mes1:=strtoint(mes(ZqresumoP.FieldByName('dataref').AsDateTime));
        mes2:=strtoint(mes(ZqresumoP.FieldByName('Dt_Vencimento').AsDateTime)); //strtoint(mes(ZqresumoEvencimento.Value));
        ano1:=strtoint(ano(ZqresumoP.FieldByName('dataref').AsDateTime));
        ano2:=strtoint(ano(ZqresumoP.FieldByName('Dt_Vencimento').AsDateTime)); //strtoint(mes(ZqresumoEvencimento.Value));
        if ((mes1>mes2) or (ano1>ano2)) or ((mes1>mes2) and (ano1=ano2))then
        begin
          xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('valor').AsFloat;//+ZqresumoPJuros_Vr.Value;
          parcela_atrazada.Value:=parcela_atrazada.Value+ZqresumoP.FieldByName('valor').AsFloat;//+ZqresumoPJuros_Vr.Value;
        end
        else
        begin
          xnparcelas_fora.Value:=xnparcelas_fora.Value+ZqresumoP.FieldByName('valor').AsFloat;//+ZqresumoPJuros_Vr.Value;
          parcela_adiantada.Value:=parcela_adiantada.Value+ZqresumoP.FieldByName('valor').AsFloat;//+ZqresumoPJuros_Vr.Value;
        end;
      end;
      if ZqresumoP.Active then
         ZqresumoP.Next;
    end;
    ZqresumoP.EnableControls;
    JSP1.Position:=0;
    JSP1.Visible := false;
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
    FRM_Recebi03.RLReport1.Previewmodal;
    FreeAndNil(FRM_Recebi03);
    ZQEmpree.Close;
    ZqresumoP.close;
    ZqresumoE.close;

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

    CDS_MarcaTit.Close;
    CDS_MarcaTit.CreateDataSet;
    CDS_MarcaTit.open;
    FrmCobrancaBancaria:=nil;
    if FrmCobrancaBancaria=nil then
       FrmCobrancaBancaria:=TFrmCobrancaBancaria.Create(Self);
    FrmCobrancaBancaria.CDS_MarcaTit.Close;
    FrmCobrancaBancaria.CDS_MarcaTit.CreateDataSet;
    FrmCobrancaBancaria.CDS_MarcaTit.open;

    ZQincorp_loteame.DisableControls;
    While not ZQincorp_loteame.Eof do
    Begin
      JSP1.position:=ZQincorp_loteame.recno;
      application.ProcessMessages;
      if (pos(ZQincorp_loteame.FieldByName('incorporador_idincorporador').Text,Varaleatorio)>0) or empty(Varaleatorio) Then
      Begin
          CDSEmpreendimento.Insert;
          CDSEmpreendimentoidordem.Value          := CDSEmpreendimento.RecordCount;
          CDSEmpreendimentoidloteamento.Value     := ZQincorp_loteame.FieldByName('idloteamento').AsLargeInt;
          CDSEmpreendimentocodcontabancaria.Value := ZQincorp_loteame.FieldByName('codcontabancaria').AsInteger;
          CDSEmpreendimentoidparticipante.Value   := ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt;
          CDSEmpreendimentodigito_dif.Value       := ZQincorp_loteame.FieldByName('digito_dif').AsString;
          CDSEmpreendimentomarca.Value            := '0';
          CDSEmpreendimento.Post;
      end;
      ZQincorp_loteame.Next;
    end;
    ZQincorp_loteame.EnableControls;
    FrmCobrancaBancaria.CDSEmpreendimento.Close;
    FrmCobrancaBancaria.CDSEmpreendimento.CreateDataSet;
    FrmCobrancaBancaria.CDSEmpreendimento.open;
    //clona os a tebela toda
    FrmCobrancaBancaria.CDSEmpreendimento.CloneCursor(CDSEmpreendimento, true, true);
    FrmCobrancaBancaria.CDSEmpreendimento.Data:=CDSEmpreendimento.Data;

    JSP1.position:=0;
    JSP1.visible:=false;
    application.ProcessMessages;
    DM_Tabelas.ZQRecebParticp.SQL.Clear;
    DM_Tabelas.ZQRecebParticp.SQL.Add('select * from recebimento as r join participante as p  on r.adversa=p.idpaticipante join cidade as c on c.idcidade=p.cidade_cob');
    DM_Tabelas.ZQRecebParticp.SQL.Add(' where r.recpag = '+quotedstr('R'));
    if CBEntrada.Checked <> CBParcela.Checked Then Begin
      if CBParcela.Checked Then
        DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.documento like '+quotedstr('%-P-%'));
    end;
    if (not empty(vardevedor)) and (RGContas.ItemIndex=0) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.adversa in ('+vardevedor+')');


    // coloquei aqui no dia 15/06/2010
    if not empty(varquadra) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.quadralote in ('+varquadra+')');


    if not empty(VarEmpree) then
       DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.idloteamento in ('+VarEmpree+')');


    if not empty(Varaleatorio) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.cliente in ('+Varaleatorio+')');
    if CBSaldo.ItemIndex = 1 Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and (r.saldo < r.valor or r.saldo is null)');
    if CBSaldo.ItemIndex = 2 Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.saldo > 0');
    if not empty(vardocumento) Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(vardocumento);
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.DT_Entrada between :dt1 and :dt2');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.DT_Vencimento between :dt3 and :dt4');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    DM_Tabelas.ZQRecebParticp.SQL.Add(' and r.TipDoc in ('+varDoc+')');

    if CBDesconsidera.Checked Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and not exists (select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec=idrecebimento)');

    DM_Tabelas.ZQRecebParticp.Open;
//    Mensagem(DM_Tabelas.ZQRecebParticp.SQL.Text);
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
      CDS_MarcaTit.open;
    end;

    JSP1.visible:=true;
    JSP1.maximum:=DM_Tabelas.ZQRecebParticp.recordcount;
    JSP1.position:=0;

    DM_Tabelas.ZQRecebParticp.DisableControls;
    while not DM_Tabelas.ZQRecebParticp.Eof do
    begin
      JSP1.Position := DM_Tabelas.ZQRecebParticp.RecNo;
      application.ProcessMessages;
      FrmCobrancaBancaria.CDS_MarcaTit.Insert;
      FrmCobrancaBancaria.CDS_MarcaTitIdreceb.Value       := DM_Tabelas.ZQRecebParticp.FieldByName('idrecebimento').AsLargeInt;
      FrmCobrancaBancaria.CDS_MarcaTitidparti.Value       := DM_Tabelas.ZQRecebParticp.FieldByName('adversa').AsInteger;
      FrmCobrancaBancaria.CDS_MarcaTitcliente.Value       := DM_Tabelas.ZQRecebParticp.FieldByName('cliente').AsInteger;
      FrmCobrancaBancaria.CDS_MarcaTitidloteam.Value      := DM_Tabelas.ZQRecebParticp.FieldByName('idloteamento').AsInteger;
      FrmCobrancaBancaria.CDS_MarcaTitquadralote.Value    := DM_Tabelas.ZQRecebParticp.FieldByName('quadralote').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitmarca.Value         := '0';
      FrmCobrancaBancaria.CDS_MarcaTitnumboleto.Value     := DM_Tabelas.ZQRecebParticp.FieldByName('numboleto').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitDT_Entrada.value    := DM_Tabelas.ZQRecebParticp.FieldByName('Dt_Entrada').AsDateTime;
      FrmCobrancaBancaria.CDS_MarcaTitDt_Vencimento.value := DM_Tabelas.ZQRecebParticp.FieldByName('Dt_Vencimento').AsDateTime;
      FrmCobrancaBancaria.CDS_MarcaTitValor.value         := DM_Tabelas.ZQRecebParticp.FieldByName('Valor').AsFloat;
      FrmCobrancaBancaria.CDS_MarcaTitordem.value         := DM_Tabelas.ZQRecebParticp.FieldByName('ordem').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitnome_parte.value    := DM_Tabelas.ZQRecebParticp.FieldByName('nome_parte').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitende_cob.value      := DM_Tabelas.ZQRecebParticp.FieldByName('ende_cob').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitcep_cob.value       := DM_Tabelas.ZQRecebParticp.FieldByName('cep_cob').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitdoc1.value          := DM_Tabelas.ZQRecebParticp.FieldByName('doc1').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitTipDoc.value        := DM_Tabelas.ZQRecebParticp.FieldByName('TipDoc').AsString;
      FrmCobrancaBancaria.CDS_MarcaTitnumordem.value      := DM_Tabelas.ZQRecebParticp.FieldByName('numordem').AsInteger;
      FrmCobrancaBancaria.CDS_MarcaTitbairro.value        := DM_Tabelas.ZQRecebParticp.FieldByName('bairro_cob').AsString;
      if DM_Tabelas.ZQRecebParticp.FieldByName('cidade_cob').AsLargeInt>0 then
         FrmCobrancaBancaria.CDS_MarcaTitcidade.value     := DM_Tabelas.ZQRecebParticp.FieldByName('cidade_cob').AsLargeInt
      else
        FrmCobrancaBancaria.CDS_MarcaTitcidade.value      := DM_Tabelas.ZQRecebParticp.FieldByName('cidade').AsInteger;
      FrmCobrancaBancaria.CDS_MarcaTitnomecidade.Value    := DM_Tabelas.ZQRecebParticp.FieldByName('nomecid').AsString;
      FrmCobrancaBancaria.CDS_MarcaTituf.Value            := DM_Tabelas.ZQRecebParticp.FieldByName('estado').AsString;

      posi:=pos('-',DM_Tabelas.ZQRecebParticp.FieldByName('ordem').AsString);
      inc(posi);
      FrmCobrancaBancaria.CDS_MarcaTitordem.Value:= trim(copy(DM_Tabelas.ZQRecebParticp.FieldByName('ordem').AsString,posi,15));
      FrmCobrancaBancaria.CDS_MarcaTit.Post;
      DM_Tabelas.ZQRecebParticp.Next;
    end;
    DM_Tabelas.ZQRecebParticp.EnableControls;
    if CDS_MarcaTit.Active=false then
    begin
      CDS_MarcaTit.Close;
      CDS_MarcaTit.CreateDataSet;
      CDS_MarcaTit.open;
    end;

    //clona os a tebela toda
    CDS_MarcaTit.CloneCursor(FrmCobrancaBancaria.CDS_MarcaTit, true, true);
    CDS_MarcaTit.Data:=FrmCobrancaBancaria.CDS_MarcaTit.Data;

    JSP1.position:=0;
    JSP1.visible:=false;
    application.ProcessMessages;

//    FrmCobrancaBancaria.CDS_MarcaTit.recordcount;
//    CDS_MarcaTit.recordcount;
    dxbSelecionar.enabled:=false;
    CDS_MarcaTit.IndexName:='nome_venc';
//    FrmCobrancaBancaria.Top := FrmRelRecebimento.Top + 40;
  //  FrmCobrancaBancaria.Left :=FrmRelRecebimento.Left+30;
    AbrirModal(Self, FrmCobrancaBancaria);
    CDS_MarcaTit.IndexName:='';
    CDS_MarcaTit.Close;
    CDS_MarcaTit.CreateDataSet;
    CDS_MarcaTit.open;

//    dxbSelecionar.Enabled:=true;
//    dxbselecionar.SetFocus;
    Exit;
  end;
  vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
  if not empty(XDEVencimentoInicio.Text) Then
    vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' à '+XDEVencimentoFinal.Text;
  if not empty(XDEEntradaInicio.Text) Then
    vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' à '+XDEEntradaFinal.Text;
  if not empty(XDEBaixasInicio.Text) Then
    vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' à '+XDEBaixasFinal.Text;
    if not empty(XDERefInicio.Text) Then
      vartitulo := vartitulo + ', com Ref/Baixas de '+XDERefInicio.Text+' à '+XDERefFinal.Text;

  vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;

  ZQTempReceber.Close;
  ZQTempReceber.SQL.Clear;
  ZQTempReceber.Close;
  ZQTempReceber.SQL.Clear;
  ZQTempReceber.SQL.Add('select re.idrecebimento,re.documento,re.cliente,re.usuario,re.Dt_Entrada,re.Dt_Vencimento,re.Valor,re.Observ,re.VrDoc,re.ordem,re.TipDoc,re.saldo,re.marcar,');
  ZQTempReceber.SQL.Add('       re.RefBaixa,re.refvinda,re.contabil,re.empresa,re.custodaparcela,re.origem,re.adversa,re.recpag,re.numordem,re.idloteamento,');
  ZQTempReceber.SQL.Add('       re.venda_idvenda,re.quadralote,re.numboleto,re.Substituicao,re.nomeadversa,re.data_juridico,re.dt_nao_pagou_no_mes,re.descricao_juridico,re.juridico,');
  ZQTempReceber.SQL.Add('       pa.idpaticipante,pa.nome_parte,pa.doc1,pa.doc2,pa.endereco,pa.bairro,pa.cidade,pa.cep,pa.ende_cob,pa.bairro_cob,pa.cidade_cob,pa.cep_cob,pa.Fone1,pa.fone2,');
  ZQTempReceber.SQL.Add('       pa.fone3,pa.tipopessoa,pa.aniversario,pa.nacionalidade,pa.email,pa.naturalidade,pa.localdetrab,pa.profissao,pa.estadocivil,pa.renda,pa.observacao,pa.codpaginc,');
  ZQTempReceber.SQL.Add('       pa.codrecinc,pa.codpagBx,pa.codrecBx,pa.fone4,pa.complemento,pa.complemento_cob,pa.cadastrado ');
  PrepararNomesTempReceber;
  if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
  begin
      ZQTempReceber.SQL.Add('   ,H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,H.sq,H.valor_parcela,');
      ZQTempReceber.SQL.Add('   BX.idreceb_baixa,BX.Dt_rec,BX.Vr_rec,BX.TipDoc,BX.Docum,BX.Juros_Vr,BX.Desc_Vr,BX.RefBaixa,BX.obsebx,BX.Codcontabil,BX.dataref,BX.vencimento,BX.substituicao,BX.vr_ab,BX.sq,BX.valor_parcela ');
  end;
  ZQTempReceber.SQL.Add(' from (recebimento as RE join participante as pa on pa.idpaticipante=re.cliente )');
  if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
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
    ZQTempReceber.SQL.Add(' and RE.TIPDOC in ('+varDoc+')');
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
     ZQTempReceber.SQL.Add(' and  not exists (select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec=idrecebimento)');
  //
  ZQTempReceber.SQL.Add(' group by re.idloteamento,re.idrecebimento');
  ZQTempReceber.SQL.Add(' ORDER BY idloteamento,'+VarOrdem );
  ZQTempReceber.Open;
  //2
  if RGFormato.ItemIndex = 2 Then
  Begin

   Vrtexto := '';

   if cbmensal.Checked=true then
   begin
    ZQmensal.Close;
    ZQmensal.SQL.Clear;
    ZQmensal.SQL.Add(' select month(re.dt_vencimento) as Mes, case month(re.dt_vencimento)');
    ZQmensal.SQL.Add( 'when 1 then ''Janeiro'' when 2 then ''Fevereiro'' when 3 then ''Março''');
    ZQmensal.SQL.Add(' when 4 then ''Abril'' when 5 then ''Maio'' when 6 then ''Junho''');
    ZQmensal.SQL.Add(' when 7 then ''Julho'' when 8 then ''Agosto'' when 9 then ''Setembro''');
    ZQmensal.SQL.Add(' when 10 then ''Outubro'' when 11 then ''Novembro'' when 12 then ''Dezembro''');
    ZQmensal.SQL.Add(' end AS MESDESC, CASE when  re.valor then re.valor end as vr_rel, re.idloteamento, re.recpag, re.tipdoc, re.quadralote, re.valor, re.DT_Entrada, re.DT_Vencimento, ');
    ZQmensal.SQL.Add(' year(re.dt_vencimento) as Ano from recebimento as RE where re.saldo>0');
    if not empty(varDoc) then
      ZQmensal.SQL.Add(' and RE.TIPDOC in ('+varDoc+')');
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQmensal.SQL.Add(' and re.DT_Vencimento between :dt3 and :dt4');
      ZQmensal.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQmensal.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
    end;

    { A consulta derivada dispensa criar ou substituir objetos no banco. }
    ZQrelatorio2.close;
    ZQrelatorio2.SQL.clear;
    ZQrelatorio2.SQL.Add('select M.*,sum(M.vr_rel) as valorpago, sum(M.valor) as valorpagojuros from (');
    ZQrelatorio2.SQL.Add(ZQmensal.SQL.Text);
    ZQrelatorio2.SQL.Add(') as M group by M.mes,M.ano order by M.ano,M.mes');
    ZQrelatorio2.Params.Assign(ZQmensal.Params);
    ZQrelatorio2.open;

    FrmRelReceb02_mensal_ab:=nil;
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
    FrmRelReceb02_mensal_ab.RLReport1.PreviewModal;
    FreeAndNil(FrmRelReceb02_mensal_ab);
    ZQEmpree.Close;
    ZQmensal.close;
    ZQmensal2.close;
    vrtexto:='';
    JSP1.Visible := False;
    exit;
   end;

   if (stotais.Checked=false) and (CBQuadra.Checked=false) then
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
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZqresumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
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
    FrmRelReceb02:=nil;
    if FrmRelReceb02=nil then
       FrmRelReceb02:=TFrmRelReceb02.Create(Self);
    FrmRelReceb02.XNsomabaixa.Value:=0;
    FrmRelReceb02.XNNegociado.Value:=0;
    FrmRelReceb02.XNRenova.Value:=0;
    FrmRelReceb02.XNPrincipal.Value:=0;
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
    FrmRelReceb02.RLReport1.PreviewModal;
    FreeAndNil(FrmRelReceb02);
    ZQEmpree.Close;
    vrtexto:='';
   end
   else  if (stotais.Checked=true) and (CBQuadra.Checked=false) then
   begin
     vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
     if not empty(XDEVencimentoInicio.Text) Then
        vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' à '+XDEVencimentoFinal.Text;
     if not empty(XDEEntradaInicio.Text) Then
        vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' à '+XDEEntradaFinal.Text;
     if not empty(XDEBaixasInicio.Text) Then
        vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' à '+XDEBaixasFinal.Text;
    if not empty(XDERefInicio.Text) Then
      vartitulo := vartitulo + ', com Ref/Baixas de '+XDERefInicio.Text+' à '+XDERefFinal.Text;
     vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;
     ZQTempReceber2.Close;
     ZQTempReceber2.SQL.Clear;
     ZQTempReceber2.SQL.Add('select RE.idrecebimento,RE.documento,RE.cliente,RE.usuario,RE.Dt_Entrada,RE.Dt_Vencimento,RE.Valor,RE.Observ,RE.VrDoc,RE.ordem,RE.TipDoc,RE.saldo,RE.marcar,');
     ZQTempReceber2.SQL.Add('       RE.RefBaixa,RE.refvinda,RE.contabil,RE.empresa,RE.custodaparcela,RE.origem,RE.adversa,RE.recpag,RE.numordem,RE.idloteamento,');
     ZQTempReceber2.SQL.Add('       RE.venda_idvenda,RE.quadralote,RE.numboleto,RE.Substituicao,RE.nomeadversa,RE.data_juridico,RE.dt_nao_pagou_no_mes,RE.descricao_juridico,RE.juridico,');
     ZQTempReceber2.SQL.Add('       RE.sq,RE.somar,RE.Reajustado,RE.Data_reajuste,RE.Proximo_Reajuste,RE.Parcelas_fixas ');
     { Retorne todos os campos fisicos registrados no dataset compartilhado.
       As subconsultas nao acrescentam linhas nem alteram os totais. }
     ZQTempReceber2.SQL.Add(' ,(select L.apelido from loteamento L');
     ZQTempReceber2.SQL.Add('   where L.idloteamento = RE.idloteamento) as apelido_loteamento');
     ZQTempReceber2.SQL.Add(' ,(select COALESCE(NULLIF(L.nomeloteamento, ''''), L.apelido) from loteamento L');
     ZQTempReceber2.SQL.Add('   where L.idloteamento = RE.idloteamento) as nome_loteamento');
     ZQTempReceber2.SQL.Add(' ,COALESCE((select NULLIF(P.nome_parte, '''') from participante P');
     ZQTempReceber2.SQL.Add('   where P.idpaticipante = RE.adversa), NULLIF(RE.nomeadversa, ''''), '''') as comprador_relatorio');
    if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
     begin
        ZQTempReceber2.SQL.Add('   ,H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,H.sq,H.valor_parcela,');
        ZQTempReceber2.SQL.Add('   BX.idreceb_baixa,BX.Dt_rec,BX.Vr_rec,BX.TipDoc,BX.Docum,BX.Juros_Vr,BX.Desc_Vr,BX.RefBaixa,BX.obsebx,BX.Codcontabil,BX.dataref,BX.vencimento,BX.substituicao,BX.vr_ab,BX.sq,BX.valor_parcela ');
     end;

     ZQTempReceber2.SQL.Add(',sum(RE.valor) as tvr, sum(RE.saldo) as tsld from recebimento as RE ');
    if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
         ZQTempReceber2.SQL.Add('join recbxhist as H ON RE.idrecebimento=H.idrecib join receb_baixa as BX ON BX.refbaixa=H.refer');
     ZQTempReceber2.SQL.Add(' where 1 = 1 ');
     if not empty(vardocumento) Then
        ZQTempReceber2.SQL.Add(vardocumento);
     if RGFormato.ItemIndex <> 3 Then
     Begin
        if RGContas.ItemIndex = 0 Then
           ZQTempReceber2.SQL.Add(' and RE.recpag ='+quotedstr('R'))
        else
           ZQTempReceber2.SQL.Add(' and RE.recpag ='+quotedstr('P'));
     end;
    // 13/12/2010
     if CBDesconsidera.Checked Then
       ZQTempReceber2.SQL.Add(' and ( RE.Numboleto is null or RE.Numboleto='+quotedstr('')+') ');

     if not empty(VarEmpree) then
        ZQTempReceber2.SQL.Add(' and RE.idloteamento in ('+VarEmpree+')');
       // coloquei aqui no dia 15/06/2010
     if not empty(varquadra) then
        ZQTempReceber2.SQL.Add(' and RE.quadralote in ('+varquadra+')');
     if not empty(varDoc) Then
        ZQTempReceber2.SQL.Add(' and RE.TIPDOC in ('+varDoc+')');
     if CBSaldo.ItemIndex = 1 Then
        ZQTempReceber2.SQL.Add(' and (RE.saldo < RE.valor or RE.saldo is null)');
     if CBSaldo.ItemIndex = 2 Then
        ZQTempReceber2.SQL.Add(' and RE.saldo > 0');
     if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then
     Begin
       ZQTempReceber2.SQL.Add(' and RE.DT_Entrada between :dt1 and :dt2');
       ZQTempReceber2.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
       ZQTempReceber2.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
     end;
     if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
     Begin
       ZQTempReceber2.SQL.Add(' and RE.DT_Vencimento between :dt3 and :dt4');
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
        ZQTempReceber2.SQL.Add(' and not exists (select idremessa_receb,remessa,idrec,nossonumero from remessa_receb where idrec=RE.idrecebimento)');
     //

     ZQTempReceber2.SQL.Add(' group by RE.idloteamento');

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
       Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
     end;
     if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
     Begin
{      ZqresumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZqresumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZqresumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);}
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
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
     FrmRelReceb02_total:=nil;
     if FrmRelReceb02_total=nil then
        FrmRelReceb02_total:=TFrmRelReceb02_total.Create(Self);
     FrmRelReceb02_total.RLLabel1.Caption := 'Contas a Receber';
     if RGContas.ItemIndex = 1 Then
        FrmRelReceb02_total.RLLabel1.Caption := 'Contas a Pagar';
     if FrmRelReceb02_total=nil then
        FrmRelReceb02_total:=TFrmRelReceb02_total.Create(Application);
     if not empty(vrtexto) Then
        FrmRelReceb02_total.RLLabel2.Caption := 'Período de : '+vrtexto;
     FrmRelReceb02_total.RLReport1.PreviewModal;
     FreeAndNil(FrmRelReceb02_total);
     ZQEmpree.Close;
     vrtexto:='';

   end;
   if CBQuadra.Checked then
   begin
      vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];

      if not empty(XDEVencimentoInicio.Text) Then
        vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' à '+XDEVencimentoFinal.Text;

      if not empty(XDEEntradaInicio.Text) Then
        vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' à '+XDEEntradaFinal.Text;

      if not empty(XDEBaixasInicio.Text) Then
        vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' à '+XDEBaixasFinal.Text;

      if not empty(XDERefInicio.Text) Then
        vartitulo := vartitulo + ', com Ref/Baixas de '+XDERefInicio.Text+' à '+XDERefFinal.Text;

      vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;

      ZQTempReceber2.Close;
      ZQTempReceber2.SQL.Clear;

      ZQTempReceber2.SQL.Add('select ');
      ZQTempReceber2.SQL.Add(' R.idrecebimento,R.documento,R.cliente,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,');
      ZQTempReceber2.SQL.Add(' R.Valor,R.Observ,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,');
      ZQTempReceber2.SQL.Add(' R.RefBaixa,R.refvinda,R.contabil,R.empresa,R.custodaparcela,R.origem,');
      ZQTempReceber2.SQL.Add(' R.adversa,R.recpag,R.numordem,R.idloteamento,');
      ZQTempReceber2.SQL.Add(' R.venda_idvenda,R.quadralote,R.numboleto,R.Substituicao,');
      ZQTempReceber2.SQL.Add(' R.nomeadversa,R.data_juridico,R.dt_nao_pagou_no_mes,');
      ZQTempReceber2.SQL.Add(' R.descricao_juridico,R.juridico,R.sq,R.somar,');
      ZQTempReceber2.SQL.Add(' R.Reajustado,R.Data_reajuste,R.Proximo_Reajuste,R.Parcelas_fixas,');
      ZQTempReceber2.SQL.Add(' L.apelido as apelido_loteamento,');
      ZQTempReceber2.SQL.Add(' COALESCE(NULLIF(L.nomeloteamento, ''''), L.apelido) as nome_loteamento,');
      ZQTempReceber2.SQL.Add(' COALESCE(MAX(NULLIF(P.nome_parte, '''')),');
      ZQTempReceber2.SQL.Add(' MAX(NULLIF(R.nomeadversa, '''')), '''') as comprador_relatorio ');

        if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
      begin
        ZQTempReceber2.SQL.Add(' ,H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,');
        ZQTempReceber2.SQL.Add('  H.percent_usado,H.data,H.sq as sq_hist,H.valor_parcela as valor_parcela_hist,');
        ZQTempReceber2.SQL.Add('  BX.idreceb_baixa,BX.Dt_rec,BX.Vr_rec,BX.TipDoc as TipDoc_bx,');
        ZQTempReceber2.SQL.Add('  BX.Docum,BX.Juros_Vr,BX.Desc_Vr,BX.RefBaixa as RefBaixa_bx,');
        ZQTempReceber2.SQL.Add('  BX.obsebx,BX.Codcontabil,BX.dataref,BX.vencimento,');
        ZQTempReceber2.SQL.Add('  BX.substituicao,BX.vr_ab,BX.sq as sq_bx,BX.valor_parcela as valor_parcela_bx ');
      end;

      ZQTempReceber2.SQL.Add(' ,sum(R.valor) as tvr, sum(R.saldo) as tsld ');
      ZQTempReceber2.SQL.Add(' from recebimento R ');
      ZQTempReceber2.SQL.Add(' join loteamento L on L.idloteamento = R.idloteamento ');
      ZQTempReceber2.SQL.Add(' left join participante P on P.idpaticipante = R.adversa ');

      if  (not empty(copy(XDEBaixasInicio.Text,1,2))) or (not empty(copy(XDERefInicio.Text,1,2))) Then
      begin
        ZQTempReceber2.SQL.Add(' join recbxhist H on R.idrecebimento = H.idrecib ');
        ZQTempReceber2.SQL.Add(' join receb_baixa BX on BX.refbaixa = H.refer ');
      end;

      ZQTempReceber2.SQL.Add(' where 1=1 ');

      if not empty(vardocumento) Then
        ZQTempReceber2.SQL.Add(vardocumento);

      if RGFormato.ItemIndex <> 3 Then
      begin
        if RGContas.ItemIndex = 0 Then
          ZQTempReceber2.SQL.Add(' and R.recpag = '+QuotedStr('R'))
        else
          ZQTempReceber2.SQL.Add(' and R.recpag = '+QuotedStr('P'));
      end;

      if CBDesconsidera.Checked Then
        ZQTempReceber2.SQL.Add(' and (R.numboleto is null or R.numboleto='''')');

      if not empty(VarEmpree) then
        ZQTempReceber2.SQL.Add(' and R.idloteamento in ('+VarEmpree+')');

      if not empty(varquadra) then
        ZQTempReceber2.SQL.Add(' and R.quadralote in ('+varquadra+')');

      if not empty(varDoc) Then
        ZQTempReceber2.SQL.Add(' and R.TipDoc in ('+varDoc+')');

      if CBSaldo.ItemIndex = 1 Then
        ZQTempReceber2.SQL.Add(' and (R.saldo < R.valor or R.saldo is null)');

      if CBSaldo.ItemIndex = 2 Then
        ZQTempReceber2.SQL.Add(' and R.saldo > 0');

      if not empty(copy(XDEEntradaInicio.Text,1,2)) Then
      begin
        ZQTempReceber2.SQL.Add(' and R.DT_Entrada between :dt1 and :dt2');
        ZQTempReceber2.ParamByName('dt1').AsDate := StrToDate(XDEEntradaInicio.DateText);
        ZQTempReceber2.ParamByName('dt2').AsDate := StrToDate(XDEEntradaFinal.DateText);
      end;

      if not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
      begin
        ZQTempReceber2.SQL.Add(' and R.DT_Vencimento between :dt3 and :dt4');
        ZQTempReceber2.ParamByName('dt3').AsDate := StrToDate(XDEVencimentoInicio.DateText);
        ZQTempReceber2.ParamByName('dt4').AsDate := StrToDate(XDEVencimentoFinal.DateText);
      end;

      if not empty(copy(XDEBaixasInicio.Text,1,2)) Then
      begin
        ZQTempReceber2.SQL.Add(' and BX.DT_Rec between :dt5 and :dt6');
        ZQTempReceber2.ParamByName('dt5').AsDate := StrToDate(XDEBaixasInicio.DateText);
        ZQTempReceber2.ParamByName('dt6').AsDate := StrToDate(XDEBaixasFinal.DateText);
      end;

      if not empty(copy(XDERefInicio.Text,1,2)) Then
      begin
        ZQTempReceber2.SQL.Add(' and BX.Dataref between :dt7 and :dt8');
        ZQTempReceber2.ParamByName('dt7').AsDate := StrToDate(XDERefInicio.DateText);
        ZQTempReceber2.ParamByName('dt8').AsDate := StrToDate(XDERefFinal.DateText);
      end;

      if CBDesconsidera.Checked Then
        ZQTempReceber2.SQL.Add(' and not exists (select 1 from remessa_receb RR where RR.idrec = R.idrecebimento)');
      if CBEmpre.Checked then
      begin
        ZQTempReceber2.SQL.Add(' group by COALESCE(NULLIF(L.nomeloteamento, ''''), L.apelido),');
        ZQTempReceber2.SQL.Add(' L.apelido, R.quadralote ');
        ZQTempReceber2.SQL.Add(' order by COALESCE(NULLIF(L.nomeloteamento, ''''), L.apelido), '+VarOrdem);
      end
      else
      begin
        ZQTempReceber2.SQL.Add(' group by R.quadralote ');
        ZQTempReceber2.SQL.Add(' order by '+VarOrdem);
      end;

      ZQTempReceber2.Open;
      ZQTempReceber2.First;



     Vrtexto := '';
     if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then
     Begin
       Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' à '+XDEEntradaFinal.DateText;
     end;
     if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then
     Begin
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' à '+XDEVencimentoFinal.DateText;
     end;
     if CBEmpre.Checked then
     begin
         FrmRelReceb02_total_ab_lt:=nil;
         if FrmRelReceb02_total_ab_lt=nil then
            FrmRelReceb02_total_ab_lt:=TFrmRelReceb02_total_ab_lt.Create(Application);
         FrmRelReceb02_total_ab_lt.RLLabel1.Caption := 'Contas a Receber';
         if RGContas.ItemIndex = 1 Then
            FrmRelReceb02_total_ab_lt.RLLabel1.Caption := 'Contas a Pagar';
         if not empty(vrtexto) Then
            FrmRelReceb02_total_ab_lt.RLLabel2.Caption := 'Período de : '+vrtexto;
         FrmRelReceb02_total_ab_lt.RLReport1.PreviewModal;
         FreeAndNil(FrmRelReceb02_total_ab_lt);
     end
     else
     begin
         FrmRelReceb02_total_ab:=nil;
         if FrmRelReceb02_total_ab=nil then
            FrmRelReceb02_total_ab:=TFrmRelReceb02_total_ab.Create(Application);
         FrmRelReceb02_total_ab.RLLabel1.Caption := 'Contas a Receber';
         if RGContas.ItemIndex = 1 Then
            FrmRelReceb02_total_ab.RLLabel1.Caption := 'Contas a Pagar';
         if not empty(vrtexto) Then
            FrmRelReceb02_total_ab.RLLabel2.Caption := 'Período de : '+vrtexto;
         FrmRelReceb02_total_ab.RLReport1.PreviewModal;
         FreeAndNil(FrmRelReceb02_total_ab);
     end;
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
    ZQTempReceber.DisableControls;
    while not ZQTempReceber.Eof do begin
      JSP1.Position := ZQTempReceber.RecNo;
      for varvezes:=1 to 5 do begin
        varmasc := copy(ZQTempReceber.FieldByName('planomascara').AsString,1,vartama[varvezes]);
        if length(varmasc)<>vartama[varvezes] Then
          break;
        if CDSPlano.Locate('mascara',varmasc,[]) Then
          CDSPlano.Edit
        else
          CDSPlano.Insert;
        CDSPlanomascara.Value := varmasc;
        CDSPlanosoma.Value := CDSPlanosoma.Value + ZQTempReceber.FieldByName('Valor').AsFloat;
        CDSPlano.Post;
      end;
      CDSDetal.Insert;
      CDSDetalmascara.Value := ZQTempReceber.FieldByName('planomascara').AsString;
      CDSDetalvalor.Value := ZQTempReceber.FieldByName('Valor').AsFloat;
      CDSDetaldata.Value := ZQTempReceber.FieldByName('Dt_Vencimento').AsDateTime;
      CDSDetaldetalhe.Value := 'Lanç: '+ZQTempReceber.FieldByName('idrecebimento').Text+'  Participante: '+ZQTempReceber.FieldByName('nome_parte').AsString+'      Adversa: '+ZQTempReceber.FieldByName('adversanome').AsString+chr(13);
      CDSDetaldetalhe.Value := CDSDetaldetalhe.Value + ZQTempReceber.FieldByName('Observ').AsString;
      ZQTempReceber.Next;
    end;
    ZQTempReceber.EnableControls;
    if FrmRelRecebGerencial=nil then
       FrmRelRecebGerencial:=TFrmRelRecebGerencial.Create(Application);

    if CBObsLanc.Checked Then
      FrmRelRecebGerencial.RLBand4.Visible := True
    else
      FrmRelRecebGerencial.RLBand4.Visible := False;
    FrmRelRecebGerencial.RLReport1.PreviewModal;
    FreeAndNil(FrmRelRecebGerencial);
    ZQEmpree.Close;
  end;
  JSP1.Visible := False;
end;

procedure TFrmRelRecebimento.DXBMarcDocClick(Sender: TObject);
Var
  x, varcontador : Integer;
begin
  varcontador := CLBDoc.Count-1;
  if DXBMarcDoc.Caption = 'Desmarcar' Then Begin
    for x := 0 to varcontador do
      CLBDoc.Checked[x] := False;
    DXBMarcDoc.Caption := 'Marcar';
  end
  Else
  Begin
    for x := 0 to varcontador do
      CLBDoc.Checked[x] := True;
    DXBMarcDoc.Caption := 'Desmarcar';
  End;
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

procedure TFrmRelRecebimento.xquadraLoteExit(Sender: TObject);
var
  i: Integer;
  Linha, Quadra, Lote: string;
begin
  if Trim(xquadraLote.Text) <> '' then
  begin
    for i := 0 to xquadraLote.Lines.Count - 1 do
    begin
      Linha := Trim(xquadraLote.Lines[i]);

      // ignora linhas vazias
      if Linha = '' then
        Continue;

      // ====== EXTRAÇÃO ======
      // Exemplo esperado: NO.22.13.13-10
      // Quadra = NO.22.13.13
      // Lote   = 10

      if Pos('-', Linha) > 0 then
      begin
        Quadra := Copy(Linha, 1, Pos('-', Linha) - 1);
        Lote   := Copy(Linha, Pos('-', Linha) + 1, Length(Linha));
      end
      else
      begin
        mensagem('Formato inválido: ' + Linha);
        Continue;
      end;

      // joga nos edits (se quiser manter compatibilidade)
      Equadra.Text := Quadra;
      Elote.Text   := Lote;

      // ====== CONSULTA ======
      ZQVen1.Close;
      ZQVen1.SQL.Clear;
      ZQVen1.SQL.Add('Select * from venda join imovel on idimovel=imovel');
      ZQVen1.SQL.Add(' where quadra = ' + QuotedStr(Quadra));
      ZQVen1.SQL.Add('   and lote   = ' + QuotedStr(Lote));
      ZQVen1.Open;

      if ZQVen1.RecordCount > 0 then
      begin
        CDSQuadraLote.Insert;
        CDSQuadraLoteidlote.Value := ZQVen1.FieldByName('idimovel').AsLargeInt;
        CDSQuadraLoteQuadra.Value := Linha;
        CDSQuadraLote.Post;

        if ZQVen1.FieldByName('datavenda').AsDateTime < XDEEntradaInicio.DateValue then
          XDEEntradaInicio.DateValue := ZQVen1.FieldByName('datavenda').AsDateTime;
      end
      else
      begin
        mensagem('Não encontrado: ' + Linha);
      end;
    end;

    Equadra.Clear;
    Elote.Clear;
    xquadraLote.Lines.Clear;
    xquadraLote.Visible:=true;
    RGQuadraLote.ItemIndex:=0;
    Equadra.SetFocus;
  end
  else
  begin
    if CDSQuadraLote.RecordCount > 0 then
      DXBImprimir.SetFocus;
  end;
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
  xquadraLote.Lines.Clear;
  ZQBaixa.Close;
  ZQTempReceber.Close;
  ZQTempReceber2.close;
  DM_Tabelas.ZQincorp_loteame.close;
  DM_Tabelas.ZQLoteamento.close;
  FrmRelRecebimento := nil;
  Action := CaFree;
end;

procedure TFrmRelRecebimento.GBEscClienteExit(Sender: TObject);
begin
  FrmRelRecebimento.KeyPreview := True
end;

procedure TFrmRelRecebimento.ZQBaixaCalcFields(DataSet: TDataSet);
begin
  if ZQBaixa.FieldByName('percent_usado').AsFloat >0 Then
    ZQBaixa.FieldByName('vrusado').AsFloat := ZQBaixa.FieldByName('Vr_rec').AsFloat*ZQBaixa.FieldByName('percent_usado').AsFloat/100
  else
    ZQBaixa.FieldByName('vrusado').AsFloat := ZQBaixa.FieldByName('Vr_rec').AsFloat;
end;

procedure TFrmRelRecebimento.EadversaExit(Sender: TObject);
begin
  IF (Eadversa.Text)<>emptystr Then Begin
    if AchaCliente(400, 300, Eadversa.Text) then Begin
      CDSAdversaTemp2.Insert;
      CDSAdversaTemp2CodAdversa.Value := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
      CDSAdversaTemp2.Post;
    end;
    Eadversa.Text := '';
    Eadversa.SetFocus;
  end;

end;

procedure TFrmRelRecebimento.EClienteExit(Sender: TObject);
begin
  if not empty(ECliente.Text) Then Begin
    if AchaCLiente(200, 10,ECliente.Text) Then Begin
      CDSCli.Insert;
      CDSCliCCli.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
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

procedure TFrmRelRecebimento.dbxmarcempreClick(Sender: TObject);
Var
  x, varcontador : Integer;
begin
  varcontador := CLBEmpree.Count-1;
  if dbxmarcempre.Caption = 'Desmarcar' Then Begin
    for x := 0 to varcontador do
      CLBEmpree.Checked[x] := False;
    dbxmarcempre.Caption := 'Marcar';
  end
  Else
  Begin
    for x := 0 to varcontador do
      CLBEmpree.Checked[x] := True;
    dbxmarcempre.Caption := 'Desmarcar';
  End;
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
  if ( RGFormato.ItemIndex = 0 ) or ( RGFormato.ItemIndex = 1 ) or (RGFormato.ItemIndex = 4) or ( RGFormato.ItemIndex = 3) or ( RGFormato.ItemIndex = 2 ) or ( RGFormato.ItemIndex = 6 ) Then
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
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
    CBQuadra.Visible:=true;
  end
  else
  begin
    CBDesconsidera.Visible := False;
    GroupBox3.Visible:=true;
    GroupBox7.Visible:=true;
    xresumo.Visible:=false;
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
    CBQuadra.Visible:=false;
  end;
  if RGFormato.ItemIndex = 4 Then
  Begin
    dxbSelecionar.visible:=true;
    cbmensal.Visible:=false;
    CbCarteira.Visible:=false;
    CBDesconsidera.Checked:=true;
    xresumo.Visible:=false;
    CBfinan.Visible:=false;
    GroupBox5.Visible:=true;
  end
  else
  begin
    dxbSelecionar.visible:=false;
    xresumo.Visible:=false;
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
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
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
    CBQuadra.Visible:=true;
    cbtotais.Visible:=false;
  end
  else
  begin
    stotais.Visible:=false;
    xresumo.Visible:=false;
    CBfinan.Visible:=false;
    CBFinanciamento.Text:='';
    GBParcelas.Visible:=false;
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
    CBQuadra.Visible:=false;
    cbtotais.Visible:=false;
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
    CBfinan.Visible:=false;
    CBFinanciamento.Text:='';
    GBParcelas.Visible:=false;
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
    CBQuadra.Visible:=false;
    cbtotais.Visible:=true;
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
    CBfinan.Visible:=false;
    CBFinanciamento.Text:='';
    GBParcelas.Visible:=false;
    CBfinan.Visible:=false;
    GroupBox5.Visible:=true;
    CBQuadra.Visible:=false;
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
    CBfinan.Visible:=true;
    GroupBox5.Visible:=true;
//    CBFinanciamento.Text:='';
//    GBParcelas.Visible:=false;
  end
  else
  begin
    RGOrdem.Items.Clear;
    RGOrdem.Items.Add('Vencimento');
    RGOrdem.Items.Add('Comprador');
    RGOrdem.Items.Add('Quadra/Lote');
    RGOrdem.Items.Add('Tipo');
    if (RGFormato.ItemIndex=6) then
       RGOrdem.Items.Add('Empreendimento');
    RGOrdem.ItemIndex:=0;
    GroupBox5.Visible:=true;
  end;

  if RGFormato.ItemIndex = 6 Then
  Begin
    CBSaldo.ItemIndex:=2;
    CBVerbaixa.ItemIndex:=1;
    GroupBox6.Visible:=false;
    GroupBox8.Visible:=false;
    GroupBox5.Visible:=false;
    stotais.Visible:=true;
    cbmensal.Visible:=true;
    CbCarteira.Visible:=false;
    CBDesconsidera.Checked:=false;
    xresumo.Visible:=false;
    CBfinan.Visible:=false;
    RGOrdem.ItemIndex:=1;
    CBAgrupado.Visible:=true;
    GroupBox1.Caption:='Dt. Reajuste';
    CBQuadra.Visible:=false;
  end
  else
  begin
    CBAgrupado.Visible:=False;
    GroupBox1.Caption:='Entrada';
  end;

  if RGFormato.ItemIndex = 8 Then
  Begin
{    CBSaldo.ItemIndex:=2;
    CBVerbaixa.ItemIndex:=1;
    GroupBox6.Visible:=false;
    GroupBox8.Visible:=false;
    GroupBox5.Visible:=false;
    stotais.Visible:=true;
    cbmensal.Visible:=true;
    CbCarteira.Visible:=false;
    CBDesconsidera.Checked:=false;
    xresumo.Visible:=false;
    CBfinan.Visible:=false;
    RGOrdem.ItemIndex:=1;
    CBAgrupado.Visible:=true;}
    GroupBox1.Caption:='Dt.Juridico';
  end
  else
  begin
//    CBAgrupado.Visible:=False;
    GroupBox1.Caption:='Entrada';
  end;

end;


procedure TFrmRelRecebimento.RGQuadraLoteClick(Sender: TObject);
begin
  if RGQuadraLote.ItemIndex=0 then
     xquadraLote.Visible:=false
  else
     xquadraLote.Visible:=true;
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
      CDSQuadraLoteidlote.Value := ZQVen1.FieldByName('idimovel').AsLargeInt;
      CDSQuadraLoteQuadra.Value := Equadra.Text+'-'+Elote.Text;
      CDSQuadraLote.Post;
      if ZQVen1.FieldByName('datavenda').AsDateTime<XDEEntradaInicio.DateValue Then
        XDEEntradaInicio.DateValue := ZQVen1.FieldByName('datavenda').AsDateTime;
//      if ZQVen1datavenda.Value>XDEEntradaFinal.DateValue Then
//        XDEEntradaFinal.DateValue := ZQVen1datavenda.Value;
    end
    else
      Mensagem('Esta Quadra e lote não foram vendidos!!!!');
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
    AbrirModal(Self, FrmCobrancaBancaria);
    CDS_MarcaTit.IndexName:='';}    
end;

procedure TFrmRelRecebimento.CBQuadraClick(Sender: TObject);
begin
  if CBQuadra.Checked then
     CBEmpre.Visible:=true
  else
     CBEmpre.Visible:=false;
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

procedure TFrmRelRecebimento.cbtotaisClick(Sender: TObject);
begin
  if cbtotais.Checked=true then
  begin
    XDERefInicio.DateText:='01/01/1900';
    XDERefFinal.DateValue:=date;
  end
  else
  begin
    XDERefInicio.Clear;
    XDERefFinal.Clear;
  end;
end;

procedure TFrmRelRecebimento.CBfinanClick(Sender: TObject);
begin
  if CBfinan.Checked=true then
  begin
     GBParcelas.Visible:=true;
     GroupBox3.Visible:=false;
     GroupBox7.Visible:=false;
     GroupBox5.Visible:=false;
     GroupBox1.Visible:=false;
  end
  else
  begin
     CBFinanciamento.text:='';
     GBParcelas.Visible:=false;
     GBParcelas.Visible:=true;
     GroupBox3.Visible:=true;
     GroupBox7.Visible:=true;
     GroupBox5.Visible:=true;
     GroupBox1.Visible:=true;

  end;   
end;


procedure TFrmRelRecebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  { O hook de campos em tempo de execucao preserva este evento e aplica os
    demais lookups depois dele. }
  ZQRecebimento.OnCalcFields := ZQRecebimentoCalcFields;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdata_juridico', 'data_juridico', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdescricao_juridico', 'descricao_juridico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberjuridico', 'juridico', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQBaixa', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaTipdoc_1', 'Tipdoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadescricao', 'descricao', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixatipodoc', 'tipodoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavend_receb', 'vend_receb', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavend_caixa', 'vend_caixa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixareceb_caixa', 'receb_caixa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixareceb_receb', 'receb_receb', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadados_chequ', 'dados_chequ', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaso_avista', 'so_avista', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixalancabanco', 'lancabanco', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasomapaga', 'somapaga', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixalancinclus', 'lancinclus', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasubstituicao_1', 'substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQResumoReceb', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebtipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoReceblocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasubstituicao_1', 'substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQAgruTipo', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoidRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiposaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiposomasaldo', 'somasaldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiporecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiponumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiporefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiponumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTemp2Cl', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cldoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cldoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cltipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Claniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cllocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRateio', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiotipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateionumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreematriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreebairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreepasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreenomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQVenda', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendalote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendalado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendacornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendadisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQVen1', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1lado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomatit', 'somatit', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomapago', 'somapago', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomajuros', 'somajuros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomecliente', 'nomecliente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomeempree', 'nomeempree', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsubstituicao', 'substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomadescontos', 'somadescontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb1', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Vr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Docum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Juros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Desc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1obsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Valor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Observ', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1marcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1RefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1origem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1numboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomecliente', 'nomecliente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomeempree', 'nomeempree', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1substituicao', 'substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Substituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1vr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1sq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1sq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1TipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1sq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1somar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1valor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1valor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Reajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Data_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Proximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Parcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQCh', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChbanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChalias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChsubstituicao_2', 'substituicao_2', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTempReceber2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Observ', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2marcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2origem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2planomascara', 'planomascara', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2adversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2nomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2numboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2Substituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2tvr', 'tvr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2tsld', 'tsld', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2nomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2apelido_loteamento', 'apelido_loteamento', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2nome_loteamento', 'nome_loteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber2', 'ZQTempReceber2comprador_relatorio', 'comprador_relatorio', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQmensal', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalTIPDOC', 'TIPDOC', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalvr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalnomelote', 'nomelote', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalMESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal', 'ZQmensalvr_rel', 'vr_rel', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQmensal2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2MESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2vr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2tipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2DT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2DT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2vr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2juros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQmensal2', 'ZQmensal2valor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRelatorio2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2Mes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2MESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2vr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2tipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2DT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2DT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2Ano', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio2', 'ZQRelatorio2valorpago', 'valorpago', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRelatorio', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioMESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriovr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatoriotipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelatorio', 'ZQRelatorioquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoETipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoERefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoECodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoESubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoETipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoERefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoErefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoErecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoESubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoE', 'ZqresumoEvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZqresumoP', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPvalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZqresumoP', 'ZqresumoPParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQcarteira', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraMESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteirarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiratipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraDT_Entrada', 'DT_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraDT_Vencimento', 'DT_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiradataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraAno', 'Ano', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteirajuros_vr', 'juros_vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiraap', 'ap', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQcarteira', 'ZQcarteiravalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRelCarteira', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraMes', 'Mes', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraMESDESC', 'MESDESC', TWideStringField, fkData, 9, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiravr_rel', 'vr_rel', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteirarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiratipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
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
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRelCarteira', 'ZQRelCarteiraap', 'ap', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZRdiario', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioVR_AB', 'VR_AB', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariorefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarionumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarionomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariojuros_1', 'juros_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiariodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRdiario', 'ZRdiarioApelido', 'Apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQincorp_loteame', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamepercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamenconta', 'nconta', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameconta_v', 'conta_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_conta_v', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameagencia', 'agencia', TWideStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameagencia_v', 'agencia_v', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_agencia_v', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamen_banco', 'n_banco', TWideStringField, fkLookup, 3, 0, False, '', '', '', '', 0, 'codcontabancaria', 'DM_Tabelas.ZQContaBancaria', 'idconta_bancaria', 'n_banco', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamematriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamebairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamepasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamenomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteametipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamecodcontabancaria_1', 'codcontabancaria_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteameapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQincorp_loteame', 'ZQincorp_loteamedigito_dif_1', 'digito_dif_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'zqfinanciamento', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoidloteamento_1', 'idloteamento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentocidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentomatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentodatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentodatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentodataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentologotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentomapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentotestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentotestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentotestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentotestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentobairro', 'bairro', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentopasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentonomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentotipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentoapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'zqfinanciamento', 'zqfinanciamentodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRecebimento', 'ZQRecebimentoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZRTotPago', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosomatit', 'somatit', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosomapago', 'somapago', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosomajuros', 'somajuros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagonomeempree', 'nomeempree', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagonomecliente', 'nomecliente', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagovencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosubstituicao', 'substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagovr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagovalor_parcela', 'valor_parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagorefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagovalor_parcela_1', 'valor_parcela_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoTipDoc_1', 'TipDoc_1', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagoSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodata_juridico', 'data_juridico', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagodescricao_juridico', 'descricao_juridico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagojuridico', 'juridico', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosq_2', 'sq_2', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZRTotPago', 'ZRTotPagosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
