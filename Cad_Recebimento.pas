unit Cad_Recebimento;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, ExtCtrls, XBanner,
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids, DBGrids,
  wwdbdatetimepicker, XDate, XNum, XEdit, DBClient, ComCtrls, TabNotBk,
  XDBNum, JvExControls, JvArrayButton, GradBtn, JvXPCore, JvXPBar,
  XDBDate, FnpNumericEdit, DBDateTimePicker, Vcl.Samples.Gauges;

type
  TFrmCad_Recebimento = class(TForm)
    Label5: TLabel;
    PPainelBotoes: TPanel;
    DXBExcluir: TdxButtonArround;
    DXBEditar: TdxButtonArround;
    DXBIncluir: TdxButtonArround;
    DXBCancelar: TdxButtonArround;
    DXBGravar: TdxButtonArround;
    DXBPesquisar: TdxButtonArround;
    DXBRelatorios: TdxButtonArround;
    DXBFechar: TdxButtonArround;
    DXBUltimo: TdxButtonArround;
    DXBProximo: TdxButtonArround;
    DXBAnterior: TdxButtonArround;
    DXBPrimeiro: TdxButtonArround;
    Pag_Receb: TTabbedNotebook;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Vencimento: TLabel;
    Label8: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    DBELancamento: TDBEdit;
    DBEDocomento: TDBEdit;
    DBEValor: TDBEdit;
    DBMObos: TDBMemo;
    ECliente: TEdit;
    DBCBTipDoc: TDBComboBox;
    DBEOrdem: TDBEdit;
    Label10: TLabel;
    LReg: TLabel;
    Label11: TLabel;
    GBCheque: TGroupBox;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    DBEAgencia: TDBEdit;
    DBEBanco: TDBEdit;
    DBEDono: TDBEdit;
    DBECpf_Cnpj: TDBEdit;
    DBEAlias: TDBEdit;
    Label17: TLabel;
    Label18: TLabel;
    WDDBProrrogado: TwwDBDateTimePicker;
    DS_VIncRec: TDataSource;
    ZQVIncRec: TZQuery;
    DS_Origem: TDataSource;
    ZQOrigem: TZQuery;
    Label15: TLabel;
    Label21: TLabel;
    Label24: TLabel;
    DX: TdxButtonArround;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    Label19: TLabel;
    EContabil: TEdit;
    DBGReceb: TDBGrid;
    DBEIDReceb: TDBEdit;
    DBGrid4: TDBGrid;
    Label22: TLabel;
    DEBConta: TDBEdit;
    EAdversa: TEdit;
    PRecPag: TPanel;
    Label25: TLabel;
    GBVisualizar: TGroupBox;
    CBRec: TCheckBox;
    CBPag: TCheckBox;
    CBAberto: TCheckBox;
    CBFinal: TCheckBox;
    TimeRecebimento: TTimer;
    DBMemo1: TDBMemo;
    dxButton2: TdxButtonArround;
    XDBNumEdit1: TXDBNumEdit;
    Label20: TLabel;
    RGDElete: TRadioGroup;
    DS_TempDivid: TDataSource;
    ZQTempDivid: TZQuery;
    dxButton3: TdxButtonArround;
    DBGrid5: TDBGrid;
    Label23: TLabel;
    DS_Gerou: TDataSource;
    ZQGerou: TZQuery;
    Label26: TLabel;
    Label27: TLabel;
    DS_result: TDataSource;
    CDSResult: TClientDataSet;
    DS_result1: TDataSource;
    ZQResult1: TZQuery;
    CDSResultareceber: TFloatField;
    CDSResultapagar: TFloatField;
    CDSResultrecebido: TFloatField;
    CDSResultpago: TFloatField;
    DS_Feito: TDataSource;
    ZQFeito: TZQuery;
    CDSResultdata: TDateField;
    CDSResultordem: TWideStringField;
    XBanner2: TXBanner;
    XBanner4: TXBanner;
    XBanner6: TXBanner;
    XBanner5: TXBanner;
    XBanner8: TXBanner;
    DBENumero: TDBEdit;
    Label32: TLabel;
    DBXProcesso: TdxButtonArround;
    JBBaixa: TJvXPBar;
    GroupBox1: TGroupBox;
    Label33: TLabel;
    Label34: TLabel;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    XBanner7: TXBanner;
    DBEdit1: TDBEdit;
    Label35: TLabel;
    DBXBaixa: TdxButtonArround;
    JBProcesso: TJvXPBar;
    ZQReceb_Baixa: TZQuery;
    DS_Receb_baixa: TDataSource;
    Panel1: TPanel;
    Panel2: TPanel;
    ZQVenda: TZQuery;
    DS_Venda: TDataSource;
    ZQParcela: TZQuery;
    DS_Parcela: TDataSource;
    DS_Entrada: TDataSource;
    ZQEntrada: TZQuery;
    CDSResultpareceber: TFloatField;
    CDSResultprecebido: TFloatField;
    CDSResultpatrazado: TFloatField;
    WDBEEntrada: TXDBDateEdit;
    WDBEVencimetno: TXDBDateEdit;
    tab_graficos: TTabbedNotebook;
    XBanner3: TXBanner;
    Ecodcli: TEdit;
    XDEfim: TXDateEdit;
    XDEIni: TXDateEdit;
    Label28: TLabel;
    Label31: TLabel;
    Label29: TLabel;
    EParticipante: TEdit;
    Edevedor: TEdit;
    Ecoddeved: TEdit;
    Label30: TLabel;
    dxButton4: TdxButtonArround;
    PnlGrafico1: TPanel;
    XBanner9: TXBanner;
    XBanner10: TXBanner;
    XDEIni1: TXDateEdit;
    Label36: TLabel;
    XDEFim1: TXDateEdit;
    Label37: TLabel;
    Edloteamento: TEdit;
    EdNomeLoteamento: TEdit;
    dxButton1: TdxButtonArround;
    PnlGrafico2: TPanel;
    XDVar1: TXDateEdit;
    Label38: TLabel;
    XBanner11: TXBanner;
    DBcobranca: TDBGrid;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;
    Label40: TLabel;
    ZQRecebimento2: TZQuery;
    DS_Recebimento2: TDataSource;
    nparcelas: TFnpNumericEdit;
    Label41: TLabel;
    vencidos: TXNumEdit;
    Label42: TLabel;
    Label44: TLabel;
    Xmulta: TXNumEdit;
    Label45: TLabel;
    XMora: TXNumEdit;
    XCorrigido: TXNumEdit;
    Label43: TLabel;
    Xpagto: TXDateEdit;
    Panel3: TPanel;
    Label46: TLabel;
    Shape1: TShape;
    Label47: TLabel;
    Shape2: TShape;
    Shape4: TShape;
    CDnegocio: TClientDataSet;
    DS_negocio: TDataSource;
    DBNegociacao: TDBGrid;
    CDnegocionparcela: TIntegerField;
    CDnegociovr_parcela: TFloatField;
    CDnegociovencimento: TDateField;
    CDnegocioJuros: TFloatField;
    CDnegociovr_parcela_corrigida: TFloatField;
    dxButton7: TdxButtonArround;
    dxButton8: TdxButtonArround;
    fpg: TMaskEdit;
    Label48: TLabel;
    Bar1: TProgressBar;
    XN: TXNumEdit;
    Label49: TLabel;
    ZQRecebimento3: TZQuery;
    DS_Recebimento3: TDataSource;
    ZQRecebimento4: TZQuery;
    DS_Recebimento4: TDataSource;
    ListBox1: TListBox;
    CDnegocio2: TClientDataSet;
    DS_negocio2: TDataSource;
    CDnegocio2nparcela: TIntegerField;
    CDnegocio2vencimento: TDateField;
    CDnegocio2vr_parcela: TFloatField;
    CDnegocio2Juros: TFloatField;
    CDnegocio2vr_parcela_corrigida: TFloatField;
    cbmontante: TCheckBox;
    CDnegocioDias: TIntegerField;
    XBanner12: TXBanner;
    PnlGrafico3: TPanel;
    ZQInadimplentes: TZQuery;
    DataZQInadimplentes: TDataSource;
    Gauge3: TGauge;
    dxButton9: TdxButtonArround;
    dxButton10: TdxButtonArround;
    ZQRecebimento_bancario: TZQuery;
    DS_ZQRecebimento_bancario: TDataSource;
    XBanner13: TXBanner;
    DBMemo2: TDBMemo;
    Label50: TLabel;
    Label51: TLabel;
    DBMemo3: TDBMemo;
    Label52: TLabel;
    Lmora: TLabel;
    RGSelecao: TRadioGroup;
    dt_selecao: TDateTimePicker;
    XBanner1: TXBanner;
    XBanner14: TXBanner;
    DBComboBox1: TDBComboBox;
    Label53: TLabel;
    Label54: TLabel;
    DBMemo4: TDBMemo;
    Label55: TLabel;
    ZQVenda_cobr: TZQuery;
    DS_venda_cobr: TDataSource;
    XDBDateEdit1: TXDBDateEdit;
    Gauge1: TGauge;
    dxButton5: TdxButtonArround;
    XDVar: TXDateEdit;
    Gauge2: TGauge;
    Label39: TLabel;
    dxButton6: TdxButtonArround;
    procedure FormShow(Sender: TObject);
    procedure Ativa_campos;
    procedure Desativa_campos;
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure DXBCancelarClick(Sender: TObject);
    procedure DXBIncluirClick(Sender: TObject);
    procedure DXBEditarClick(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure DXBRelatoriosClick(Sender: TObject);
    procedure botoes_setas;
    procedure PrepararLookupsRecebimento;
    procedure AtualizarLookupsRecebimento;
    procedure ECliente1Exit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure Atualiza_Telas;
    procedure Atualiza_DBcobranca;
    procedure FinalizarPesquisaCobranca;
    procedure CobrancaAntesReabrir(DataSet: TDataSet);
    procedure DBGRecebTitleClick(Column: TColumn);
    procedure DBGRecebKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGRecebMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure EClienteExit(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DBMObosExit(Sender: TObject);
    procedure DXBBaixaIncluirClick(Sender: TObject);
    procedure DXBBaixaEditarClick(Sender: TObject);
    procedure DXBBaixaCancelarClick(Sender: TObject);
    procedure DesativaBaixa_Campos;
    procedure ativaBaixa_Campos;
    procedure Calc_Saldo;
    procedure DXBBaixaExcluirClick(Sender: TObject);
    procedure DXBBaixarClick(Sender: TObject);
    procedure DBGrid1Enter(Sender: TObject);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DXBExBaixaClick(Sender: TObject);
    procedure WDDBProrrogadoExit(Sender: TObject);
    procedure DBGrid2CellClick(Column: TColumn);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure EContabilExit(Sender: TObject);
    procedure DBGRecebDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid3CellClick(Column: TColumn);
    procedure DBGrid3MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGrid4MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure EAdversaExit(Sender: TObject);
    procedure DBGRecebEnter(Sender: TObject);
    procedure TimeRecebimentoTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RGDEleteExit(Sender: TObject);
    procedure RGDEleteClick(Sender: TObject);
    procedure DXBExcluirClick(Sender: TObject);
    procedure CBRecClick(Sender: TObject);
    Procedure Filtrar;
    procedure CBAbertoClick(Sender: TObject);
    procedure CBFinalClick(Sender: TObject);
    procedure CBPagClick(Sender: TObject);
    procedure DXClick(Sender: TObject);
    procedure EParticipanteExit(Sender: TObject);
    procedure EdevedorExit(Sender: TObject);
    procedure dxButton4Click(Sender: TObject);
    procedure dxButton5Click(Sender: TObject);
    procedure Pag_RecebClick(Sender: TObject);
    procedure DBXBaixaClick(Sender: TObject);
    procedure JvXPBar1Items1Click(Sender: TObject);
    procedure JBBaixaExit(Sender: TObject);
    procedure JvXPBar1Items0Click(Sender: TObject);
    procedure DBXProcessoClick(Sender: TObject);
    procedure JBProcessoExit(Sender: TObject);
    procedure JBBaixaItems0Click(Sender: TObject);
    procedure JBProcessoItems0Click(Sender: TObject);
    procedure JBProcessoItems2Click(Sender: TObject);
    procedure JBProcessoItems1Click(Sender: TObject);
    procedure XDBNumEdit1Enter(Sender: TObject);
    procedure XDBNumEdit1Exit(Sender: TObject);
    procedure EdNomeLoteamentoEnter(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure dxButton6Click(Sender: TObject);
    procedure tab_graficosClick(Sender: TObject);
    procedure ZQRecebimentoCalcFields(DataSet: TDataSet);
    procedure DBcobrancaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBcobrancaEnter(Sender: TObject);
    procedure fpgEnter(Sender: TObject);
    procedure fpgExit(Sender: TObject);
    procedure dxButton8Click(Sender: TObject);
    procedure ZQRecebimento4CalcFields(DataSet: TDataSet);
    procedure DBcobrancaDblClick(Sender: TObject);
    procedure dxButton7Click(Sender: TObject);
    procedure dxButton9Click(Sender: TObject);
    procedure JBProcessoItems4Click(Sender: TObject);
    procedure dxButton10Click(Sender: TObject);
    procedure Label5MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure JBProcessoItems5Click(Sender: TObject);
    procedure RGSelecaoClick(Sender: TObject);
    procedure dt_selecaoExit(Sender: TObject);
    procedure JBProcessoItems6Click(Sender: TObject);
    procedure JBProcessoItems7Click(Sender: TObject);
    procedure JBProcessoItems8Click(Sender: TObject);
    procedure JBProcessoItems9Click(Sender: TObject);
    procedure JBProcessoItems10Click(Sender: TObject);
  private
    FCobrancaAtualizada: Boolean;
    FTaxasTotaisCobranca: TZQuery;
    FTitulosTotaisCobranca: TZQuery;
    FContextoCobranca: string;
    procedure AbrirTitulosTotaisCobranca;
    procedure PrepararTaxasVendaCobranca(const AIdVenda: Int64);
    procedure PrepararTaxasTituloCobranca;
    procedure GarantirVendaCobranca(const AIdVenda: string);
    function ContextoCobranca: string;
  private
    { Private declarations }

    FPaintBoxGrafico1: TPaintBox;
    FPaintBoxGrafico2: TPaintBox;
    FPaintBoxGrafico3: TPaintBox;
    FPintandoGrafico1: Boolean;
    FPintandoGrafico2: Boolean;
    procedure CriarGraficosRuntime;
    procedure PintarGrafico1(Sender: TObject);
    procedure PintarGrafico2(Sender: TObject);
    procedure PintarGrafico3(Sender: TObject);
    procedure AtualizarGrafico(PaintBox: TPaintBox);
    procedure PrepararLookupsCobranca;
    procedure PreencherLookupsCobranca;
    procedure AtualizarCamposCobranca;

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCad_Recebimento: TFrmCad_Recebimento;


implementation

uses tabelas, Funcoes, Inc_Recebimento, RecebBaixa, UAchaReceb,
  RelRecebimento, PesqRecebimento, ImpRecibo, ReajusteDeParcelas,
  BaixaAutomatica, Aditamento, quitacao, Cessao, principal,
  ReajusteDeParcelas2, PesqCobranca, RelCobranca, RecebBaixa_subst,
  RelCobranca2, Uobs_estorno, trocar_lote, Uparcelanaopaga, Balao, Ucobranca,
  trocar_empreendimento, zerarNossoNumero, uRuntimeFields, ChartGenerator, uSiaiPerformance;

type
  { A API publica de TDataSet nao expoe GetCalcFields, mas este metodo e o
    ciclo seguro que coloca o dataset em dsCalcFields antes de preencher os
    campos calculados. }
  TDataSetCalcFieldsAccess = class(TDataSet);

procedure RecalcularCamposAtuais(ADataSet: TDataSet);
begin
  if (ADataSet = nil) or (not ADataSet.Active) or ADataSet.IsEmpty then
    Exit;
  TDataSetCalcFieldsAccess(ADataSet).GetCalcFields(ADataSet.ActiveBuffer);
end;

procedure CalcularValoresCobranca(const ASaldo, AMultaContrato, AMoraContrato,
  ATabelaPrice: Double; const ADias: Integer; const AUsarTabelaPrice: Boolean;
  out AMulta, AMora, ACorrigido: Double);
var
  LTaxaMora: Double;
begin
  if AUsarTabelaPrice then
    LTaxaMora := ATabelaPrice
  else
    LTaxaMora := AMoraContrato;
  AMora := ExRound(((((ASaldo / 30) * LTaxaMora) / 100) * ADias), 2);
  AMulta := ExRound(ExRound(ASaldo * AMultaContrato, 2) / 100, 2);
  ACorrigido := ASaldo + AMulta + AMora;
end;

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


procedure GetSelectedRecord(DataSet: TDataSet; FieldName: String;
                            BookMarkList: TBookMarkList; Strings : TStrings);
var i: Integer;
    bm: TBookMark;
begin
  Strings.Clear;
  FrmCad_Recebimento.CDnegocio2.Close;
  FrmCad_Recebimento.CDnegocio2.CreateDataSet;
  FrmCad_Recebimento.CDnegocio2.open;

  With DataSet do
  begin
    DisableControls;
    bm := GetBookMark;
    try
      for i := 0 to BookMarkList.Count -1 do
      begin
        BookMark := BookMarkList[i];
        Strings.Add(DataSet.FindField(FieldName).AsString);
        FrmCad_Recebimento.ZQRecebimento.Filtered:=false;
        FrmCad_Recebimento.ZQRecebimento.Filter:='idrecebimento='+quotedstr(DataSet.FindField(FieldName).AsString);
        FrmCad_Recebimento.ZQRecebimento.Filtered:=true;
        if FrmCad_Recebimento.ZQRecebimento.RecordCount>0 then
        begin
          FrmCad_Recebimento.CDnegocio2.Insert;
          FrmCad_Recebimento.CDnegocio2nparcela.value:=FrmCad_Recebimento.ZQRecebimento.FieldByName('Dias').AsInteger;
          if FrmCad_Recebimento.ZQRecebimento.FieldByName('Dt_Vencimento').IsNull then
            FrmCad_Recebimento.CDnegocio2vencimento.Clear
          else
            FrmCad_Recebimento.CDnegocio2vencimento.Value := FrmCad_Recebimento.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
          FrmCad_Recebimento.CDnegocio2vr_parcela.value:=FrmCad_Recebimento.ZQRecebimento.FieldByName('saldo').AsFloat;
          FrmCad_Recebimento.CDnegocio2Juros.value:=FrmCad_Recebimento.ZQRecebimento.FieldByName('Multa_Contrato').AsFloat+FrmCad_Recebimento.ZQRecebimento.FieldByName('mora_dia').AsFloat;
          FrmCad_Recebimento.CDnegocio2vr_parcela_corrigida.value:=FrmCad_Recebimento.ZQRecebimento.FieldByName('parcela_corrigida').AsFloat;
          FrmCad_Recebimento.CDnegocio2.post;
          FrmCad_Recebimento.ZQRecebimento.Filtered:=false;
        end
        else
        begin
          FrmCad_Recebimento.ZQRecebimento.Filtered:=false;
        end;
      end;
    finally
      FrmCad_Recebimento.ZQRecebimento.Filtered:=false;
      GotoBookMark(bm);
      FreeBookMark(bm);
      EnableControls;
    end;
  end; // With DataSet do
end;



procedure TFrmCad_Recebimento.FormShow(Sender: TObject);
begin
  DBELancamento.Enabled := False;
  DBCBTipDoc.Clear;

  if ZQRecebimento_bancario.Active=false then
  begin
    ZQRecebimento_bancario.close;
    ZQRecebimento_bancario.SQL.Clear;
    ZQRecebimento_bancario.SQL.Add(' SELECT r.idrecebimento,r.documento,r.cliente,r.usuario,r.dt_entrada,r.dt_vencimento,r.valor,r.observ,r.VrDoc,r.ordem,r.tipdoc,r.saldo,r.marcar,r.refbaixa,r.refvinda,r.contabil,r.empresa,r.custodaparcela,r.origem,r.adversa,r.recpag,r.numordem,');
    ZQRecebimento_bancario.SQL.Add(' r.idloteamento,r.venda_idvenda,r.quadralote,r.numboleto,r.substituicao,r.sq,r.somar,r.nomeadversa,r.reajustado,r.data_reajuste,r.proximo_reajuste,r.parcelas_fixas,r.juros,r.descontos,r.data_quitacao,r.observ_estorno,r.tip,r.sld_antes_reajuste,');
    ZQRecebimento_bancario.SQL.Add(' r.Percentual_reajuste,r.juridico,r.data_juridico,r.dt_nao_pagou_no_mes,r.descricao_juridico,r.multa,r.mora,il.incorporador_idincorporador,il.loteamento_idloteamento,il.percent_parte,il.codcontabancaria,il.digito_dif,cb.idconta_bancaria,cb.n_banco,');
    ZQRecebimento_bancario.SQL.Add(' cb.n_agencia,cb.n_agencia_v,cb.n_conta,cb.n_conta_v,cb.nomebanco,cb.n_no_banco,cb.nomeresposavel,cb.nometitular,cb.aberta,cb.senha_interna,cb.senha_conta,cb.ativa,cb.jurosemboleto,cb.taxadiaria,cb.valoroupercent,cb.carteira,cb.boleto,cb.idparticipante,');
    ZQRecebimento_bancario.SQL.Add(' cb.n_dif_empreed,cb.convenio,cb.doc_titular,cb.n_sequencial,cb.nosso_numero,cb.apelido,cb.cod_transmissao,cb.complemento,cb.moramensal,cb.variacao,cb.data_maxima_desconto,cb.Perc_descontos ');
    ZQRecebimento_bancario.SQL.Add(' FROM recebimento as r left join incorporador_loteamento as il on il.loteamento_idloteamento=r.idloteamento ');
    ZQRecebimento_bancario.SQL.Add(' left join conta_bancaria as cb on cb.idconta_bancaria=il.codcontabancaria group by r.idloteamento order by r.idloteamento LIMIT 10');
//    ZQRecebimento_bancario.SQL.Add(' SELECT * FROM recebimento as r join incorporador_loteamento as il on il.loteamento_idloteamento=r.idloteamento join conta_bancaria as cb on cb.idconta_bancaria=il.codcontabancaria group by r.idloteamento order by r.idloteamento LIMIT 0, 10');
    ZQRecebimento_bancario.open;
  end;
  if DM_tabelas.ZQTipodoc.Active= false then
  begin
    DM_tabelas.ZQTipodoc.close;
    DM_tabelas.ZQTipodoc.SQL.Clear;
    DM_tabelas.ZQTipodoc.sql.Add('select idtipodocumento,tipodoc,descricao,vend_receb,vend_caixa,receb_receb,receb_caixa,dados_chequ,so_avista,lancabanco,somapaga,lancinclus from tipodocumento');
    DM_tabelas.ZQTipodoc.open;
    DM_tabelas.ZQTipodoc.First;
  end;
  DM_tabelas.ZQTipodoc.First;
  if DM_tabelas.ZQPlanoDeContas.Active=false then
  begin
    DM_tabelas.ZQPlanoDeContas.close;
    DM_tabelas.ZQPlanoDeContas.SQL.clear;
    DM_tabelas.ZQPlanoDeContas.SQL.Add('Select codigo,doccomum,mascara,descricao,classificacao,valor,cod_reduzido,cred_debi,usoativo,CDPlano,DespFixVar from Plano_contas');
    DM_tabelas.ZQPlanoDeContas.open;
  end;
  if DM_tabelas.ZQReBxHi.Active=false then
  begin
    DM_tabelas.ZQReBxHi.close;
    DM_tabelas.ZQReBxHi.SQL.Clear;
    DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist  LIMIT 10');
    DM_tabelas.ZQReBxHi.open;
  end;
  if DM_TAbelas.ZQRecebBxTemp.Active=false then
  begin
    DM_TAbelas.ZQRecebBxTemp.Close;
    DM_TAbelas.ZQRecebBxTemp.SQL.Clear;
    DM_TAbelas.ZQRecebBxTemp.SQL.Add('SELECT H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.valor_parcela,H.data,');
    DM_TAbelas.ZQRecebBxTemp.SQL.Add('       R.idrecebimento,R.documento,R.cliente,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.Observ,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.RefBaixa,');
    DM_TAbelas.ZQRecebBxTemp.SQL.Add('       R.refvinda,R.contabil,R.empresa,R.custodaparcela,R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.numboleto,');
    DM_TAbelas.ZQRecebBxTemp.SQL.Add('       B.idreceb_baixa,B.Dt_rec,B.Vr_rec,B.TipDoc,B.Docum,B.Juros_Vr,B.Desc_Vr,B.RefBaixa,B.obsebx,B.Codcontabil,B.dataref,B.vencimento,B.valor_parcela ');
    DM_TAbelas.ZQRecebBxTemp.SQL.Add('  FROM recbxhist as H left join recebimento as R ON R.idrecebimento=H.idrecib join receb_baixa as B ON B.refbaixa=H.refer  LIMIT 10');
    DM_TAbelas.ZQRecebBxTemp.open;
  end;
  if DM_Tabelas.ZQCheque.Active=false then
  begin
    DM_Tabelas.ZQCheque.Close;
    DM_Tabelas.ZQCheque.SQL.Clear;
    DM_Tabelas.ZQCheque.SQL.Add('select Banco,Dono,CPF_CNPJ,valor,emissao,deposito,alias,prorrogado,conta,idrecebimento,idpagamento,idcheque,idvinculo,numero,agencia,Substituicao,sq from cheque_rec');
    DM_Tabelas.ZQCheque.open;
  end;

  { Os campos nomecli, adversanome e nome_loteamento são calculados a
    partir destes dois datasets. Eles precisam estar abertos antes do
    primeiro posicionamento do grid para que First/Last também calculem o
    registro atual. }
  PrepararLookupsRecebimento;

  if DM_tabelas.ZQRecebimento.Active=false then
  begin
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento  r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) order by r.DT_Vencimento LIMIT 200, 50');
    DM_tabelas.ZQRecebimento.Open;
//     DM_tabelas.ZQRecebimento.Open;
  end;
  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_Tabelas.ZQContaBancaria.Close;
  DM_Tabelas.ZQContaBancaria.SQL.Clear;
  DM_Tabelas.ZQContaBancaria.SQL.Add('Select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,nometitular,aberta,senha_interna,senha_conta,ativa,');
  DM_Tabelas.ZQContaBancaria.SQL.Add('        jurosemboleto,taxadiaria,valoroupercent,carteira,boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,Apelido,cod_transmissao,');
  DM_Tabelas.ZQContaBancaria.SQL.Add('        complemento,moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote ');
  DM_Tabelas.ZQContaBancaria.SQL.Add('   from conta_bancaria');
  DM_Tabelas.ZQContaBancaria.Open;
//  DM_Tabelas.ZQMovBancaria.Open;
  ZQGerou.Open;

  DM_Tabelas.ZQRecebimento.Last;
  AtualizarLookupsRecebimento;
  DBGReceb.SetFocus;
  botoes_setas;
  Pag_Receb.PageIndex := 0;
  if Pag_Receb.PageIndex = 0 Then
  begin
    ECliente.Text := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
    EAdversa.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
    EContabil.Text:= DM_Tabelas.ZQRecebimento.FieldByName('contabil').Text;
  end;
  Desativa_campos;
  Panel1.Visible:=false;
  Panel2.Visible:=false;
  Frm_principal.Panel1.Visible:=false;

  DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').AsString,[]);
  IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    GBCheque.Visible := True
  else
    GBCheque.Visible := False;

  Application.ProcessMessages;

end;

procedure TFrmCad_Recebimento.Ativa_campos;
Begin
  TimeRecebimento.Enabled := false;
  GBVisualizar.Enabled := False;
  DBXBaixa.Enabled := False;
  DXBPrimeiro.Enabled := False;
  DXBAnterior.Enabled := False;
  DXBProximo.Enabled := False;
  DXBUltimo.Enabled := False;
  DXBIncluir.Enabled := False;
  DXBEditar.Enabled := False;
  DXBExcluir.Enabled := False;
  DXBPesquisar.Enabled := False;
  DXBRelatorios.Enabled := False;
  DXBGravar.Enabled := True;
  DXBCancelar.Enabled := True;
  DBGReceb.Enabled := False;
  DBEOrdem.ReadOnly := False;
  DBCBTipDoc.ReadOnly := False;
  DBEDocomento.ReadOnly := False;
  WDBEEntrada.ReadOnly := False;
  WDBEVencimetno.ReadOnly := False;
  DBEValor.ReadOnly := False;
  ECliente.ReadOnly := False;
  EAdversa.ReadOnly := False;
  DBMObos.ReadOnly := False;
  DBEBanco.ReadOnly := False;
  DBEAgencia.ReadOnly := False;
  DEBConta.ReadOnly := False;
  DBECpf_Cnpj.ReadOnly := False;
  DBEDono.ReadOnly := False;
  DBEAlias.ReadOnly := False;
  WDDBProrrogado.ReadOnly := False;
End;

procedure TFrmCad_Recebimento.Desativa_campos;
Begin
//  TimeRecebimento.Enabled := True;
  GBVisualizar.Enabled := True;
  DBXBaixa.Enabled := True;
  PPainelBotoes.Enabled := True;
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := True;
  DXBUltimo.Enabled := True;
  DXBIncluir.Enabled := True;
  DXBEditar.Enabled := True;
  DXBExcluir.Enabled := True;
  DXBPesquisar.Enabled := True;
  DXBRelatorios.Enabled := True;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := False;
  DBGReceb.Enabled := True;
  DBELancamento.ReadOnly := True;
  DBEOrdem.ReadOnly := True;
  DBCBTipDoc.ReadOnly := True;
  DBEDocomento.ReadOnly := True;
  WDBEEntrada.ReadOnly := True;
  WDBEVencimetno.ReadOnly := True;
  DBEValor.ReadOnly := True;
  ECliente.ReadOnly := True;
  EAdversa.ReadOnly := True;
  DBMObos.ReadOnly := True;
  DBEBanco.ReadOnly := True;
  DBEAgencia.ReadOnly := True;
  DEBConta.ReadOnly := True;
  DBECpf_Cnpj.ReadOnly := True;
  DBEDono.ReadOnly := True;
  DBEAlias.ReadOnly := True;
  WDDBProrrogado.ReadOnly := True;
End;

procedure TFrmCad_Recebimento.DXBPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQRecebimento.First;
  AtualizarLookupsRecebimento;
  botoes_setas;
  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQRecebimento.Prior;
  AtualizarLookupsRecebimento;
  botoes_setas;
  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQRecebimento.Next;
  AtualizarLookupsRecebimento;
  botoes_setas;
  DBGReceb.SetFocus;  
end;

procedure TFrmCad_Recebimento.DXBUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQRecebimento.Last;
  AtualizarLookupsRecebimento;
  botoes_setas;
  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBGravarClick(Sender: TObject);
var
ql:string;
begin
  if not (DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit]) Then Begin
    mensagem('Houve problema de controle de tabelas...  Registro não será gravado...');
    Desativa_campos;
    DBGReceb.SetFocus;
    exit;
  end;
  ql:=DM_tabelas.ZQRecebimento.FieldByName('quadralote').AsString;
  if not Verif_senha('Financeiro','Gravar ','Documento: '+DM_tabelas.ZQRecebimento.FieldByName('documento').AsString+'  '+DM_tabelas.ZQRecebimento.FieldByName('ordem').AsString+'  Interno: '+DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text) then exit;
  DM_tabelas.ZQRecebimento.Post;
  Desativa_campos;
  Panel1.Visible:=True;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  if DM_TAbelas.ZQRecebimento.active=true then
     DM_Tabelas.ZQRecebimento.Refresh;

  // atualizar venda
  if DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString='R' then
  begin
    if simnao('Deseja Atualizar valor Final da Venda ?','S') then
    begin
      ZQVenda.Close;
      ZQVenda.SQL.Clear;
      ZQVenda.SQL.Add('Select  idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao from venda where idvenda='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Text));
      ZQVenda.open;
      if ZQVenda.RecordCount>0 then
      begin
        ZQEntrada.Close;
        ZQEntrada.SQL.Clear;
        ZQEntrada.SQL.Add('Select idrecebimento,valor,quadralote,documento,ordem,cliente, venda_idvenda, sum(valor) as entrada from recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.quadralote='+quotedstr(ql)+' and r.documento like  ''%-E-%''  group by r.quadralote  order by r.ordem');
        ZQEntrada.open;

        ZQParcela.Close;
        ZQParcela.SQL.Clear;
        ZQParcela.SQL.Add('Select idrecebimento,valor,quadralote,documento,ordem,cliente, venda_idvenda, sum(valor) as parcela from recebimento as r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda)  and r.quadralote='+quotedstr(ql)+' and r.documento like  ''%-P-%''  group by r.quadralote  order by r.ordem');
        ZQParcela.open;

        ZQVenda.Edit;
        ZQVenda.FieldByName('valorvenda').AsFloat:=ZQEntrada.FieldByName('entrada').AsFloat+ZQParcela.FieldByName('parcela').AsFloat;
        ZQVenda.Post;
      end;
      ZQVenda.close;
      ZQEntrada.close;
      ZQParcela.close;
    end;
  end;
// até aqui


  Panel1.Visible:=false;
  Panel2.Visible:=false;
  Application.ProcessMessages;

  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Financeiro','Cancelar ','Documento: '+DM_tabelas.ZQRecebimento.FieldByName('documento').AsString+'  '+DM_tabelas.ZQRecebimento.FieldByName('ordem').AsString+'  Interno: '+DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text) then exit;
  DM_tabelas.ZQRecebimento.Cancel;
  Desativa_campos;
end;

procedure TFrmCad_Recebimento.DXBIncluirClick(Sender: TObject);
begin
  if Frm_Inc_Recebimento=nil then
     Frm_Inc_Recebimento:=TFrm_Inc_Recebimento.Create(Self);
  Frm_Inc_Recebimento.Label20.Caption := 'DIRETA';
  Frm_Inc_Recebimento.Label23.Caption := '';
  AbrirModal(Self, Frm_Inc_Recebimento);
  Frm_Inc_Recebimento:=nil;
  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  if DM_TAbelas.ZQRecebBxTemp.active=true then
     DM_TAbelas.zqrecebbxTemp.Refresh;
  Panel1.Visible:=false;
  Panel2.Visible:=false;  
  Application.ProcessMessages;


  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBEditarClick(Sender: TObject);
begin
  if not Verif_senha('Financeiro','Editar ','Documento: '+DM_tabelas.ZQRecebimento.FieldByName('documento').AsString+'  '+DM_tabelas.ZQRecebimento.FieldByName('ordem').AsString+'  Interno: '+DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text) then exit;
  Ativa_campos;
  DM_tabelas.ZQRecebimento.Edit;
  DBEOrdem.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBPesquisarClick(Sender: TObject);
begin
//  FrmAchaReceb.showmodal;
  FrmPesqRecebimento.Top := FrmCad_Recebimento.Top+100;
  FrmPesqRecebimento.Left := FrmCad_Recebimento.Left;
  AbrirModal(Self, FrmPesqRecebimento);
  FrmPesqRecebimento.ZQTempCliReceb.close;
  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBRelatoriosClick(Sender: TObject);
begin
  if not Verif_senha('Financeiro','Relatório ','') then exit;
  if FrmRelRecebimento=nil then
     FrmRelRecebimento:=TFrmRelRecebimento.Create(Self);
  FrmRelRecebimento.Top := FrmCad_Recebimento.Top+100;
  FrmRelRecebimento.Left := FrmCad_Recebimento.Left;
//  FrmRelRecebimento.showmodal;
  AbrirModal(Self, FrmRelRecebimento);
  FrmRelRecebimento:=nil;
end;

procedure TFrmCad_Recebimento.botoes_setas;
Begin
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  if DM_tabelas.ZQRecebimento.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  if DM_tabelas.ZQRecebimento.RecNo = DM_tabelas.ZQRecebimento.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
//  Atualiza_Telas;
End;

procedure TFrmCad_Recebimento.PrepararLookupsRecebimento;
begin
  if DM_Tabelas = nil then
    Exit;

  { Não altere SQL nem feche datasets que possam estar sendo usados por
    outras telas. Apenas abra os lookups quando a tela financeira precisar
    deles. }
  if not DM_Tabelas.ZqParticipante.Active then
    DM_Tabelas.ZqParticipante.Open;
  if not DM_Tabelas.ZQLoteamento.Active then
    DM_Tabelas.ZQLoteamento.Open;
end;

procedure TFrmCad_Recebimento.AtualizarLookupsRecebimento;
var
  LNomeCliente: string;
  LNomeComprador: string;
begin
  if (DM_Tabelas = nil) or
     (not DM_Tabelas.ZQRecebimento.Active) or
     DM_Tabelas.ZQRecebimento.IsEmpty then
    Exit;

  PrepararLookupsRecebimento;

  LNomeCliente := DM_Tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  LNomeComprador := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;

  ECliente.Text := LNomeCliente;
  EAdversa.Text := LNomeComprador;

  DBGReceb.Invalidate;
end;

procedure TFrmCad_Recebimento.ECliente1Exit(Sender: TObject);
begin
  If DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit] Then Begin
    if not AchaCLiente(450, 140,Ecliente.Text) Then Begin
      Ecliente.SetFocus;
      Exit;
    End;
    DM_tabelas.ZQRecebimento.FieldByName('cliente').AsInteger := DM_tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
    Ecliente.Text := DM_tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
  end;
end;

procedure TFrmCad_Recebimento.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCad_Recebimento.Atualiza_Telas;
Var
  Varrefer, Varidrec : string;
Begin
  Varrefer := '(0)';
  Varidrec := '(0)';
  LABEL25.CAPTION := 'R'+CHR(13)+'E'+CHR(13)+'C'+CHR(13)+'E'+CHR(13)+'B'+CHR(13)+'I'+CHR(13)+'M'+CHR(13)+'E'+CHR(13)+'N'+CHR(13)+'T'+CHR(13)+'O';
  PRecPag.Color := ClRed;
  IF DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString = 'P' Then Begin
    LABEL25.CAPTION := 'P'+CHR(13)+'A'+CHR(13)+'G'+CHR(13)+'A'+CHR(13)+'M'+CHR(13)+'E'+CHR(13)+'N'+CHR(13)+'T'+CHR(13)+'O';
    PRecPag.Color := $0000BB00;
  end;
  LReg.Caption := inttostr(DM_tabelas.ZQRecebimento.RecordCount);
  ECliente.Text := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Text;
  EAdversa.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EContabil.Text :='';
  if DM_tabelas.ZQPlanodeContas.Locate('codigo',DM_tabelas.ZQRecebimento.FieldByName('contabil').AsInteger,[]) then
    EContabil.Text := DM_tabelas.ZQPlanoDeContas.FieldByName('codigo').Text+' - '+DM_tabelas.ZQPlanoDeContas.FieldByName('descricao').AsString+'  -  '+masccontabil(DM_tabelas.ZQPlanoDeContas.FieldByName('mascara').AsString);
  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_baixa where refBaixa='+quotedstr(DM_tabelas.ZQRecebimento.FieldByName('RefBaixa').Text));
  DM_tabelas.ZQReceb_Baixa.Open;
  ZQOrigem.Close;
  ZQOrigem.SQL.Clear;
  ZQOrigem.SQL.Add('select idrecebimento,documento,dt_entrada,dt_vencimento,valor,tipdoc,recpag,refbaixa,refvinda ');
  ZQOrigem.SQL.Add('from recebimento where refbaixa='+quotedstr(DM_tabelas.ZQRecebimento.FieldByName('refvinda').Text));
  ZQOrigem.Open;
  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReBxHi.SQL.Clear;
  DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from recbxhist where idrecib='+quotedstr(DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
  DM_tabelas.ZQReBxHi.Open;

  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  if DM_TAbelas.ZQRecebBxTemp.active=true then
     DM_TAbelas.ZQRecebBxTemp.Refresh;
  Panel1.Visible:=false;
  Panel2.Visible:=false;  
  Application.ProcessMessages;


  if DM_tabelas.ZQReBxHi.RecordCount>0 then Begin
    DM_tabelas.ZQReBxHi.First;
    Varrefer := '(';
    DM_tabelas.ZQReBxHi.DisableControls;
    while not DM_tabelas.ZQReBxHi.Eof do Begin
      Varrefer := Varrefer + DM_tabelas.ZQReBxHi.FieldByName('refer').Text+',';
      DM_tabelas.ZQReBxHi.Next;
    end;
    DM_tabelas.ZQReBxHi.EnableControls;
    Varrefer := copy(Varrefer,1,length(Varrefer)-1)+')';
    DM_tabelas.ZQReBxHi.close;
    DM_tabelas.ZQReBxHi.SQL.Clear;
    DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from recbxhist where refer in '+Varrefer);
    DM_tabelas.ZQReBxHi.Open;
    DM_tabelas.ZQReBxHi.First;
    Varidrec := '(';
    DM_tabelas.ZQReBxHi.DisableControls;
    while not DM_tabelas.ZQReBxHi.Eof do Begin
      VAridrec := Varidrec + DM_tabelas.ZQReBxHi.FieldByName('idrecib').Text+',';
     DM_tabelas.ZQReBxHi.Next;
    end;
    DM_tabelas.ZQReBxHi.EnableControls;
    Varidrec := copy(Varidrec,1,length(Varidrec)-1)+')';
    DM_tabelas.ZQReBxHi.Close;    
    DM_tabelas.ZQReBxHi.SQL.Clear;
    DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from recbxhist where idrecib='+DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text);
    DM_tabelas.ZQReBxHi.Open;
  end;
  ZQVincRec.Close;
  ZQVincRec.SQL.Clear;
  ZQVincRec.SQL.Add('select  idrecebimento,documento, ordem,quadralote,refbaixa,refvinda,tipdoc,valor,saldo,dt_vencimento,recpag from recebimento where IdRecebimento in '+Varidrec+' order by idrecebimento');
  ZQVincRec.Open;

  { The generated-title grid needs these display fields in the result set.
    The previous query supplied only the internal link fields. }
  ZQGerou.Close;
  ZQGerou.SQL.Clear;
  ZQGerou.SQL.Add('select R.idrecebimento,H.idrecib,R.quadralote,R.refbaixa,R.refvinda,H.refer,');
  ZQGerou.SQL.Add('       R.ordem,coalesce(P.nome_parte,R.nomeadversa) as nomecli,R.valor ');
  ZQGerou.SQL.Add('from recbxhist H join recebimento R on R.refvinda=H.refer ');
  ZQGerou.SQL.Add('left join participante P on P.idpaticipante=R.cliente ');
  ZQGerou.SQL.Add('where H.idrecib='+quotedstr(DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text)+' order by R.idrecebimento');
  ZQGerou.Open;
  Label26.Caption := '('+inttostr(ZQGerou.RecordCount)+') Registros';

  DM_TAbelas.ZQRecebBxTemp.Close;
  DM_TAbelas.ZQRecebBxTemp.SQL.Clear;
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('SELECT H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.valor_parcela,H.data,');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('       R.idrecebimento,R.documento,R.cliente,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor AS Valor_1,R.Observ,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.RefBaixa,');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('       R.refvinda,R.contabil,R.empresa,R.custodaparcela,R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.numboleto,');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('       B.idreceb_baixa,B.Dt_rec,B.Vr_rec,B.TipDoc AS TipDoc_1,B.Docum,B.Juros_Vr,B.Desc_Vr,B.RefBaixa AS RefBaixa_1,B.obsebx,B.Codcontabil,B.dataref,B.vencimento,B.valor_parcela AS valor_parcela_1,');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('       PC.descricao AS desccontabil ');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('FROM recbxhist H join recebimento R ON R.idrecebimento=H.idrecib ');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('join receb_baixa B ON B.refbaixa=H.refer ');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('left join plano_contas PC ON PC.codigo=B.Codcontabil ');
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('where H.refer in '+Varrefer+' order by H.refer,H.idrecib,H.idrecbxhist');
  DM_TAbelas.ZQRecebBxTemp.open;
  RecalcularCamposAtuais(DM_TAbelas.ZQRecebBxTemp);


  Label27.Caption := '('+inttostr(ZQVincRec.RecordCount)+') Registros';
  DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').AsString,[]);
  IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    GBCheque.Visible := True
  else
    GBCheque.Visible := False;
end;

procedure TFrmCad_Recebimento.PrepararLookupsCobranca;
begin
  { A cobrança usa campos calculados que dependem destes dois lookups.
    Abra-os somente neste fluxo, sem alterar o estado global das demais telas. }
  if not DM_Tabelas.ZqParticipante.Active then
        DM_Tabelas.ZqParticipante.Open;
  if not DM_Tabelas.ZQLoteamento.Active then
       DM_Tabelas.ZQLoteamento.Open;
end;

procedure TFrmCad_Recebimento.PreencherLookupsCobranca;
var
  LChave, LDestino, LOrigem: TField;
  LValor, LNome, LDocumento: Variant;
  LStep: UInt64;
begin
  if (not ZQRecebimento.Active) or (DM_Tabelas = nil) then
    Exit;
  if not (ZQRecebimento.State in [dsEdit, dsInsert, dsCalcFields]) then
    Exit;

  LNome := Null;
  LDocumento := Null;
  LChave := ZQRecebimento.FindField('adversa');
  if (LChave <> nil) and (not LChave.IsNull) and (LChave.AsLargeInt <> 0) and DM_Tabelas.ZqParticipante.Active then
  begin
    { Um unico acesso ao participante, sem reter valores entre operacoes. }
    LStep := PerformanceStart;
    try
      LValor := DM_Tabelas.ZqParticipante.Lookup('idpaticipante', LChave.AsLargeInt, 'nome_parte;doc1');
    finally
      PerformanceSample(smLookupParticipante, LStep);
    end;
    if VarIsArray(LValor) then
    begin
      LNome := LValor[0];
      LDocumento := LValor[1];
    end;
  end;

  LDestino := ZQRecebimento.FindField('adversanome');
  if LDestino <> nil then
  begin
    LDestino.Clear;
    if not VarIsNull(LNome) and not VarIsEmpty(LNome) then
      LDestino.AsString := VarToStr(LNome);
    { Preserva o nome do titulo antigo quando o cadastro nao fornece nome. }
    if LDestino.AsString = '' then
    begin
      LOrigem := ZQRecebimento.FindField('nomeadversa');
      if (LOrigem <> nil) and (not LOrigem.IsNull) then
        LDestino.AsString := LOrigem.AsString;
    end;
  end;

  LDestino := ZQRecebimento.FindField('CPF');
  if LDestino <> nil then
  begin
    LDestino.Clear;
    if not VarIsNull(LDocumento) and not VarIsEmpty(LDocumento) then
      LDestino.AsString := VarToStr(LDocumento);
  end;

  LChave := ZQRecebimento.FindField('idloteamento');
  LDestino := ZQRecebimento.FindField('nome_loteamento');
  if LDestino <> nil then
  begin
    LDestino.Clear;
    if (LChave <> nil) and (not LChave.IsNull) and
       (LChave.AsLargeInt <> 0) and DM_Tabelas.ZQLoteamento.Active then
    begin
      LStep := PerformanceStart;
      try
        LValor := DM_Tabelas.ZQLoteamento.Lookup('idloteamento', LChave.AsLargeInt, 'apelido');
      finally
        PerformanceSample(smLookupLoteamento, LStep);
      end;
      if not VarIsNull(LValor) and not VarIsEmpty(LValor) then
          LDestino.AsString := VarToStr(LValor);
    end;
  end;
end;

procedure TFrmCad_Recebimento.CobrancaAntesReabrir(DataSet: TDataSet);
begin
  FCobrancaAtualizada := False;
  ZQVenda_cobr.Close;
end;

function TFrmCad_Recebimento.ContextoCobranca: string;
begin
  Result := '';
  if ZQRecebimento.Active and not ZQRecebimento.IsEmpty then
     Result := ZQRecebimento.FieldByName('adversa').AsString + '|' + Label52.Caption + '|' + Lmora.Caption + '|' + DateToStr(Date);
end;

procedure TFrmCad_Recebimento.GarantirVendaCobranca(const AIdVenda: string);
var
  LInicio: UInt64;
begin
  { Preserva o dataset local tambem utilizado pela negociacao. }
  if ZQVenda_cobr.Active and not ZQVenda_cobr.IsEmpty and (ZQVenda_cobr.FieldByName('idvenda').Text = AIdVenda) then
    Exit;
  LInicio := PerformanceStart;
  try
    ZQVenda_cobr.Close;
    ZQVenda_cobr.SQL.Text := 'Select idvenda,multa,mora from venda where idvenda=' + QuotedStr(AIdVenda);
    ZQVenda_cobr.Open;
  finally
    PerformanceElapsed('Cobranca: consultar venda para calculo', LInicio);
  end;
end;

procedure TFrmCad_Recebimento.AtualizarCamposCobranca;
var
  LIdRecebimento: Int64;
  LNomeLoteamento: string;
  LComprador: string;
  LCPF: string;
  LInicio, LStep: UInt64;
begin
  if (not ZQRecebimento.Active) or ZQRecebimento.IsEmpty then
    Exit;

  PerformanceCheckpoint('Cobranca: inicio do recalculo completo');
  FlushPerformanceLog;
  LInicio := PerformanceStart;
  LIdRecebimento := ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
  ZQRecebimento.DisableControls;
  try
    LStep := PerformanceStart;
    ZQRecebimento.First;
    PerformanceSample(smPrimeiroTitulo, LStep);
    while not ZQRecebimento.Eof do
    begin
      { First e Next ja executam OnCalcFields quando AutoCalcFields esta
        habilitado. Nao chame GetCalcFields novamente: isso repetia todo o
        calculo e os lookups do registro sem alterar o resultado. }
      LNomeLoteamento := ZQRecebimento.FieldByName('nome_loteamento').AsString;
      LComprador := ZQRecebimento.FieldByName('adversanome').AsString;
      LCPF := ZQRecebimento.FieldByName('CPF').AsString;
      LStep := PerformanceStart;
      ZQRecebimento.Next;
      PerformanceSample(smProximoTitulo, LStep);
    end;
  finally
    LStep := PerformanceStart;
    ZQRecebimento.EnableControls;
    PerformanceSample(smReativarControles, LStep);
  end;
  LStep := PerformanceStart;
  ZQRecebimento.Locate('idrecebimento', LIdRecebimento, []);
  PerformanceSample(smRestaurarTitulo, LStep);
  PerformanceElapsed('Cobranca: recalcular todos os titulos', LInicio);
  FlushPerformanceLog;
end;

procedure TFrmCad_Recebimento.AbrirTitulosTotaisCobranca;
var
  LSQL: string;
begin
  if FTitulosTotaisCobranca = nil then
  begin
    FTitulosTotaisCobranca := TZQuery.Create(Self);
    FTitulosTotaisCobranca.Connection := ZQRecebimento4.Connection;
  end;
  FTitulosTotaisCobranca.Close;
  { A origem e exatamente o SQL ja montado para ZQRecebimento4; somente os
    campos que participam da soma sao trazidos. }
  LSQL := StringReplace(ZQRecebimento4.SQL.Text, 'Select *',
    'Select venda_idvenda,Dt_Vencimento,saldo', [rfReplaceAll, rfIgnoreCase]);
  FTitulosTotaisCobranca.SQL.Text := LSQL;
  FTitulosTotaisCobranca.ParamByName('dt').AsDate :=
    ZQRecebimento4.ParamByName('dt').AsDate;
  FTitulosTotaisCobranca.Open;
end;

procedure TFrmCad_Recebimento.PrepararTaxasVendaCobranca(
  const AIdVenda: Int64);
begin
  { Consulta exclusiva dos totais: nao muda a venda usada pela negociacao. }
  if FTaxasTotaisCobranca = nil then
  begin
    FTaxasTotaisCobranca := TZQuery.Create(Self);
    FTaxasTotaisCobranca.Connection := ZQRecebimento4.Connection;
    FTaxasTotaisCobranca.SQL.Text :=
      'select idvenda,multa,mora,tabela_price from venda where idvenda=:id';
  end;
  if FTaxasTotaisCobranca.Active and not FTaxasTotaisCobranca.IsEmpty and
     (FTaxasTotaisCobranca.FieldByName('idvenda').AsLargeInt = AIdVenda) then
    Exit;
  FTaxasTotaisCobranca.Close;
  FTaxasTotaisCobranca.ParamByName('id').AsLargeInt := AIdVenda;
  FTaxasTotaisCobranca.Open;
  if FTaxasTotaisCobranca.IsEmpty then
    raise EDatabaseError.Create(
      'Venda do titulo nao encontrada para calcular a cobranca.');
end;

procedure TFrmCad_Recebimento.PrepararTaxasTituloCobranca;
begin
  PrepararTaxasVendaCobranca(
    ZQRecebimento4.FieldByName('venda_idvenda').AsLargeInt);
end;
procedure TFrmCad_Recebimento.FinalizarPesquisaCobranca;
var
  LStep: UInt64;
begin
  { O Refresh invalida os calculos. Concluir a atualizacao somente depois dele,
    mesmo quando SetFocus nao dispara OnEnter porque o grid ja tem o foco. }
  LStep := PerformanceStart;
  ZQRecebimento.Refresh;
  PerformanceElapsed('Cobranca: Refresh apos pesquisa', LStep);
  DBcobrancaEnter(nil);
end;

procedure TFrmCad_Recebimento.Atualiza_DBcobranca;
var
Zdias: Integer;
LDias: Integer;
LMulta, LMoraCalculada, LCorrigido: Double;
LInicio, LStep: UInt64;
Begin
  FCobrancaAtualizada := False;
  if FTaxasTotaisCobranca <> nil then
    FTaxasTotaisCobranca.Close;
  if FTitulosTotaisCobranca <> nil then
    FTitulosTotaisCobranca.Close;
  LInicio := PerformanceStart;
  try
  LStep := PerformanceStart;
  PrepararLookupsCobranca;
  PerformanceElapsed('Cobranca: abrir bases dos lookups', LStep);
  nparcelas.Value:=0;
  vencidos.Value:=0;
  Xmulta.Value:=0;
  XMora.Value:=0;
  XCorrigido.Value:=0;
  Zdias:=0;
  ZQRecebimento2.close;
  ZQRecebimento2.SQL.Clear;
  // 27/08/2012 tony pediu pra mudar
//  ZQRecebimento2.SQL.Add('Select adversa,quadralote,saldo, count(quadralote) as parcelas, sum(saldo) as titulos from Recebimento where Dt_Vencimento< :dt and saldo>''0'' and quadralote='+quotedstr(ZQRecebimentoquadralote.value)+' group by quadralote');
  if empty(Label52.Caption) then
  begin
    ZQRecebimento2.SQL.Add('Select venda_idvenda,multa,mora,idloteamento,adversa,quadralote,saldo, count(quadralote) as parcelas, sum(saldo) as titulos from Recebimento where Dt_Vencimento< :dt and saldo>''0'' and adversa='+quotedstr(ZQRecebimento.FieldByName('adversa').text)+' group by adversa');
  end
  else
  begin
   ZQRecebimento2.SQL.Add('Select venda_idvenda,multa,mora,idloteamento,adversa,quadralote,saldo, count(quadralote) as parcelas, sum(saldo) as titulos from Recebimento where idloteamento='+quotedstr(Label52.Caption)+' and Dt_Vencimento< :dt and saldo>''0'' and adversa='+quotedstr(ZQRecebimento.FieldByName('adversa').text)+' group by adversa');
  end;
  ZQRecebimento2.ParamByName('dt').AsDate:=date;
  LStep := PerformanceStart;
  ZQRecebimento2.Open;
  PerformanceElapsed('Cobranca: consulta de quantidade e saldo', LStep);
  ZQRecebimento2.First;
  nparcelas.Value:=ZQRecebimento2.FieldByName('parcelas').AsLargeInt;
  vencidos.Value:=ZQRecebimento2.FieldByName('titulos').AsFloat;
  Xpagto.DateText:=datetostr(date);
  ZQRecebimento3.close;
  ZQRecebimento3.SQL.Clear;
  ZQRecebimento3.SQL.Add('DROP TABLE IF EXISTS cobranca');
  LStep := PerformanceStart;
  ZQRecebimento3.ExecSQL;
  PerformanceElapsed('Cobranca: DROP da tabela auxiliar', LStep);

  ZQVenda_cobr.Close;
  ZQVenda_cobr.SQL.Clear;
  ZQVenda_cobr.SQL.Add('Select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao from venda where idvenda='+quotedstr(ZQRecebimento2.FieldByName('venda_idvenda').Text));
  LStep := PerformanceStart;
  ZQVenda_cobr.open;
  PerformanceElapsed('Cobranca: venda para totais', LStep);

  ZQRecebimento3.close;
  ZQRecebimento3.SQL.Clear;
  // 27/08/2012 tony pediu pra mudar
//  ZQRecebimento3.SQL.Add('CREATE TABLE IF NOT EXISTS cobranca SELECT * FROM recebimento where Dt_Vencimento< :dt and saldo>''0'' and quadralote='+quotedstr(ZQRecebimentoquadralote.value));
  if empty(Label52.Caption) then
  begin
    ZQRecebimento3.SQL.Add('CREATE TABLE IF NOT EXISTS cobranca SELECT * FROM recebimento where Dt_Vencimento< :dt and saldo>''0'' and adversa='+quotedstr(ZQRecebimento.FieldByName('adversa').text));
  end
  else
  begin
    ZQRecebimento3.SQL.Add('CREATE TABLE IF NOT EXISTS cobranca SELECT * FROM recebimento where idloteamento='+quotedstr(Label52.Caption)+' and Dt_Vencimento< :dt and saldo>''0'' and adversa='+quotedstr(ZQRecebimento.FieldByName('adversa').text));
  end;
  ZQRecebimento3.ParamByName('dt').AsDate:=date;
  LStep := PerformanceStart;
  ZQRecebimento3.ExecSQL;
  PerformanceElapsed('Cobranca: CREATE e copia dos titulos', LStep);


  ZQRecebimento4.close;
  ZQRecebimento4.SQL.Clear;
  // 27/08/2012 tony pediu pra mudar
//  ZQRecebimento4.SQL.Add('Select * from Cobranca where Dt_Vencimento< :dt and saldo>''0'' and quadralote='+quotedstr(ZQRecebimentoquadralote.value)+' order by nomeadversa,DT_Vencimento');
  if empty(Label52.Caption) then
  begin
    ZQRecebimento4.SQL.Add('Select * from Cobranca where Dt_Vencimento< :dt and saldo>''0'' and adversa='+quotedstr(ZQRecebimento.FieldByName('adversa').text)+' order by nomeadversa,DT_Vencimento');
  end
  else
  begin
    ZQRecebimento4.SQL.Add('Select * from Cobranca where idloteamento='+quotedstr(Label52.Caption)+' and Dt_Vencimento< :dt and saldo>''0'' and adversa='+quotedstr(ZQRecebimento.FieldByName('adversa').text)+' order by nomeadversa,DT_Vencimento');
  end;
  ZQRecebimento4.ParamByName('dt').AsDate:=date;
  { Mantem ZQRecebimento4 aberto para seus consumidores existentes, mas nao
    percorre seus campos lookup somente para calcular tres totais. }
  LStep := PerformanceStart;
  ZQRecebimento4.Open;
  PerformanceElapsed('Cobranca: abrir titulos dos totais', LStep);
  LStep := PerformanceStart;
  AbrirTitulosTotaisCobranca;
  PerformanceElapsed('Cobranca: abrir titulos para soma sem lookups', LStep);

  Bar1.Max := FTitulosTotaisCobranca.RecordCount;
  Bar1.Position := 0;
  LStep := PerformanceStart;
  FTitulosTotaisCobranca.DisableControls;
  try
    FTitulosTotaisCobranca.First;
    while not FTitulosTotaisCobranca.Eof do
    begin
      Bar1.Position := FTitulosTotaisCobranca.RecNo;
      PrepararTaxasVendaCobranca(
        FTitulosTotaisCobranca.FieldByName('venda_idvenda').AsLargeInt);
      LDias := StrToInt(DifDias(
        FTitulosTotaisCobranca.FieldByName('Dt_Vencimento').AsDateTime, Date));
      CalcularValoresCobranca(
        FTitulosTotaisCobranca.FieldByName('saldo').AsFloat,
        FTaxasTotaisCobranca.FieldByName('Multa').AsFloat,
        FTaxasTotaisCobranca.FieldByName('Mora').AsFloat,
        FTaxasTotaisCobranca.FieldByName('tabela_price').AsFloat, LDias,
        Lmora.Caption <> 'N', LMulta, LMoraCalculada, LCorrigido);
      XMora.Value := XMora.Value + LMoraCalculada;
      XMulta.Value := XMulta.Value + LMulta;
      XCorrigido.Value := XCorrigido.Value + LCorrigido;
      FTitulosTotaisCobranca.Next;
    end;
  finally
    FTitulosTotaisCobranca.EnableControls;
  end;
  PerformanceElapsed('Cobranca: percorrer e somar totais', LStep);
  Bar1.Position := 0;
  ZQRecebimento4.First;    LStep := PerformanceStart;
    DBcobranca.Refresh;
    DBcobranca.Repaint;
    PerformanceElapsed('Cobranca: atualizar desenho do grid', LStep);
    fpg.clear;
    XN.Value:=0;
    CDnegocio.close;
    PerformanceElapsed('Cobranca: totais e recriacao da tabela', LInicio);
    Application.ProcessMessages;
  except
    { Uma falha nao pode deixar totais parciais disponiveis para negociacao. }
    nparcelas.Value  := 0;
    vencidos.Value   := 0;
    Xmulta.Value     := 0;
    XMora.Value      := 0;
    XCorrigido.Value := 0;
    Bar1.Position    := 0;
    fpg.Clear;
    XN.Value         := 0;
    CDnegocio.Close;
    CDnegocio2.Close;
    raise;
  end;
end;

procedure TFrmCad_Recebimento.DBGRecebTitleClick(Column: TColumn);
Var
  campo, Vartipo, varsql :string;
  varcodigo : integer;
begin
  vartipo := '';
  if CBPag.Checked then
    vartipo := vartipo+quotedstr('P')+',';
  if CBRec.Checked then
    vartipo := vartipo+quotedstr('R')+',';
  vartipo := copy(vartipo,1,length(vartipo)-1);
  if empty(vartipo) Then Begin
    vartipo := quotedstr('R')+','+quotedstr('P');
  end;
  campo:=column.fieldname;
  if  (not empty(campo)) and (pos(campo,'ordem'+'adversanome'+'idrecebimento'+'TipDoc'+'documento'+'Dt_Entrada'+'Dt_Vencimento'+'saldo'+'nome_loteamento'+'Apelido')>0) Then
  Begin
    DBGReceb.Columns[0].Color := clWindow;
    DBGReceb.Columns[1].Color := clWindow;
    DBGReceb.Columns[3].Color := clWindow;
    DBGReceb.Columns[4].Color := clWindow;
    DBGReceb.Columns[5].Color := clWindow;
    DBGReceb.Columns[6].Color := clWindow;
    DBGReceb.Columns[7].Color := clWindow;
    DBGReceb.Columns[9].Color := clWindow;
    application.processmessages;
    if campo = 'nome_loteamento' Then
      campo := 'idloteamento';
    if campo = 'Apelido' Then
      campo := 'usuario';
    if campo = 'adversanome' then
      campo := 'adversa';
    varcodigo := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
    varsql := DM_tabelas.ZQRecebimento.sql.Text;
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.sql.clear;
    DM_tabelas.ZQRecebimento.SQL.Add('Select * from recebimento  r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) ');
    DM_tabelas.ZQRecebimento.SQL.Add(' and r.recpag in ('+vartipo+')');
    if (not CBAberto.Checked) and (CBFinal.Checked) Then
      DM_tabelas.ZQRecebimento.SQL.Add(' and (r.saldo<=0 or r.saldo is null)');
    if (not CBFinal.Checked) and (CBAberto.Checked) Then
      DM_tabelas.ZQRecebimento.SQL.Add(' and r.saldo>0');
    DM_tabelas.ZQRecebimento.SQL.Add(' order by r.'+campo+', r.dt_vencimento');
    if (pos(' order by r.'+campo,varsql)>0) and ( pos(' desc ',varsql)=0) Then
      DM_tabelas.ZQRecebimento.SQL.Add(' desc ');
    DM_tabelas.ZQRecebimento.Open;
    DM_tabelas.ZQRecebimento.Locate('r.idRecebimento',varcodigo,[]);
    column.color:=$00EDF0F1;
    botoes_setas;
    if Pag_Receb.PageIndex = 1 Then
       atualiza_telas;
  end;
  if Pag_Receb.PageIndex = 0 Then
  begin
    ECliente.Text := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
    EAdversa.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
    EContabil.Text:= DM_Tabelas.ZQRecebimento.FieldByName('contabil').Text;
  end;

end;

procedure TFrmCad_Recebimento.DBGRecebKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_setas;
  AtualizarLookupsRecebimento;
  IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    GBCheque.Visible := True
  else
    GBCheque.Visible := False;
  if Pag_Receb.PageIndex = 1 Then
     atualiza_telas;
  if Pag_Receb.PageIndex = 0 Then
  begin
    ECliente.Text := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
    EAdversa.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
    EContabil.Text:= DM_Tabelas.ZQRecebimento.FieldByName('contabil').Text;
  end;
end;

procedure TFrmCad_Recebimento.DBGRecebMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes_setas;
  AtualizarLookupsRecebimento;
    IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    GBCheque.Visible := True
  else
    GBCheque.Visible := False;
  if Pag_Receb.PageIndex = 1 Then
     atualiza_telas;

  if Pag_Receb.PageIndex = 0 Then
  begin
    ECliente.Text := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
    EAdversa.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
    EContabil.Text:= DM_Tabelas.ZQRecebimento.FieldByName('contabil').Text;
  end;

end;

procedure TFrmCad_Recebimento.EClienteExit(Sender: TObject);
begin
  If DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit] Then Begin
    if empty(ECliente.text) then
    begin
      if not AchaCLiente(450, 140,Ecliente.Text) Then Begin
        Ecliente.SetFocus;
        Exit;
      End;
      DM_tabelas.ZQRecebimento.FieldByName('cliente').AsInteger := DM_tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;
      Ecliente.Text := DM_tabelas.ZQParticipante.FieldByName('nome_parte').AsString;
    end;
  end;
end;

procedure TFrmCad_Recebimento.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmCad_Recebimento.DBMObosExit(Sender: TObject);
begin
  if DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit] Then Begin
    if pos(DM_tabelas.ZQRecebimento.FieldByName('TipDoc').AsString,('CH/CP'))>0 Then Begin
      GBCheque.Visible := True;
      DBEBanco.SetFocus;
      exit;
    end;
    DXBGravar.SetFocus;
  end;
end;

procedure TFrmCad_Recebimento.DXBBaixaIncluirClick(Sender: TObject);
begin
  AtivaBaixa_Campos;
  DM_tabelas.ZQReceb_Baixa.Insert;
end;

procedure TFrmCad_Recebimento.DXBBaixaEditarClick(Sender: TObject);
begin
  AtivaBaixa_Campos;
  DM_tabelas.ZQReceb_Baixa.Edit;
end;

procedure TFrmCad_Recebimento.DXBBaixaCancelarClick(Sender: TObject);
begin
  DM_tabelas.ZQReceb_Baixa.Cancel;
  Desativa_Campos;
end;

procedure TFrmCad_Recebimento.DesativaBaixa_Campos;
Begin
//
End;

procedure TFrmCad_Recebimento.ativaBaixa_Campos;
Begin
//
End;

procedure TFrmCad_Recebimento.Calc_Saldo;
Var
  VarSaldo : Double;
Begin
  VarSaldo := 0;
  DM_tabelas.ZQReceb_Baixa.First;
  DM_tabelas.ZQReceb_Baixa.DisableControls;
  while not DM_tabelas.ZQReceb_Baixa.Eof do begin
    VarSaldo := VarSaldo+DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').AsFloat+DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').AsFloat-DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').AsFloat;
    DM_tabelas.ZQReceb_Baixa.Next;
  end;
  DM_tabelas.ZQReceb_Baixa.EnableControls;
  DM_tabelas.ZQRecebimento.Edit;
  DM_tabelas.ZQRecebimento.FieldByName('saldo').AsFloat := DM_tabelas.ZQRecebimento.FieldByName('Valor').AsFloat - VarSaldo;
  DM_tabelas.ZQRecebimento.Post;
  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  if DM_TAbelas.ZQRecebimento.active=true then
     DM_tabelas.ZQRecebimento.Refresh;
  Panel1.Visible:=false;
  Panel2.Visible:=false;  
  Application.ProcessMessages;

End;


procedure TFrmCad_Recebimento.DXBBaixaExcluirClick(Sender: TObject);
begin
  if DM_tabelas.ZQReceb_Baixa.RecordCount>0 Then Begin
    DM_tabelas.ZQReceb_Baixa.Delete;
    Calc_Saldo;
  end;
end;

procedure TFrmCad_Recebimento.DXBBaixarClick(Sender: TObject);
Var
  Varregidrec : integer;
begin
  if not Verif_senha('Financeiro','Baixar ','Cliente: '+DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString) then exit;
  Varregidrec := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
  FrmRecebBaixa.Label9.Caption := 'Baixando Títulos em Aberto do Contas a Pagar';
  FrmRecebBaixa.Label9.Color := $0000BB00;
  if DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString = 'R' Then Begin
    FrmRecebBaixa.Label9.Caption := 'Baixando Títulos em Aberto do Contas a Receber';
    FrmRecebBaixa.Label9.Color := ClRed;
  end;
  AbrirModal(Self, FrmRecebBaixa);
  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  if DM_TAbelas.ZQRecebimento.active=true then
     DM_tabelas.ZQRecebimento.Refresh;
  if DM_TAbelas.ZQRecebBxTemp.active=true then
     DM_TAbelas.zqrecebbxTemp.Refresh;
  if ZQGerou.active=true then
     ZQGerou.Refresh;
  DM_tabelas.zqRecebimento.Locate('IdRecebimento',Varregidrec,[]);
  DBGReceb.SetFocus;
  Panel1.Visible:=false;
  Panel2.Visible:=false;  
  Application.ProcessMessages;
  
//  Botoes_setas;
end;

procedure TFrmCad_Recebimento.DBGrid1Enter(Sender: TObject);
begin
  botoes_setas;
end;

procedure TFrmCad_Recebimento.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_setas;
end;

procedure TFrmCad_Recebimento.DBGrid1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes_setas;
end;

procedure TFrmCad_Recebimento.DXBExBaixaClick(Sender: TObject);
Var
  Varrefbaixa, varregistro, varidrec : integer;
  varsqltexto : string;
  VarRestoreSaldo : Double;
  VarVinculo : string;
  VarRefAnterior : Integer;

  procedure ExecutarSQLDireto(const ASQL: string);
  begin
    with TZQuery.Create(nil) do
    try
      Connection := DM_tabelas.zconeccao;
      SQL.Text := ASQL;
      ExecSQL;
    finally
      Free;
    end;
  end;

  function BuscarRefAnterior(const AIdRecib, ARefAtual: Integer): Integer;
  begin
    Result := 0;
    with TZQuery.Create(nil) do
    try
      Connection := DM_tabelas.zconeccao;
      SQL.Text := 'select max(refer) as ref_anterior from recbxhist ' +
                  'where idrecib = :idrecib and refer <> :refer';
      Params.ParamByName('idrecib').AsInteger := AIdRecib;
      Params.ParamByName('refer').AsInteger := ARefAtual;
      Open;
      if not Fields[0].IsNull then
        Result := Fields[0].AsInteger;
    finally
      Free;
    end;
  end;

  procedure RestaurarRecebimento(const AIdRecib: Integer; const ARestoreSaldo: Double);
  begin
    with TZQuery.Create(nil) do
    try
      Connection := DM_tabelas.zconeccao;
      SQL.Text := 'select idrecebimento, valor, saldo, refbaixa, tip from recebimento ' +
                  'where idrecebimento = :idrecebimento';
      Params.ParamByName('idrecebimento').AsInteger := AIdRecib;
      Open;
      if not IsEmpty then
      begin
        Edit;
        FieldByName('saldo').AsFloat :=
          FieldByName('saldo').AsFloat + ARestoreSaldo;

        if FieldByName('saldo').AsFloat > FieldByName('valor').AsFloat then
          FieldByName('saldo').AsFloat := FieldByName('valor').AsFloat;

        VarRefAnterior := BuscarRefAnterior(AIdRecib, Varrefbaixa);
        if VarRefAnterior > 0 then
          FieldByName('refbaixa').AsInteger := VarRefAnterior
        else
          FieldByName('refbaixa').Clear;

        if FieldByName('saldo').AsFloat > 0 then
          FieldByName('tip').AsString := 'P'
        else if FieldByName('tip').AsString = 'P' then
          FieldByName('tip').Clear;
        Post;
      end;
    finally
      Free;
    end;
  end;

begin
  if not Verif_senha('Financeiro','Excluir Baixa','Cliente: '+DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString+' | Nº baixa: '+DM_tabelas.ZQRecebimento.FieldByName('RefBaixa').Text+' - Lanç.'+DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text) then exit;
  Varrefbaixa:=0;
  varidrec:=0;
  if FrmObs_estorno=nil then
     FrmObs_estorno:=TFrmObs_estorno.Create(Self);
  if DM_tabelas.ZQRecebimento.FieldByName('RefBaixa').AsInteger>0 Then
  Begin
    if (DM_tabelas.ZQRecebimento.FieldByName('TipDoc').AsString='CH') then
    begin
      FrmObs_estorno.obs_est.Text:=DM_tabelas.ZQRecebimento.FieldByName('observ_estorno').AsString;
      AbrirModal(Self, FrmObs_estorno);
    end;
//    if DM_tabelas.ZQRecebBxTemprefer.Value=0 then
//    begin
//       Varrefbaixa := DM_tabelas.ZQRecebimentorefbaixa.Value;
//       Atualiza_Telas;
//    end
//    else
    Varidrec := DM_tabelas.ZQRecebBxTemp.FieldByName('idrecebimento').AsLargeInt;
    Varrefbaixa := DM_Tabelas.ZQRecebBxTemp.FieldByName('refer').AsInteger;
    VarVinculo := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString + IntToStr(Varrefbaixa);

    DM_TAbelas.zqrecebbxTemp.Close;

    DM_tabelas.ZQReBxHi.Close;    
    DM_tabelas.ZQReBxHi.SQL.Clear;
    DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist where refer='+IntToStr(Varrefbaixa)+' order by idrecib,idrecbxhist');
    DM_tabelas.ZQReBxHi.Open;
    DM_tabelas.ZQReBxHi.First;
    varregistro := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
    varsqltexto := DM_tabelas.ZQRecebimento.SQL.Text;

    DM_tabelas.zconeccao.StartTransaction;
    try
      ExecutarSQLDireto('delete from caixa where vinculo = ' + QuotedStr(VarVinculo));
      ExecutarSQLDireto('delete from mov_banco where vinculo = ' + QuotedStr(VarVinculo));
      ExecutarSQLDireto('delete from cheque_rec where idvinculo = ' + IntToStr(Varrefbaixa));
      ExecutarSQLDireto('delete from recebimento where RefVinda = ' + IntToStr(Varrefbaixa));
      ExecutarSQLDireto('delete from receb_baixa where refbaixa = ' + IntToStr(Varrefbaixa));

      DM_tabelas.ZQRecebimento.Close;
      DM_tabelas.ZQRecebimento.SQL.Clear;
      DM_tabelas.ZQRecebimento.SQL.Add(varsqltexto);
      DM_tabelas.ZQRecebimento.Open;

      DM_tabelas.ZQReBxHi.First;
      DM_tabelas.ZQReBxHi.DisableControls;
      while not DM_tabelas.ZQReBxHi.Eof do
      begin
        if DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').AsFloat > 0 then
          VarRestoreSaldo := DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').AsFloat
        else
          VarRestoreSaldo := DM_tabelas.ZQReBxHi.FieldByName('valor').AsFloat -
                             DM_tabelas.ZQReBxHi.FieldByName('juros').AsFloat +
                             DM_tabelas.ZQReBxHi.FieldByName('descon').AsFloat;

        if VarRestoreSaldo < 0 then
          VarRestoreSaldo := 0;

        RestaurarRecebimento(DM_tabelas.ZQReBxHi.FieldByName('idrecib').AsInteger, VarRestoreSaldo);

        DM_tabelas.ZQReBxHi.Next;
      end;
      DM_tabelas.ZQReBxHi.EnableControls;

      ExecutarSQLDireto('delete from RecBxHist where refer = ' + IntToStr(Varrefbaixa));
      DM_tabelas.zconeccao.Commit;
    except
      on E: Exception do
      begin
        if DM_tabelas.ZQReBxHi.ControlsDisabled then
          DM_tabelas.ZQReBxHi.EnableControls;
        DM_tabelas.zconeccao.Rollback;
        mensagem('Erro ao excluir baixa: ' + E.Message);
        Exit;
      end;
    end;

    ZQReceb_Baixa.Close;
    ZQReceb_Baixa.SQL.clear;
    ZQReceb_Baixa.SQL.Add(' Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela From receb_baixa where 1=0');
    ZQReceb_Baixa.open;

    DM_tabelas.ZQCaixa.Close;
    DM_tabelas.ZQCaixa.SQL.Clear;
    DM_tabelas.ZQCaixa.SQL.Add('select codigo,plano_contas_codigo,data_lan,vr_lan,obs,documento,empresa,vinculo,saldo,credeb,contraquem,contracodigo,Substituicao,sq from caixa');
    DM_tabelas.ZQCaixa.Open;

    DM_Tabelas.ZQMovBancaria.close;
    DM_Tabelas.ZQMovBancaria.SQL.Clear;
    DM_Tabelas.ZQMovBancaria.SQL.Add('select lancamento,conta_bancaria_cod_banco,plano_contas_codigo,documento,dt_lanc,hora_lanc,dt_conciliado,vr_lanc,credeb,saldo_lanc,saldo_conci,obs,contraquem,');
    DM_Tabelas.ZQMovBancaria.SQL.Add('       contracodigo,vinculo,Substituicao,sq ');
    DM_Tabelas.ZQMovBancaria.SQL.Add(' from mov_banco');
    DM_Tabelas.ZQMovBancaria.Open;

    DM_tabelas.ZQReBxHi.Filtered:=false;
    DM_tabelas.ZQReBxHi.close;
    DM_tabelas.ZQReBxHi.SQL.Clear;
    DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist');
    DM_tabelas.ZQReBxHi.Open;
    DM_tabelas.ZQRecebimento.Locate('IdRecebimento',varregistro,[]);



    DM_tabelas.ZQRecebimento.Close;

    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add(varsqltexto);
    DM_tabelas.ZQRecebimento.Open;

    Panel1.Visible:=true;
    Panel2.Visible:=true;
    Application.ProcessMessages;
    if ZQVincRec.active=true then
       ZQVIncRec.Refresh;
    if ZQGerou.active=true then
       ZQGerou.Refresh;
    DM_TAbelas.zqrecebbxTemp.Open;
    DM_TAbelas.zqrecebbxTemp.Refresh;
    Panel1.Visible:=false;
    Panel2.Visible:=false;
    Application.ProcessMessages;
    atualiza_telas;

    if (DM_tabelas.ZQRecebimento.FieldByName('TipDoc').AsString='CH') then
    begin
      DM_tabelas.ZQRecebimento.edit;
      DM_tabelas.ZQRecebimento.FieldByName('observ_estorno').AsString:= FrmObs_estorno.obs_est.Text;
      DM_tabelas.ZQRecebimento.post;
    end;
    FrmObs_estorno:=nil;
    FrmObs_estorno.Free;
  end;
end;

procedure TFrmCad_Recebimento.WDDBProrrogadoExit(Sender: TObject);
begin
  DXBGravar.SetFocus;
end;

procedure TFrmCad_Recebimento.DBGrid2CellClick(Column: TColumn);
begin
  DM_tabelas.ZQRecebimento.Locate('IdRecebimento',ZQVIncRec.FieldByName('idRecebimento').AsLargeInt,[]);
  botoes_setas;
end;

procedure TFrmCad_Recebimento.DBGrid1CellClick(Column: TColumn);
begin
  DM_tabelas.ZQRecebimento.Locate('refvinda',DM_TAbelas.ZQRecebBxTemp.FieldByName('idreceb_baixa').AsLargeInt,[]);
  botoes_setas;
end;

procedure TFrmCad_Recebimento.EContabilExit(Sender: TObject);
begin
  if (not ((DXBCancelar.Focused) or (DXBFechar.Focused))) and (DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit]) and not(empty(Econtabil.text)) Then Begin
    if not AchaPlanoDeConta(300, 20, 'S', EContabil.Text, '') then Begin
      EContabil.SetFocus;
      exit;
    end;
    EContabil.Text := DM_tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' - '+DM_tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString+'  -  '+masccontabil(DM_tabelas.ZQAchaPlanoDeContas.FieldByName('mascara').AsString);
    DM_tabelas.ZQRecebimento.FieldByName('contabil').AsInteger := DM_tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
  end;
end;

procedure TFrmCad_Recebimento.DBGRecebDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DM_tabelas.ZQRecebimento.FieldByName('saldo').AsFloat <=0 Then
    DBGReceb.Canvas.Brush.Color:= $00FFEEDD; // coloque aqui a cor desejada
  if (DM_tabelas.ZQRecebimento.FieldByName('saldo').AsFloat >0) and (DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime<date) Then
    DBGReceb.Canvas.Brush.Color:= $00DDDDFF; // coloque aqui a cor desejada
  if DBEIDReceb.Text = DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text Then Begin
    DBGReceb.Canvas.Brush.Color := $006CFFFF;
    DBGReceb.Canvas.Font.Color := $00A80000;
//    DBGReceb.Canvas.Font.Style := [FsBold];
  end;
  DBGReceb.DefaultDrawDataCell(Rect, DBGReceb.columns[datacol].field, State);
end;

procedure TFrmCad_Recebimento.DBGrid3CellClick(Column: TColumn);
begin
  DM_tabelas.ZQRecebimento.Locate('IdRecebimento',ZQOrigem.FieldByName('idRecebimento').AsLargeInt,[]);
  botoes_setas;
end;

procedure TFrmCad_Recebimento.DBGrid3MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
//  DM_tabelas.ZQRecebimento.Locate('IdRecebimento',ZQOrigemIdRecebimento.Value,[]);
//  botoes_setas;

end;

procedure TFrmCad_Recebimento.DBGrid4MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
//  DM_tabelas.ZQRecebimento.Locate('IdRecebimento',DM_TAbelas.zqrecebbxTempidRecebimento.Value,[]);
//  botoes_setas;
end;

procedure TFrmCad_Recebimento.EAdversaExit(Sender: TObject);
begin
  if (not ((DXBCancelar.Focused) or (DXBFechar.Focused))) and (DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit]) Then Begin
    AchaCliente(400, 300, Eadversa.Text);
    if not (DM_tabelas.ZQRecebimento.State in [DsInsert, DsEdit]) Then
      DM_tabelas.ZQRecebimento.Edit;
    Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
    DM_Tabelas.ZQRecebimento.FieldByName('adversa').AsInteger := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
  end;
end;

procedure TFrmCad_Recebimento.DBGRecebEnter(Sender: TObject);
begin
  botoes_setas;
  AtualizarLookupsRecebimento;
  IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    GBCheque.Visible := True
  else
    GBCheque.Visible := False;
  if Pag_Receb.PageIndex = 1 Then 
     atualiza_telas;
  if Pag_Receb.PageIndex = 0 Then
  begin
    ECliente.Text := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
    EAdversa.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
    EContabil.Text:= DM_Tabelas.ZQRecebimento.FieldByName('contabil').Text;
  end;
end;

procedure TFrmCad_Recebimento.TimeRecebimentoTimer(Sender: TObject);
begin
//  DM_Tabelas.ZQRecebimento.Refresh;
end;

procedure TFrmCad_Recebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQRecebimento_bancario.Close;
  TimeRecebimento.Enabled := False;
  DM_tabelas.ZQTipodoc.close;
  DM_tabelas.ZQPlanoDeContas.close;
  DM_tabelas.ZQReBxHi.close;
  DM_TAbelas.ZQRecebBxTemp.close;
  DM_Tabelas.ZQCheque.close;
  DM_tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQContaBancaria.close;
  ZQVenda_cobr.Close;
  Panel1.Visible:=false;
  Panel2.Visible:=false;
  FrmCad_Recebimento:=nil;
  Action := CaFree;
  Application.ProcessMessages;
end;

procedure TFrmCad_Recebimento.RGDEleteExit(Sender: TObject);
var
  varsql, varmensa : string;
  Varord, varre : integer;
  vartem : boolean;
begin
  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;
  if DM_tabelas.ZQRecebimento.RecordCount<1 Then
    Exit;
  varmensa := '';
  if DM_Tabelas.ZQRecebBxTemp.RecordCount>0 Then
    varmensa := 'Não vai excluir porque existe baixa deste lanç.';
  if not Verif_senha('Financeiro','Excluir ','Documento: '+DM_tabelas.ZQRecebimento.FieldByName('documento').AsString+'  '+DM_tabelas.ZQRecebimento.FieldByName('ordem').AsString+'  Interno: '+DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Text+' '+varmensa) then Begin
    RGDElete.Visible := False;
    exit;
  end;
  if DM_Tabelas.ZQRecebBxTemp.RecordCount>0 Then Begin
    mensagem('Este lançamento só pode ser excluido depois que todas as baixas dele forem excluidas!!!');
    RGDelete.Visible := False;
    exit;
  end;
  if RGDElete.ItemIndex=0 Then Begin
    if (DM_Tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').AsString,[])) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S') then Begin
      if DM_Tabelas.ZQCheque.Locate('idrecebimento',DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]) Then
        DM_Tabelas.ZQCheque.Delete;
    end;
    DM_tabelas.ZQRecebimento.Delete
  end
  else Begin
    varsql := DM_Tabelas.ZQRecebimento.SQL.Text;
    varre := DM_Tabelas.ZQRecebimento.RecNo;
    Varord := DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger;
    DM_TAbelas.ZQRecebimento.close;
    DM_TAbelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add('select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
    DM_tabelas.ZQRecebimento.SQL.Add('     from recebimento r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.numordem='+inttostr(varord));
    DM_Tabelas.ZQRecebimento.Open;
    vartem := True;
    DM_Tabelas.ZQRecebimento.DisableControls;
    while not DM_Tabelas.ZQRecebimento.Eof do Begin
      if DM_Tabelas.ZQRecebimento.FieldByName('saldo').AsFloat = DM_Tabelas.ZQRecebimento.FieldByName('Valor').AsFloat then Begin
         if (DM_Tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').AsString,[])) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S') then Begin
          if DM_Tabelas.ZQCheque.Locate('idrecebimento',DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]) Then
            DM_Tabelas.ZQCheque.Delete;
        end;
        DM_Tabelas.ZQRecebimento.Delete;
      end
      else Begin
        vartem := False;
        DM_Tabelas.ZQRecebimento.next;
      end;
    end;
    DM_Tabelas.ZQRecebimento.EnableControls;
    if vartem Then Begin
      ZQTempDivid.Close;
      ZQTempDivid.SQL.Clear;
      ZQTempDivid.SQL.Add('select iddividido,participante,percentual,planodeconta,tipodocumento,ordem from dividido where ordem='+inttostr(varord));
      ZQTempDivid.Open;
      ZQTempDivid.DisableControls;
      while not ZQTempDivid.Eof do Begin
        ZQTempDivid.Delete;
//        ZQTempDivid.Next;
      end;
      ZQTempDivid.EnableControls;
    end;
    DM_Tabelas.ZQRecebimento.Close;
    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_TAbelas.ZQRecebimento.SQL.Add(varsql);
    DM_Tabelas.ZQRecebimento.Open;
    DM_Tabelas.ZQRecebimento.RecNo := varre;
  end;
  if ZQGerou.active=true then
     ZQGerou.Refresh;
  RGDelete.Visible := False;
  if DM_TAbelas.ZQRecebimento.active=true then
     DM_Tabelas.ZQRecebimento.Refresh;
  Panel1.Visible:=false;
  Panel2.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrmCad_Recebimento.RGDEleteClick(Sender: TObject);
begin
  DXBExcluir.SetFocus;
end;

procedure TFrmCad_Recebimento.DXBExcluirClick(Sender: TObject);
begin
  if simnao('Confirma a exclusão do registro ?','SIM') then
  begin
    RGDElete.Visible := True;
    RGDElete.Items.Strings[0] := 'Registro '+DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text;
    RGDElete.Items.Strings[1] := 'Ordem '+DM_Tabelas.ZQRecebimento.FieldByName('numordem').Text;
    RGDElete.SetFocus;
  end;
end;


procedure TFrmCad_Recebimento.CBRecClick(Sender: TObject);
begin
  Filtrar;
end;

Procedure TFrmCad_Recebimento.Filtrar;
var
  Vartipo : string;
begin
  vartipo := '';
  if CBPag.Checked then
    vartipo := vartipo+quotedstr('P')+',';
  if CBRec.Checked then
    vartipo := vartipo+quotedstr('R')+',';
  vartipo := copy(vartipo,1,length(vartipo)-1);
  if empty(vartipo) Then Begin
    vartipo := quotedstr('R')+','+quotedstr('P');
  end;
  DM_tabelas.ZQRecebimento.Close;  
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.recpag in ('+vartipo+')');
  if (not CBAberto.Checked) and (CBFinal.Checked) Then
    DM_tabelas.ZQRecebimento.SQL.Add(' and (r.saldo<=0 or r.saldo is null)');
  if (not CBFinal.Checked) and (CBAberto.Checked) Then
    DM_tabelas.ZQRecebimento.SQL.Add(' and r.saldo>0');
  DM_tabelas.ZQRecebimento.SQL.Add(' order by r.DT_Vencimento');
  DM_tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQRecebimento.Last;
  DBGReceb.SetFocus;
end;

procedure TFrmCad_Recebimento.CBAbertoClick(Sender: TObject);
begin
  Filtrar;
end;

procedure TFrmCad_Recebimento.CBFinalClick(Sender: TObject);
begin
  Filtrar;
end;

procedure TFrmCad_Recebimento.CBPagClick(Sender: TObject);
begin
  Filtrar;
end;

procedure TFrmCad_Recebimento.DXClick(Sender: TObject);
var
  LReferencia: string;
begin
  LReferencia := '';
  if DM_Tabelas.ZQReBxHi.Active and (not DM_Tabelas.ZQReBxHi.IsEmpty) then
    LReferencia := DM_Tabelas.ZQReBxHi.FieldByName('refer').AsString;
  if Empty(LReferencia) and DM_Tabelas.ZQRecebBxTemp.Active and
     (not DM_Tabelas.ZQRecebBxTemp.IsEmpty) then
    LReferencia := DM_Tabelas.ZQRecebBxTemp.FieldByName('refer').AsString;
  if Empty(LReferencia) then
    LReferencia := DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').AsString;
  if Empty(LReferencia) then
    LReferencia := DM_Tabelas.ZQRecebimento.FieldByName('refvinda').AsString;
  if Empty(LReferencia) then
  begin
    mensagem('Nao foi encontrada uma baixa para reimprimir o recibo deste titulo.');
    Exit;
  end;

  if FrmImpRecibo=nil then
     FrmImpRecibo:=TFrmImpRecibo.Create(Self);
  FrmImpRecibo.N_Baixa.Text := LReferencia;
  FrmImpRecibo.Nomecli.Text := DM_Tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  if DM_Tabelas.ZQRecebBxTemp.Active and (not DM_Tabelas.ZQRecebBxTemp.IsEmpty) then
    FrmImpRecibo.Memoobs.Text := DM_Tabelas.ZQRecebBxTemp.FieldByName('obsebx').AsString
  else
    FrmImpRecibo.Memoobs.Text := DM_Tabelas.ZQRecebimento.FieldByName('Observ').AsString;
  AbrirModal(Self, FrmImpRecibo);
  FreeAndNil(FrmImpRecibo);
  FrmImpRecibo:=nil;
end;

procedure TFrmCad_Recebimento.EParticipanteExit(Sender: TObject);
begin
  if not empty(EParticipante.Text) Then Begin
    if AchaCLiente(450, 140,EParticipante.Text) Then Begin
      EParticipante.Text := DM_tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
      Ecodcli.Text := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
    end
    else
      Ecodcli.Text := '';
  end
  else
    Ecodcli.Text := '';
  Atualiza_Telas;
end;

procedure TFrmCad_Recebimento.EdevedorExit(Sender: TObject);
begin
  if not empty(Edevedor.Text) Then Begin
    if AchaCliente(400, 300, Edevedor.Text) then Begin
      Edevedor.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
      Ecoddeved.Text := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
    end
    else
     Ecoddeved.Text := '';
  end
  else
    Ecoddeved.Text := '';
end;

procedure TFrmCad_Recebimento.dxButton4Click(Sender: TObject);
var
  varinc : integer;
  vardias : Double;
begin
  Gauge1.Visible := True;
  CDSResult.Close;
  CDSResult.CreateDataSet;
  XDVar.DateValue := XDEIni.DateValue;
  varinc := 0;
  If copy(XDEfim.Text,4,7)=copy(XDEIni.Text,4,7) Then Begin
    vardias := (XDEfim.DateValue - XDEIni.DateValue);
    Gauge1.MaxValue := strtoint(floattostr(vardias));
    while XDVar.DateValue <= XDEfim.DateValue do begin
      CDSResult.Insert;
      CDSResultdata.Value := XDVar.DateValue;
      CDSResultordem.Value := copy(XDVar.Text,1,2); // copy(XDVar.Text,7,4)+'/'+copy(XDVar.Text,4,2)+'/'+
      CDSResult.Post;
      inc(varinc);
      Gauge1.Progress := varinc;
      XDVar.DateValue := XDVar.DateValue +1;
    end;
  end
  else Begin
    while copy(XDVar.Text,7,4)+copy(XDVar.Text,4,2)<=copy(XDEfim.Text,7,4)+copy(XDEfim.Text,4,2) do begin
      CDSResult.Insert;
      CDSResultordem.Value := copy(XDVar.Text,7,4)+'/'+copy(XDVar.Text,4,2);
      if  copy(XDVar.Text,4,7)=copy(XDEfim.Text,4,7) then
        CDSResultdata.Value :=XDEfim.DateValue
      else
        CDSResultdata.Value := strtodate('01/'+copy(datetostr(IncMonth(XDVar.DateValue,1)),4,2)+'/'+copy(XDVar.Text,7,4))-1;
      CDSResult.Post;
      XDVar.DateValue := IncMonth(XDVar.DateValue,1);
    end;
  end;
  CDSResult.IndexFieldNames := 'ordem';
  CDSResult.First;
  for varinc := 1 to CDSResult.RecordCount do begin
    Gauge1.Progress := varinc;
    CDSResult.RecNo := varinc;
    ZQResult1.Close;
    ZQResult1.SQL.Clear;
    ZQResult1.SQL.Add('select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQResult1.SQL.Add('       custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQResult1.SQL.Add('       somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    ZQResult1.SQL.Add('       dt_nao_pagou_no_mes,descricao_juridico,');
    ZQResult1.SQL.Add('       sum(valor) as soma from recebimento');
    ZQResult1.SQL.Add(' where dt_vencimento between :dt1 and :dt2');
    if not empty(Ecodcli.Text) Then
      ZQResult1.SQL.Add(' and cliente='+quotedstr(Ecodcli.Text));
    if not empty(Ecoddeved.Text) Then
      ZQResult1.SQL.Add(' and adversa='+quotedstr(Ecoddeved.Text));
    ZQResult1.ParamByName('dt1').AsDate:=XDEIni.DateValue;
    ZQResult1.ParamByName('dt2').AsDate:=CDSResultdata.Value;
    ZQResult1.SQL.Add(' group by recpag');
    ZQResult1.Open;
    CDSResult.Edit;
    if ZQResult1.Locate('recpag','R',[]) then
      CDSResultareceber.Value := ZQResult1.FieldByName('soma').AsFloat;
    if ZQResult1.Locate('recpag','P',[]) then
      CDSResultapagar.Value := ZQResult1.FieldByName('soma').AsFloat;
    ZQFeito.close;
    ZQFeito.SQL.Clear;
    ZQFeito.SQL.Add('SELECT *,sum(vr_rec) as somafeito FROM (select dt_rec,vr_rec,recpag from receb_baixa join recbxhist on refbaixa=refer');
    ZQFeito.SQL.Add(' join recebimento on idrecib=idrecebimento');
    ZQFeito.SQL.Add(' where dt_rec between :dt1 and :dt2');
    if not empty(Ecodcli.Text) Then
      ZQFeito.SQL.Add(' and cliente='+QuotedStr(Ecodcli.Text));
    if not empty(Ecoddeved.Text) Then
      ZQFeito.SQL.Add(' and adversa='+quotedstr(Ecoddeved.Text));
    ZQFeito.SQL.Add(' group by idreceb_baixa) as  A');
    ZQFeito.SQL.Add(' group by recpag');
    ZQFeito.ParamByName('dt1').AsDate:=XDEIni.DateValue;
    ZQFeito.ParamByName('dt2').AsDate:=CDSResultdata.Value;
    ZQFeito.Open;
    if ZQFeito.Locate('recpag','R',[]) then
      CDSResultrecebido.Value := ZQFeito.FieldByName('somafeito').AsFloat;
    if ZQFeito.Locate('recpag','P',[]) then
      CDSResultpago.Value := ZQFeito.FieldByName('somafeito').AsFloat;
  end;
  CDSResult.Post;
  Gauge1.Visible := false;
  AtualizarGrafico(FPaintBoxGrafico1);
end;

procedure TFrmCad_Recebimento.CriarGraficosRuntime;
  procedure CriarPaintBox(Container: TPanel; var PaintBox: TPaintBox;
    PaintHandler: TNotifyEvent);
  begin
    if PaintBox <> nil then
      Exit;
    if Container = nil then
      Exit;
    PaintBox := TPaintBox.Create(Container);
    PaintBox.Parent := Container;
    PaintBox.Align := alClient;
    PaintBox.OnPaint := PaintHandler;
  end;
begin
  CriarPaintBox(PnlGrafico1, FPaintBoxGrafico1, PintarGrafico1);
  CriarPaintBox(PnlGrafico2, FPaintBoxGrafico2, PintarGrafico2);
  CriarPaintBox(PnlGrafico3, FPaintBoxGrafico3, PintarGrafico3);
end;

procedure TFrmCad_Recebimento.PintarGrafico1(Sender: TObject);
var
  Series: TChartSeriesArray;
  Options: TChartRenderOptions;
  PaintBox: TPaintBox;
begin
  if FPintandoGrafico1 or not (Sender is TPaintBox) then
    Exit;
  FPintandoGrafico1 := True;
  try
    PaintBox := TPaintBox(Sender);
    LoadChartSeriesFromDataSet(CDSResult,['ordem', 'ordem', 'ordem', 'ordem'],['A Pagar', 'A Receber', 'Recebido', 'Pago'],
                               ['apagar', 'areceber', 'recebido', 'pago'],[clRed, clTeal, clBlue, TColor(234)], Series);
    Series[0].LineStyle := psDot;
    Series[0].LineWidth := 4;
    Series[1].LineStyle := psDash;
    Series[1].LineWidth := 2;
    Series[2].LineWidth := 2;
    Series[3].LineWidth := 2;

    Options := DefaultChartRenderOptions(clWhite);
    Options.UseBackgroundGradient := True;
    Options.BackgroundStartColor := TColor(RGB(228, 239, 255));
    Options.BackgroundEndColor := clWhite;
    Options.BorderColor := clBlue;
    Options.BorderWidth := 2;
    GenerateStyledMultiSeriesChart(PaintBox.Canvas, Series, ctLine,PaintBox.ClientWidth, PaintBox.ClientHeight,
                                   'Recebimentos e Pagamentos do Periodo Informado', Options, True, True);
  finally
    FPintandoGrafico1 := False;
  end;
end;

procedure TFrmCad_Recebimento.PintarGrafico2(Sender: TObject);
var
  Series: TChartSeriesArray;
  ChartTitle: string;
  Options: TChartRenderOptions;
  PaintBox: TPaintBox;
begin
  if FPintandoGrafico2 or not (Sender is TPaintBox) then
    Exit;
  FPintandoGrafico2 := True;
  try
    PaintBox := TPaintBox(Sender);
    if Trim(EdNomeLoteamento.Text) = '' then
      ChartTitle := 'Lot.: Todos - Previsao de Recebimentos no Vencimento do Periodo de ' + XDEIni1.Text + ' a ' + XDEFim1.Text
    else
      ChartTitle := 'Lot.:' + EdNomeLoteamento.Text + ' - Previsao de Recebimentos no Vencimento do Periodo de ' + XDEIni1.Text + ' a ' + XDEFim1.Text;
    LoadChartSeriesFromDataSet(CDSResult,['pareceber', 'precebido', 'patrazado'],['A Receber no Periodo', 'Recebido no Periodo', 'Aberto no Periodo'],
                              ['areceber', 'recebido', 'apagar'],[clLime, clBlue, clRed], Series);
    Series[0].BarPenWidth := 2;
    Series[0].MarkStyle := cmsValue;
    Series[1].BarPenWidth := 2;
    Series[1].MarkStyle := cmsValue;
    Series[2].BarPenWidth := 2;
    Series[2].MarkStyle := cmsValue;

    Options := DefaultChartRenderOptions(clWhite);
    Options.UseBackgroundGradient := True;
    Options.BackgroundStartColor := TColor(RGB(228, 239, 255));
    Options.BackgroundEndColor := clWhite;
    Options.BorderColor := clBlue;
    Options.BorderWidth := 2;
    GenerateStyledMultiSeriesChart(PaintBox.Canvas, Series, ctBar,PaintBox.ClientWidth, PaintBox.ClientHeight,ChartTitle, Options, True, True);
  finally
    FPintandoGrafico2 := False;
  end;
end;

procedure TFrmCad_Recebimento.PintarGrafico3(Sender: TObject);
var
  Series: TChartSeriesArray;
  Options: TChartRenderOptions;
begin
  if FPaintBoxGrafico3 = nil then
    Exit;
  SetLength(Series, 1);
  Series[0].Title := 'Inadimplencia';
  Series[0].Color := clRed;
  SetLength(Series[0].Data, 0);
  if ZQInadimplentes.Active and (ZQInadimplentes.FindField('nome_loteamento_sql') <> nil) and (ZQInadimplentes.FindField('inadimplente') <> nil) then
     LoadChartSeriesFromDataSet(ZQInadimplentes, ['nome_loteamento_sql'],['Inadimplencia'], ['inadimplente'], [clRed], Series);
  if Length(Series) > 0 then
  begin
    Series[0].ColorEachPoint := True;
    Series[0].MarkStyle := cmsValue;
  end;

  Options := DefaultChartRenderOptions(clWhite);
  Options.BorderColor := clBlue;
  Options.BorderWidth := 2;
  Options.LegendBySeries := False;
  GenerateStyledMultiSeriesChart(FPaintBoxGrafico3.Canvas, Series, ctHorizontalBar, FPaintBoxGrafico3.ClientWidth,
                                 FPaintBoxGrafico3.ClientHeight,'Grafico de Inadimplencias ate a data de ' + DateToStr(Date - 1), Options, True, True);
end;

procedure TFrmCad_Recebimento.AtualizarGrafico(PaintBox: TPaintBox);
begin
  if (PaintBox <> nil) and (PaintBox.Parent <> nil) and
     (PaintBox.Width > 0) and (PaintBox.Height > 0) then
    PaintBox.Invalidate;
end;

procedure TFrmCad_Recebimento.dxButton5Click(Sender: TObject);
begin
  if FPaintBoxGrafico1 <> nil then
    PrintPaintBoxChart(FPaintBoxGrafico1, 'Recebimentos e Pagamentos do Periodo Informado', True);
end;

procedure TFrmCad_Recebimento.Pag_RecebClick(Sender: TObject);
begin
  // 28/10/2010
  Lmora.Caption:='';
  if Pag_Receb.PageIndex = 1 Then
  begin
     Atualiza_Telas;
  end;
  if Pag_Receb.PageIndex = 3 Then
  begin
    if simnao('Deseja utilizar a Tabela Price Para Mora/Dia ?','SIM') then
       Lmora.Caption:='S'
    else
       Lmora.Caption:='N';
    ZQRecebimento.Close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Add(' Select R.idrecebimento,R.documento,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.empresa, R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.nomeadversa,R.somar,');
    ZQRecebimento.SQL.Add('        R.juros,R.descontos,R.Data_Quitacao,R.tip,R.dt_nao_pagou_no_mes,R.Multa,R.Mora,R.Cliente, V.idvenda,V.datavenda,V.imovel,V.valorvenda,V.forma_reajuste,V.marca,V.Multa,V.Mora,V.tabela_price');
    ZQRecebimento.SQL.Add('  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.quadralote='+quotedstr(DM_tabelas.ZQRecebimento.FieldByName('quadralote').AsString)+' and R.recpag ='+quotedstr('R')+' and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento');

    //ZQRecebimento.SQL.Add('Select *  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.quadralote='+quotedstr(DM_tabelas.ZQRecebimentoquadralote.Value)+' and R.recpag ='+quotedstr('R')+' and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento');
    ZQRecebimento.ParamByName('dt').AsDate:=date;
    ZQRecebimento.open;

    ZQVenda_cobr.Close;
    ZQVenda_cobr.SQL.Clear;
    ZQVenda_cobr.SQL.Add('Select idvenda,multa,mora from venda where idvenda='+quotedstr(ZQRecebimento.FieldByName('venda_idvenda').Text));
    ZQVenda_cobr.open;

    RGSelecao.Visible:=true;
  //  Atualiza_DBcobranca;
    DBcobranca.SetFocus;
  end
  else
  begin
    ZQRecebimento.Close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento2.Close;
    ZQRecebimento2.SQL.Clear;
  end;
  if Pag_Receb.PageIndex = 2 Then
  begin
    ZQInadimplentes.Close;
    ZQInadimplentes.SQL.Clear;
//    ZQInadimplentes.SQL.Add('Select *, sum(valor) as inadimplente from Recebimento where TipDoc ='+quotedstr('BO')+' and recpag ='+quotedstr('R')+' and Dt_Vencimento between :dt3 and :dt4 and saldo>0 group by idloteamento order by inadimplente desc');

    ZQInadimplentes.SQL.Add('Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,VrDoc,ordem,TipDoc,saldo,marcar,empresa,');
    ZQInadimplentes.SQL.Add(' origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,nomeadversa,somar,juros,descontos,Data_Quitacao,');
    ZQInadimplentes.SQL.Add(' (select l.apelido from loteamento l where l.idloteamento=recebimento.idloteamento) as nome_loteamento_sql,');
    ZQInadimplentes.SQL.Add(' sum(valor) as inadimplente from Recebimento where TipDoc ='+quotedstr('BO')+' and recpag ='+quotedstr('R')+' and Dt_Vencimento between :dt3 and :dt4 and saldo>0 group by idloteamento order by inadimplente desc');

    ZQInadimplentes.ParamByName('dt3').AsDate:=strtodate('01/01/1900');
    ZQInadimplentes.ParamByName('dt4').AsDate:=(date-1);
    ZQInadimplentes.open;
  end
  else
  begin
    ZQInadimplentes.Close;
    ZQInadimplentes.SQL.Clear;
  end;
end;

procedure TFrmCad_Recebimento.DBXBaixaClick(Sender: TObject);

begin
  JBBaixa.Visible := True;
  JBBaixa.SetFocus;
end;

procedure TFrmCad_Recebimento.JvXPBar1Items1Click(Sender: TObject);
Var
  Varregidrec : integer;
begin
  if not Verif_senha('Financeiro','Baixar ','Cliente: '+DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString) then exit;
  try
    if FrmRecebBaixa=nil then
       FrmRecebBaixa := TFrmRecebBaixa.Create(self);
    FrmRecebBaixa.Left := Frm_principal.Left+7;
    FrmRecebBaixa.Top := Frm_principal.Top+135;
    Varregidrec := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
    FrmRecebBaixa.Label9.Caption := 'Baixando Títulos em Aberto do Contas a Pagar';
    FrmRecebBaixa.Label9.Color := $0000BB00;
    if DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString = 'R' Then Begin
      FrmRecebBaixa.Label9.Caption := 'Baixando Títulos em Aberto do Contas a Receber';
      FrmRecebBaixa.Label9.Color := ClRed;
    end;
    JBBaixa.Visible:=false;

    AbrirModal(Self, FrmRecebBaixa);
  finally
    Panel1.Visible:=true;
    Panel2.Visible:=true;
    Application.ProcessMessages;

  //  DM_tabelas.ZQRecebimento.Refresh;
    if DM_TAbelas.ZQRecebBxTemp.active=true then
       DM_TAbelas.zqrecebbxTemp.Refresh;
    if ZQGerou.active=true then
       ZQGerou.Refresh;
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Add('select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_tabelas.ZQRecebimento.SQL.Add('       custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_tabelas.ZQRecebimento.SQL.Add('       somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_tabelas.ZQRecebimento.SQL.Add('       dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento  r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.IdRecebimento='+quotedstr(Varregidrec.ToString)+' order by r.DT_Vencimento ');
    DM_tabelas.ZQRecebimento.open;
  //  DM_tabelas.zqRecebimento.Refresh;
    DM_tabelas.zqRecebimento.Locate('IdRecebimento',Varregidrec,[]);
    DBGReceb.SetFocus;
    Panel1.Visible:=false;
    Panel2.Visible:=false;
    Application.ProcessMessages;
   // Botoes_setas;
    FrmRecebBaixa:=nil;
    FrmRecebBaixa.Free;

  end;

end;

procedure TFrmCad_Recebimento.JBBaixaExit(Sender: TObject);
begin
  JBBaixa.Visible := False;

end;

procedure TFrmCad_Recebimento.JvXPBar1Items0Click(Sender: TObject);
begin
 if DM_Tabelas.ZQConfiguracoes.FieldByName('Tipo_reajuste').AsString='DATA QUITAÇÃO GERAL' then
 begin
   try
     if Frm_ReajusteDeParcelas2=nil then
        Frm_ReajusteDeParcelas2:=TFrm_ReajusteDeParcelas2.Create(Self);
     Frm_ReajusteDeParcelas2.Top := FrmCad_Recebimento.Top + 110;
     AbrirModal(Self, Frm_ReajusteDeParcelas2);
   finally
     FreeAndNil(Frm_ReajusteDeParcelas2);
     JBProcesso.Visible:=false;
   end;

 end
 else                                                      
 begin
   try
     if Frm_ReajusteDeParcelas=nil then
        Frm_ReajusteDeParcelas:=TFrm_ReajusteDeParcelas.Create(Self);
     Frm_ReajusteDeParcelas.Top := FrmCad_Recebimento.Top + 110;
     AbrirModal(Self, Frm_ReajusteDeParcelas);
   finally
     FreeAndNil(Frm_ReajusteDeParcelas);
     JBProcesso.Visible:=false;
   end;


 end;
end;

procedure TFrmCad_Recebimento.DBXProcessoClick(Sender: TObject);
begin
  JBProcesso.Visible := True;
  JBProcesso.SetFocus;
end;

procedure TFrmCad_Recebimento.JBProcessoExit(Sender: TObject);
begin
  JBProcesso.Visible := False;
end;

procedure TFrmCad_Recebimento.JBBaixaItems0Click(Sender: TObject);
begin
  JBBaixa.Visible:=false;
  try
    if Frm_BaixaAutomatica=nil then
       Frm_BaixaAutomatica := TFrm_BaixaAutomatica.Create(self);
    Frm_BaixaAutomatica.Left := Frm_principal.Left+7;
    Frm_BaixaAutomatica.Top := Frm_principal.Top+135;
    AbrirModal(Self, Frm_BaixaAutomatica);
  finally
    Frm_BaixaAutomatica:=nil;
    Frm_BaixaAutomatica.Free;
  end;

//  Frm_BaixaAutomatica.showmodal;
end;

procedure TFrmCad_Recebimento.JBProcessoItems0Click(Sender: TObject);
begin
  if FrmAditamento=nil then
     FrmAditamento:=TFrmAditamento.Create(Self);
  FrmAditamento.Left := FrmCad_Recebimento.Left+7;
//  FrmAditamento.Top := FrmCad_Recebimento.Top+135;
  if not Verif_senha('Aditamento','Tela Principal','') then exit;
  FrmAditamento.show;
  JBProcesso.Visible:=false;
end;

procedure TFrmCad_Recebimento.JBProcessoItems2Click(Sender: TObject);
begin

  if Frmquitacao=nil then
     Frmquitacao:=TFrmquitacao.Create(Self);

  Frmquitacao.qlotes.Text:=DM_tabelas.ZQRecebimento.FieldByName('quadralote').AsString;

  Frmquitacao.Left := FrmCad_Recebimento.Left+7;
 // Frmquitacao.Top := FrmCad_Recebimento.Top+135;
  if not Verif_senha('Quitação','Tela Principal','') then exit;
  Frmquitacao.show;
  JBProcesso.Visible:=false;
end;

procedure TFrmCad_Recebimento.JBProcessoItems10Click(Sender: TObject);
begin
  if FrmzerarNossoNumero=nil then
     FrmzerarNossoNumero:=TFrmzerarNossoNumero.Create(Self);
  if not Verif_senha('Zerar Nosso Numero','Tela Principal','') then exit;
  FrmzerarNossoNumero.show;
  JBProcesso.Visible:=false;
end;

procedure TFrmCad_Recebimento.JBProcessoItems1Click(Sender: TObject);
begin
  if Frm_Cessao=nil then
     Frm_Cessao:=TFrm_Cessao.Create(Self);
  Frm_Cessao.Left := FrmCad_Recebimento.Left+7;
  Frm_Cessao.Top := FrmCad_Recebimento.Top+135;
  if not Verif_senha('Cessao','Tela Principal','') then exit;
  if DM_tabelas.ZQQuadras.Active=false then
     DM_tabelas.ZQQuadras.open;
  if DM_Tabelas.ZQincorp_loteame.Active=false then
     DM_Tabelas.ZQincorp_loteame.Open;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  if DM_tabelas.ZQNumOrdem.Active=false then
     DM_tabelas.ZQNumOrdem.Open;
  if DM_Tabelas.ZQImovel.Active=false then
     DM_Tabelas.ZQImovel.open;
  if DM_Tabelas.ZQVendedor.Active=false then
     DM_Tabelas.ZQVendedor.open;
  if DM_Tabelas.ZQVenda.Active=false then   
     DM_Tabelas.ZQVenda.open;
  if DM_Tabelas.ZQComprador.Active=false then
     DM_Tabelas.ZQComprador.open;
  Frm_Cessao.show;
  JBProcesso.Visible:=false;
end;

procedure TFrmCad_Recebimento.XDBNumEdit1Enter(Sender: TObject);
begin
  if DM_tabelas.ZQRecebimento.State in [DsEdit, DsInsert] then
  begin
      if not Verif_senha('Financeiro','Vr.Saldo','') then exit;
      XDBNumEdit1.ReadOnly:=false;
  end;
end;

procedure TFrmCad_Recebimento.XDBNumEdit1Exit(Sender: TObject);
begin
  XDBNumEdit1.ReadOnly:=true;
end;

procedure TFrmCad_Recebimento.EdNomeLoteamentoEnter(Sender: TObject);
begin
  if empty(EdNomeLoteamento.Text) Then
  Begin
    if Acha_Loteamento(EdNomeLoteamento.Text) Then Begin
      EdNomeLoteamento.Text := DM_tabelas.ZQAchaLotea.FieldByName('apelido').AsString;
      Edloteamento.Text := DM_TAbelas.ZQAchaLotea.FieldByName('idloteamento').Text;
    end
    else
      Edloteamento.Text := '';
  end
  else
    Edloteamento.Text := '';
//  Atualiza_Telas;
end;

procedure TFrmCad_Recebimento.dxButton1Click(Sender: TObject);
var
  varinc : integer;
  vardias : Double;
begin
  Gauge2.Visible := True;
  CDSResult.Close;
  CDSResult.CreateDataSet;
  CDSResult.IndexFieldNames := 'ordem';
  
  XDVar1.DateValue := XDEIni1.DateValue;
  varinc := 0;
{  If copy(XDEfim1.Text,4,7)=copy(XDEIni1.Text,4,7) Then
  Begin
    vardias := (XDEfim1.DateValue - XDEIni1.DateValue);
    Gauge2.MaxValue := strtoint(floattostr(vardias));
    while XDVar1.DateValue <= XDEfim1.DateValue do begin
      CDSResult.Insert;
      CDSResultdata.Value := XDVar1.DateValue;
      CDSResultordem.Value := copy(XDVar1.Text,1,2); // copy(XDVar.Text,7,4)+'/'+copy(XDVar.Text,4,2)+'/'+
      CDSResult.Post;
      inc(varinc);
      Gauge2.Progress := varinc;
      XDVar1.DateValue := XDVar1.DateValue +1;
    end;
  end
  else
  Begin
    while copy(XDVar1.Text,7,4)+copy(XDVar1.Text,4,2)<=copy(XDEfim1.Text,7,4)+copy(XDEfim1.Text,4,2) do begin
      CDSResult.Insert;
      CDSResultordem.Value := copy(XDVar1.Text,7,4)+'/'+copy(XDVar1.Text,4,2);
      if  copy(XDVar1.Text,4,7)=copy(XDEfim1.Text,4,7) then
        CDSResultdata.Value :=XDEfim1.DateValue
      else
        CDSResultdata.Value := strtodate('01/'+copy(datetostr(IncMonth(XDVar1.DateValue,1)),4,2)+'/'+copy(XDVar1.Text,7,4))-1;
      CDSResult.Post;
      XDVar1.DateValue := IncMonth(XDVar1.DateValue,1);
    end;
  end;
  CDSResult.IndexFieldNames := 'ordem';
  CDSResult.First;
  for varinc := 1 to CDSResult.RecordCount do
  begin
    Gauge1.Progress := varinc;
    CDSResult.RecNo := varinc;}
    ZQResult1.Close;    
    ZQResult1.SQL.Clear;
    ZQResult1.SQL.Add('SELECT idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQResult1.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQResult1.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    ZQResult1.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico');
    ZQResult1.SQL.Add('        ,sum(valor) as soma FROM recebimento');
    ZQResult1.SQL.Add(' where dt_vencimento between :dt1 and :dt2');
    if not empty(Edloteamento.Text) Then
      ZQResult1.SQL.Add(' and idloteamento='+quotedstr(Edloteamento.Text));
    ZQResult1.ParamByName('dt1').AsDate:=XDEIni1.DateValue;
    ZQResult1.ParamByName('dt2').AsDate:=XDEfim1.DateValue; //CDSResultdata.Value;
    ZQResult1.SQL.Add(' group by recpag');
    ZQResult1.Open;
//    mensagem(transform(zqresult1soma.Value,'###,###,##0.00'));
//    CDSResult.Edit;
    CDSResult.insert;
    CDSResultordem.Value:='1';
    if ZQResult1.Locate('recpag','R',[]) then
      CDSResultareceber.Value := ZQResult1.FieldByName('soma').AsFloat;
    if ZQResult1.Locate('recpag','P',[]) then
      CDSResultapagar.Value := ZQResult1.FieldByName('soma').AsFloat;

    CDSResult.Post;

    ZQFeito.Close;
    ZQFeito.SQL.Clear;

    ZQFeito.SQL.Add('select *, sum(valor) as somafeito from recebimento');
    ZQFeito.SQL.Add(' where  dt_vencimento between :dt1 and :dt2');
    ZQFeito.SQL.Add(' and saldo=0 ');
    if not empty(Edloteamento.Text) Then
      ZQFeito.SQL.Add(' and idloteamento='+quotedstr(Edloteamento.Text));
    ZQFeito.ParamByName('dt1').AsDate:=XDEIni1.DateValue;
    ZQFeito.ParamByName('dt2').AsDate:=XDEfim1.DateValue;//CDSResultdata.Value;
    ZQFeito.SQL.Add(' group by recpag');
    ZQFeito.Open;
 //   mensagem(transform(zqfeitosomafeito.Value,'###,###,##0.00'));
    if ZQFeito.FieldByName('somafeito').AsFloat>0 then
    begin
      CDSResult.insert;
      CDSResultordem.Value:='2';
      if ZQFeito.Locate('recpag','R',[]) then
        CDSResultrecebido.Value := ZQFeito.FieldByName('somafeito').AsFloat;
      if ZQFeito.Locate('recpag','P',[]) then
        CDSResultpago.Value := ZQFeito.FieldByName('somafeito').AsFloat;
      CDSResult.post;
    end
    else
    begin
      CDSResult.insert;
      CDSResultordem.Value:='2';
      if ZQFeito.Locate('recpag','R',[]) then
        CDSResultrecebido.Value := 0;
      if ZQFeito.Locate('recpag','P',[]) then
        CDSResultpago.Value := 0;
      CDSResult.post;
    end;
  //end;
  // atrasados


  if ZQFeito.FieldByName('somafeito').AsFloat>0 then
  begin
    CDSResult.insert;
    CDSResultordem.Value:='3';
    CDSResultapagar.Value := ZQResult1.FieldByName('soma').AsFloat-ZQFeito.FieldByName('somafeito').AsFloat;
    CDSResult.Post;
  end
  else
  begin
    CDSResult.insert;
    CDSResultordem.Value:='3';
    CDSResultapagar.Value := 0;
    CDSResult.Post;
  end;
  CDSResult.First;
  CDSResult.edit;
  CDSResultpareceber.Value:=100;
  CDSResult.post;

  CDSResult.Next;
  if ZQFeito.FieldByName('somafeito').AsFloat>0 then
  begin
    CDSResult.Edit;
    CDSResultprecebido.Value:=(ZQFeito.FieldByName('somafeito').AsFloat*100)/ZQResult1.FieldByName('soma').AsFloat;
    CDSResult.post;
  end
  else
  begin
    CDSResult.Edit;
    CDSResultprecebido.Value:=0;
    CDSResult.post;
  end;
  CDSResult.Next;
  if ZQFeito.FieldByName('somafeito').AsFloat>0 then
  begin
    CDSResult.Edit;
    CDSResultpatrazado.Value:=((ZQResult1.FieldByName('soma').AsFloat-ZQFeito.FieldByName('somafeito').AsFloat)*100)/ZQResult1.FieldByName('soma').AsFloat;
    CDSResult.post;
  end
  else
  begin
    CDSResult.Edit;
    CDSResultpatrazado.Value:=0;
    CDSResult.post;
  end;
  CDSResult.First;


  Gauge2.Visible := false;
  AtualizarGrafico(FPaintBoxGrafico2);
end;

procedure TFrmCad_Recebimento.dxButton6Click(Sender: TObject);
var
  ChartTitle: string;
begin
  if FPaintBoxGrafico2 <> nil then
  begin
    if Trim(EdNomeLoteamento.Text) = '' then
      ChartTitle := 'Lot.: Todos - Previsao de Recebimentos no Vencimento do Periodo de ' +
        XDEIni1.Text + ' a ' + XDEFim1.Text
    else
      ChartTitle := 'Lot.:' + EdNomeLoteamento.Text +
        ' - Previsao de Recebimentos no Vencimento do Periodo de ' +
        XDEIni1.Text + ' a ' + XDEFim1.Text;
    PrintPaintBoxChart(FPaintBoxGrafico2, ChartTitle, True);
  end;
end;

procedure TFrmCad_Recebimento.tab_graficosClick(Sender: TObject);
begin
  if tab_graficos.PageIndex = 0 Then
  Begin
    ZQFeito.close;
    ZQResult1.Close;
    CDSResult.Close;
    Gauge1.Visible := False;
    XDEIni.SetFocus;
  end;
  if tab_graficos.PageIndex = 1 Then
  Begin
    ZQFeito.close;
    ZQResult1.Close;
    CDSResult.Close;
    Gauge2.Visible := False;
    XDEIni1.SetFocus;
  end;

  if tab_graficos.PageIndex = 2 then
    AtualizarGrafico(FPaintBoxGrafico3);

end;

procedure TFrmCad_Recebimento.ZQRecebimentoCalcFields(DataSet: TDataSet);
var
zmora:double;
LStep: UInt64;
begin
  { O evento de campos calculados só pode preencher campos calculados.
    Se algum código solicitar o cálculo fora do ciclo interno do dataset,
    não tente escrever no registro em modo dsBrowse. }
  if (DataSet = nil) or
     not (DataSet.State in [dsCalcFields, dsEdit, dsInsert]) then
    Exit;

  zmora:=0;

  // 03/09/2025
  GarantirVendaCobranca(ZQRecebimento.FieldByName('venda_idvenda').Text);
  LStep := PerformanceStart;

{  ZQRecebimentoDias.Value:= strtoint(DifDias(ZQRecebimentoDt_Vencimento.Value, date));
  zmora:=ExRound(ExRound((DM_Tabelas.ZQConfiguracoesMora.value/30),2)*ZQRecebimentoDias.Value,2);
  ZQRecebimentomora_dia.Value:=ExRound((ZQRecebimentosaldo.Value*zmora)/100,2);
  ZQRecebimentoMulta_Contrato.Value:=ExRound(ExRound((ZQRecebimentosaldo.Value*DM_Tabelas.ZQConfiguracoesMulta.value),2)/100,2);
  ZQRecebimentoparcela_corrigida.Value:=ZQRecebimentosaldo.Value+ZQRecebimentoMulta_Contrato.Value+ZQRecebimentomora_dia.Value;}

  // Tony pediu para mudar a formula no dia 28/08/2012


  ZQRecebimento.FieldByName('Dias').AsInteger:= strtoint(DifDias(ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime, date));
//  zmora:=ExRound(ExRound((DM_Tabelas.ZQConfiguracoesMora.value/30),2)*ZQRecebimentoDias.Value,2);
  if Lmora.Caption='N' then
     ZQRecebimento.FieldByName('mora_dia').AsFloat:=ExRound(((((ZQRecebimento.FieldByName('saldo').AsFloat/30)*ZQVenda_cobr.FieldByName('Mora').AsFloat)/100)*ZQRecebimento.FieldByName('Dias').AsInteger),2)
  else
     ZQRecebimento.FieldByName('mora_dia').AsFloat:=ExRound(((((ZQRecebimento.FieldByName('saldo').AsFloat/30)*ZQRecebimento.FieldByName('tabela_price').AsFloat)/100)*ZQRecebimento.FieldByName('Dias').AsInteger),2);
  ZQRecebimento.FieldByName('Multa_Contrato').AsFloat:=ExRound(ExRound((ZQRecebimento.FieldByName('saldo').AsFloat*ZQVenda_cobr.FieldByName('Multa').AsFloat),2)/100,2);
  ZQRecebimento.FieldByName('parcela_corrigida').AsFloat:=ZQRecebimento.FieldByName('saldo').AsFloat+ZQRecebimento.FieldByName('Multa_Contrato').AsFloat+ZQRecebimento.FieldByName('mora_dia').AsFloat;

  PerformanceSample(smCalculoFinanceiro, LStep);
  PreencherLookupsCobranca;

end;

procedure TFrmCad_Recebimento.DBcobrancaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
 // Atualiza_DBcobranca;
end;

procedure TFrmCad_Recebimento.DBcobrancaEnter(Sender: TObject);
begin
  if cbmontante.Checked=true then
  begin
    DBcobranca.Options:=DBcobranca.Options+[dgMultiSelect];
  end
  else
  begin
    DBcobranca.Options:=DBcobranca.Options-[dgMultiSelect];
  end;
  if FCobrancaAtualizada and (FContextoCobranca = ContextoCobranca) then
  begin
    PerformanceCheckpoint('Cobranca: retorno ao grid sem recarregar');
    FlushPerformanceLog;
    Exit;
  end;
  if not FCobrancaAtualizada then
    PerformanceElapsed('Cobranca: atualizar por dados invalidados', PerformanceStart)
  else
    PerformanceElapsed('Cobranca: atualizar por participante/filtro/modo/data alterado', PerformanceStart);
  Atualiza_DBcobranca;
  AtualizarCamposCobranca;
  FContextoCobranca := ContextoCobranca;
  FCobrancaAtualizada := FContextoCobranca <> '';
  DBcobranca.Invalidate;
  FlushPerformanceLog;
end;

procedure TFrmCad_Recebimento.fpgEnter(Sender: TObject);
begin
  fpg.clear;
  XN.value:=0;
  CDnegocio.IndexName:='';
  CDnegocio.Close;
end;

procedure TFrmCad_Recebimento.fpgExit(Sender: TObject);
var
formapg:string;
pag0,pag1,pag2,pag3,pag4,pag5,pag6,pag7,pag8,pag9:string;
pr2,parc,totpag,wDias:integer;
vr2,vrtotal,kmora,wmora,wMulta,wcorrigido:double;
vcto2:TDateTime;
begin
 if not empty(fpg.Text) then
 begin
  if cbmontante.Checked=false then
  begin
    wmora:=0;
    wMulta:=0;
    wcorrigido:=0;
    wDias:=0;
    kmora:=0;
    vrtotal:=0;
    CDnegocio.Close;
    CDnegocio.CreateDataSet;
    CDnegocio.open;
    formapg:='';
    formapg:=copy(Fpg.Text,1,3);
    if RGSelecao.ItemIndex=0 then
       formapg:=charrem(formapg);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,4,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,7,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,10,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,13,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,16,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,19,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,22,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,25,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,28,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=Xpagto.DateValue+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';

    pag0:= '';
    pag1:= '';
    pag2:= '';
    pag3:= '';
    pag4:= '';
    pag5:= '';
    pag6:= '';
    pag7:= '';
    pag8:= '';
    pag9:= '';
    totpag:= 0;
    pag0:=copy(Fpg.Text,1,3);
    if not empty(pag0) then
    begin
      totpag:=1;
    end;
    pag1:=copy(Fpg.Text,4,3);
    if not empty(pag1) then
    begin
      totpag:=2;
    end;
    pag2:=copy(Fpg.Text,7,3);
    if not empty(pag2) then
    begin
      totpag:=3;
    end;
    pag3:=copy(Fpg.Text,10,3);
    if not empty(pag3) then
    begin
      totpag:=4;
    end;
    pag4:=copy(Fpg.Text,13,3);
    if not empty(pag4) then
    begin
      totpag:=5;
    end;
    pag5:=copy(Fpg.Text,16,3);
    if not empty(pag5) then
    begin
      totpag:=6;
    end;
    pag6:=copy(Fpg.Text,19,3);
    if not empty(pag6) then
    begin
      totpag:=7;
    end;
    pag7:=copy(Fpg.Text,22,3);
    if not empty(pag7) then
    begin
      totpag:=8;
    end;
    pag8:=copy(Fpg.Text,25,3);
    if not empty(pag8) then
    begin
      totpag:=9;
    end;
    pag9:=copy(Fpg.Text,28,3);
    if not empty(pag9) then
    begin
      totpag:=10;
    end;
    if totpag<=0 then
    begin
       mensagem('Informe os Vencimentos.');
       fpg.SetFocus;
       exit;
    end;
    
    vrtotal:=ExRound(XCorrigido.Value/totpag,2);

    CDnegocio.first;
    Bar1.Position:=0;
    Bar1.Max:=CDnegocio.RecordCount;
    XN.value:=0;
    CDnegocio.DisableControls;
    while not CDnegocio.Eof do
    begin
      Bar1.Position:=CDnegocio.RecNo;
      wDias:=0;
      kmora:=0;
      wMora:=0;
      wMulta:=0;
      wcorrigido:=0;
      wDias:= strtoint(DifDias(Xpagto.DateValue,CDnegociovencimento.Value));
      if wDias>0 then
      begin
//        kmora:=ExRound(ExRound((DM_Tabelas.ZQConfiguracoesMora.value/30),2)*wDias,2);
        if Lmora.Caption='N' then
           kmora:=(ZQVenda_cobr.FieldByName('Mora').AsFloat/30)*wDias
        else
           kmora:=(ZQRecebimento.FieldByName('tabela_price').AsFloat/30)*wDias;

        wMora:=ExRound((vrtotal*kmora)/100,2);
//        wMulta:=ExRound(ExRound((vrtotal*DM_Tabelas.ZQConfiguracoesMulta.value),2)/100,2);
      end;
      wcorrigido:=vrtotal+wMulta+wMora;
      CDnegocio.edit;
      CDnegociovr_parcela.Value:=vrtotal;
      CDnegocioJuros.Value:=wMora+wMulta;
      CDnegociovr_parcela_corrigida.Value:=wcorrigido;
      CDnegocio.post;
      CDnegocio.Next;
    end;
    CDnegocio.EnableControls;
    Bar1.Position:=0;
    CDnegocio.IndexName:='Vencimentoidx';
    CDnegocio.first;
    Bar1.Position:=0;
    Bar1.Max:=CDnegocio.RecordCount;
    XN.value:=0;
    parc:=1;
    CDnegocio.DisableControls;
    while not CDnegocio.Eof do
    begin
      Bar1.Position:=CDnegocio.RecNo;
      CDnegocio.edit;
      CDnegocionparcela.Value:=parc;
      CDnegocio.post;
      XN.value:=XN.value+CDnegociovr_parcela_corrigida.Value;
      CDnegocio.Next;
      inc(parc);
    end;
    CDnegocio.EnableControls;
    Bar1.Position:=0;
  end
  else
  begin
    { A linha corrente nao equivale a uma parcela selecionada no grid. }
    if not ZQRecebimento.Active or ZQRecebimento.IsEmpty or
       (DBcobranca.SelectedRows.Count = 0) then
    begin
      CDnegocio.Close;
      CDnegocio2.Close;
      XN.Value := 0;
      mensagem('Selecione as parcelas no quadro de cobranca com Ctrl+clique antes de informar o vencimento.');
      Exit;
    end;
    GetSelectedRecord(ZQRecebimento, 'idrecebimento',DBcobranca.SelectedRows, ListBox1.Items);
    if CDnegocio2.IsEmpty then
    begin
      CDnegocio.Close;
      XN.Value := 0;
      mensagem('Nenhuma parcela disponivel para negociar. Selecione novamente as parcelas.');
      Exit;
    end;
    CDnegocio2.First;
    while not CDnegocio2.Eof do
    begin
      if CDnegocio2vencimento.IsNull or (CDnegocio2vencimento.AsDateTime <= 0) or
         CDnegocio2vr_parcela.IsNull or (CDnegocio2vr_parcela.AsFloat <= 0) then
      begin
        CDnegocio.Close;
        CDnegocio2.Close;
        XN.Value := 0;
        mensagem('Existe parcela selecionada sem vencimento ou saldo valido. Confira as parcelas antes de negociar.');
        Exit;
      end;
      CDnegocio2.Next;
    end;

    wmora:=0;
    wMulta:=0;
    wcorrigido:=0;
    wDias:=0;
    kmora:=0;
    vrtotal:=0;
    CDnegocio.Close;
    CDnegocio.CreateDataSet;
    CDnegocio.open;
    CDnegocio2.IndexName:='idxdias';
    CDnegocio2.First;

    wDias:=CDnegocio2nparcela.Value;

    formapg:='';
    formapg:=copy(Fpg.Text,1,3);
    if RGSelecao.ItemIndex=0 then
       formapg:=charrem(formapg);
    if not empty(formapg) then
    begin
      CDnegocio.Insert;
      CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
      CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
      CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,4,3);
    if not empty(formapg) then
    begin
      CDnegocio.Insert;
      CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
      CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
      CDnegocio.post;
    end;
     formapg:='';
     formapg:=copy(Fpg.Text,7,3);
    if not empty(formapg) then
    begin
       CDnegocio.Insert;
       CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
       CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
       CDnegocio.post;
    end;
    formapg:='';
    formapg:=copy(Fpg.Text,10,3);
    if not empty(formapg) then
    begin
        CDnegocio.Insert;
        CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
        CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
        CDnegocio.post;
     end;
     formapg:='';
     formapg:=copy(Fpg.Text,13,3);
     if not empty(formapg) then
     begin
         CDnegocio.Insert;
         CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
         CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
         CDnegocio.post;
     end;
     formapg:='';
     formapg:=copy(Fpg.Text,16,3);
     if not empty(formapg) then
     begin
         CDnegocio.Insert;
         CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
         CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
         CDnegocio.post;
     end;
     formapg:='';
     formapg:=copy(Fpg.Text,19,3);
     if not empty(formapg) then
     begin
         CDnegocio.Insert;
         CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
         CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
         CDnegocio.post;
     end;
     formapg:='';
     formapg:=copy(Fpg.Text,22,3);
     if not empty(formapg) then
     begin
         CDnegocio.Insert;
         CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
         CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
         CDnegocio.post;
     end;
     formapg:='';
     formapg:=copy(Fpg.Text,25,3);
     if not empty(formapg) then
     begin
         CDnegocio.Insert;
         CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
         CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
         CDnegocio.post;
     end;
     formapg:='';
     formapg:=copy(Fpg.Text,28,3);
     if not empty(formapg) then
     begin
         CDnegocio.Insert;
         CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
         CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
         CDnegocio.post;
     end;
     formapg:='';
     vrtotal:=vrtotal+CDnegocio2vr_parcela.Value;
     CDnegocio2.Next;
    CDnegocio2.DisableControls;
    while not CDnegocio2.Eof do
    begin
      if wDias=CDnegocio2nparcela.Value then
      begin
        formapg:='';
        vrtotal:=vrtotal+CDnegocio2vr_parcela.Value;
        CDnegocio2.Next;
      end
      else
      begin
        formapg:='';
        formapg:=copy(Fpg.Text,1,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,4,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,7,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,10,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,13,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,16,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,19,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,22,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,25,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
        formapg:=copy(Fpg.Text,28,3);
        if not empty(formapg) then
        begin
           CDnegocio.Insert;
           CDnegociovencimento.Value:=CDnegocio2vencimento.Value+CDnegocio2nparcela.value+strtoint(formapg);
           CDnegociodias.Value:=CDnegocio2nparcela.value+strtoint(formapg);
           CDnegocio.post;
        end;
        formapg:='';
//        vrtotal:=vrtotal+CDnegocio2vr_parcela.Value;
//        CDnegocio2.Next;
        wDias:=CDnegocio2nparcela.Value;
      end;
    end;
    CDnegocio2.EnableControls;
    CDnegocio2.close;
    CDnegocio.First;

    pag0:= '';
    pag1:= '';
    pag2:= '';
    pag3:= '';
    pag4:= '';
    pag5:= '';
    pag6:= '';
    pag7:= '';
    pag8:= '';
    pag9:= '';
    totpag:= 0;
    pag0:=copy(Fpg.Text,1,3);
    if not empty(pag0) then
    begin
      totpag:=1;
    end;
    pag1:=copy(Fpg.Text,4,3);
    if not empty(pag1) then
    begin
      totpag:=2;
    end;
    pag2:=copy(Fpg.Text,7,3);
    if not empty(pag2) then
    begin
      totpag:=3;
    end;
    pag3:=copy(Fpg.Text,10,3);
    if not empty(pag3) then
    begin
      totpag:=4;
    end;
    pag4:=copy(Fpg.Text,13,3);
    if not empty(pag4) then
    begin
      totpag:=5;
    end;
    pag5:=copy(Fpg.Text,16,3);
    if not empty(pag5) then
    begin
      totpag:=6;
    end;
    pag6:=copy(Fpg.Text,19,3);
    if not empty(pag6) then
    begin
      totpag:=7;
    end;
    pag7:=copy(Fpg.Text,22,3);
    if not empty(pag7) then
    begin
      totpag:=8;
    end;
    pag8:=copy(Fpg.Text,25,3);
    if not empty(pag8) then
    begin
      totpag:=9;
    end;
    pag9:=copy(Fpg.Text,28,3);
    if not empty(pag9) then
    begin
      totpag:=10;
    end;
    if totpag<=0 then
    begin
       mensagem('Informe os Vencimentos.');
       fpg.SetFocus;
       exit;
    end;

    totpag:=CDnegocio.RecordCount;

    vrtotal:=ExRound(vrtotal/totpag,2);

    CDnegocio.first;
    Bar1.Position:=0;
    Bar1.Max:=CDnegocio.RecordCount;
    XN.value:=0;
    parc:=1;
    CDnegocio.DisableControls;
    while not CDnegocio.Eof do
    begin
      Bar1.Position:=CDnegocio.RecNo;
      wDias:=0;
      kmora:=0;
      wMora:=0;
      wMulta:=0;
      wcorrigido:=0;
//      if CDnegociovencimento.Value>Xpagto.DateValue then
//         wDias:= strtoint(DifDias(Xpagto.DateValue,CDnegociovencimento.Value))
  //    else
//         wDias:= strtoint(DifDias(CDnegociovencimento.Value,Xpagto.DateValue));
      wDias:=CDnegociodias.Value;

      if wDias>0 then
      begin
        if Lmora.Caption='N' then
           kmora:=(ZQVenda_cobr.FieldByName('Mora').AsFloat/30)*wDias
        else
           kmora:=(ZQRecebimento.FieldByName('tabela_price').AsFloat/30)*wDias;
//      kmora:=ExRound(ExRound((DM_Tabelas.ZQConfiguracoesMora.value/30),2)*wDias,2);
        wMora:=ExRound((vrtotal*kmora)/100,2);
        wMulta:=ExRound(ExRound((vrtotal*ZQVenda_cobr.FieldByName('Multa').AsFloat),2)/100,2);
      end;
      wcorrigido:=vrtotal+wMulta+wMora;
      CDnegocio.edit;
      CDnegociovr_parcela.Value:=vrtotal;
      CDnegocioJuros.Value:=wMora+wMulta;
      CDnegociovr_parcela_corrigida.Value:=wcorrigido;
      CDnegocio.post;

      XN.value:=XN.value+wcorrigido;
      CDnegocio.Next;
    end;
    CDnegocio.EnableControls;
    Bar1.Position:=0;
    CDnegocio.IndexName:='Vencimentoidx';
    CDnegocio.First;

    //---------------//

    kmora:=0;
    wMora:=0;
    wMulta:=0;
    wcorrigido:=0;
    CDnegocio2.close;
    CDnegocio2.CreateDataSet;
    CDnegocio2.IndexName:='';
    CDnegocio2.Open;

    wDias:=0;
    vcto2:=CDnegociovencimento.Value;
    Bar1.Position:=0;
    Bar1.Max:=CDnegocio.RecordCount;
    XN.value:=0;
    CDnegocio.DisableControls;
    while not CDnegocio.Eof do
    begin
      Bar1.Position:=CDnegocio.RecNo;
      if vcto2=CDnegociovencimento.Value then
      begin
        wMora:=wMora+CDnegocioJuros.Value;
        wMulta:=wMulta+CDnegociovr_parcela.Value;
        wcorrigido:=wcorrigido+CDnegociovr_parcela_corrigida.Value;
        XN.value:=XN.value+wcorrigido;
        CDnegocio.Next;
      end
      else
      begin
        CDnegocio2.Insert;
        CDnegocio2vencimento.Value:=vcto2;
        CDnegocio2vr_parcela.Value:=wMulta;
        CDnegocio2Juros.Value:=wMora;
        CDnegocio2vr_parcela_corrigida.Value:=wcorrigido;
        CDnegocio2.post;
        wDias:=0;
        kmora:=0;
        wMora:=0;
        wMulta:=0;
        wcorrigido:=0;
        vcto2:=CDnegociovencimento.Value;
      end;
    end;
    CDnegocio.EnableControls;
    CDnegocio2.Insert;
    CDnegocio2vencimento.Value:=vcto2;
    CDnegocio2vr_parcela.Value:=wMulta;
    CDnegocio2Juros.Value:=wMora;
    CDnegocio2vr_parcela_corrigida.Value:=wcorrigido;
    CDnegocio2.post;
    wDias:=0;
    kmora:=0;
    wMora:=0;
    wMulta:=0;
    wcorrigido:=0;
    vcto2:=CDnegociovencimento.Value;


    CDnegocio.IndexName:='';
    CDnegocio.close;
    CDnegocio.CreateDataSet;
    CDnegocio.open;
    CDnegocio2.IndexName:='vencimentoidx';
    CDnegocio2.First;
    parc:=1;
    Bar1.Position:=0;
    Bar1.Max:=CDnegocio2.RecordCount;
    XN.value:=0;
    CDnegocio2.DisableControls;
    while not CDnegocio2.Eof do
    begin
      Bar1.Position:=CDnegocio2.RecNo;
      CDnegocio.Insert;
      CDnegocionparcela.Value:=parc;
      CDnegociovencimento.Value:=CDnegocio2vencimento.Value;
      CDnegociovr_parcela.Value:=CDnegocio2vr_parcela.Value;
      CDnegocioJuros.Value:=CDnegocio2Juros.Value;
      CDnegociovr_parcela_corrigida.Value:=CDnegocio2vr_parcela_corrigida.Value;
      CDnegocio.post;
      XN.value:=XN.value+CDnegocio2vr_parcela_corrigida.Value;
      CDnegocio2.Next;
      inc(parc);
    end;
    CDnegocio2.EnableControls;
    CDnegocio2.close;
    CDnegocio.IndexName:='Vencimentoidx';
    CDnegocio.First;
    Bar1.Position:=0;
  end;
 end;
end;

procedure TFrmCad_Recebimento.dxButton8Click(Sender: TObject);
var
posi:integer;
LStep, LOperacao: UInt64;
begin
  FrmPesqCobranca.Top := FrmCad_Recebimento.Top+100;
  FrmPesqCobranca.Left := FrmCad_Recebimento.Left;
  AbrirModal(Self, FrmPesqCobranca);
  LOperacao := PerformanceStart;
  PerformanceCheckpoint('Cobranca: retorno da selecao na pesquisa');
  FlushPerformanceLog;
  ZQRecebimento.close;
  ZQRecebimento.SQL.Clear;
  PrepararLookupsCobranca;
  // 27/08/2012 tony pediu pra mudar
//  ZQRecebimento.SQL.Add('Select *  from Recebimento where quadralote='+quotedstr(FrmPesqCobranca.ZQTempCliRecebquadralote.Value)+' and adversa='+quotedstr(FrmPesqCobranca.ZQTempCliRecebadversa.Text)+' and Dt_Vencimento< :dt and saldo>0 order by nomeadversa,DT_Vencimento ');
  if FrmPesqCobranca.RGFiltro.ItemIndex=0 then
  begin

    if FrmPesqCobranca.ednome.Text='' then
    begin
       ZQRecebimento.SQL.Add('Select R.idrecebimento,R.documento,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.empresa, R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.nomeadversa,R.somar,');
       ZQRecebimento.SQL.Add('       R.juros,R.descontos,R.Data_Quitacao,R.tip,R.dt_nao_pagou_no_mes,R.Multa,R.Mora,R.Cliente, V.idvenda,V.datavenda,V.imovel,V.valorvenda,V.forma_reajuste,V.marca,V.Multa,V.Mora,V.tabela_price ');
       ZQRecebimento.SQL.Add('  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.adversa='+quotedstr(FrmPesqCobranca.ZQTempCliReceb.FieldByName('adversa').Text)+' and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento ');
       Label52.Caption:='';
    end
    else
    begin
       posi := pos('|',FrmPesqCobranca.ednome.Text);
       dec(posi);
       ZQRecebimento.SQL.Add('Select R.idrecebimento,R.documento,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.empresa, R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.nomeadversa,R.somar,');
       ZQRecebimento.SQL.Add('       R.juros,R.descontos,R.Data_Quitacao,R.tip,R.dt_nao_pagou_no_mes,R.Multa,R.Mora,R.Cliente, V.idvenda,V.datavenda,V.imovel,V.valorvenda,V.forma_reajuste,V.marca,V.Multa,V.Mora,V.tabela_price ');
       ZQRecebimento.SQL.Add('  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.idloteamento='+quotedstr(copy(FrmPesqCobranca.ednome.Text,1,posi))+' and R.adversa='+quotedstr(FrmPesqCobranca.ZQTempCliReceb.FieldByName('adversa').Text)+' and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento ');
       Label52.Caption:=copy(FrmPesqCobranca.ednome.Text,1,posi);
    end;
    ZQRecebimento.ParamByName('dt').AsDate:=date;
    LStep := PerformanceStart;
    ZQRecebimento.open;
    PerformanceElapsed('Cobranca: abrir titulos apos pesquisa', LStep);
    // 27/08/2012 tony pediu pra mudar
  //  ZQRecebimento.Locate('quadralote',FrmPesqCobranca.ZQTempCliRecebquadralote.Value,[]);
    ZQRecebimento.Locate('adversa',FrmPesqCobranca.ZQTempCliReceb.FieldByName('adversa').AsInteger,[]);
  end
  else
  begin


    if FrmPesqCobranca.ednome.Text='' then
    begin
       ZQRecebimento.SQL.Add('Select R.idrecebimento,R.documento,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.empresa, R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.nomeadversa,R.somar,');
       ZQRecebimento.SQL.Add('       R.juros,R.descontos,R.Data_Quitacao,R.tip,R.dt_nao_pagou_no_mes,R.Multa,R.Mora,R.Cliente, V.idvenda,V.datavenda,V.imovel,V.valorvenda,V.forma_reajuste,V.marca,V.Multa,V.Mora,V.tabela_price ');
//       ZQRecebimento.SQL.Add('  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.quadralote='+quotedstr(FrmPesqCobranca.ZQTempCliRecebquadralote.Value)+' and  R.adversa='+quotedstr(FrmPesqCobranca.ZQTempCliRecebadversa.Text)+' and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento ');
       ZQRecebimento.SQL.Add('  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.quadralote='+quotedstr(FrmPesqCobranca.ZQTempCliReceb.FieldByName('quadralote').AsString)+'  and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento ');
       Label52.Caption:='';
    end
    else
    begin
       posi := pos('|',FrmPesqCobranca.ednome.Text);
       dec(posi);
       ZQRecebimento.SQL.Add('Select R.idrecebimento,R.documento,R.usuario,R.Dt_Entrada,R.Dt_Vencimento,R.Valor,R.VrDoc,R.ordem,R.TipDoc,R.saldo,R.marcar,R.empresa, R.origem,R.adversa,R.recpag,R.numordem,R.idloteamento,R.venda_idvenda,R.quadralote,R.nomeadversa,R.somar,');
       ZQRecebimento.SQL.Add('       R.juros,R.descontos,R.Data_Quitacao,R.tip,R.dt_nao_pagou_no_mes,R.Multa,R.Mora,R.Cliente, V.idvenda,V.datavenda,V.imovel,V.valorvenda,V.forma_reajuste,V.marca,V.Multa,V.Mora,V.tabela_price ');
       ZQRecebimento.SQL.Add('  from Recebimento as R join venda as V on R.venda_idvenda=V.idvenda where R.idloteamento='+quotedstr(copy(FrmPesqCobranca.ednome.Text,1,posi))+' and R.adversa='+quotedstr(FrmPesqCobranca.ZQTempCliReceb.FieldByName('adversa').Text)+' and R.Dt_Vencimento< :dt and R.saldo>0 order by R.nomeadversa,R.DT_Vencimento ');
       Label52.Caption:=copy(FrmPesqCobranca.ednome.Text,1,posi);
    end;
    ZQRecebimento.ParamByName('dt').AsDate:=date;
    LStep := PerformanceStart;
    ZQRecebimento.open;
    PerformanceElapsed('Cobranca: abrir titulos apos pesquisa', LStep);
    // 27/08/2012 tony pediu pra mudar
    ZQRecebimento.Locate('quadralote',FrmPesqCobranca.ZQTempCliReceb.FieldByName('quadralote').AsString,[]);
  //  ZQRecebimento.Locate('adversa',FrmPesqCobranca.ZQTempCliRecebadversa.Value,[]);

  end;

  { A atualizacao completa ocorre depois do Refresh, uma unica vez. }

  ZQRecebimento_bancario.close;
  ZQRecebimento_bancario.SQL.Clear;
  ZQRecebimento_bancario.SQL.Add(' SELECT r.idrecebimento,r.documento,r.cliente,r.usuario,r.Dt_Entrada,r.Dt_Vencimento,r.Valor,r.Observ,r.VrDoc,r.ordem,r.TipDoc,r.saldo,r.marcar,r.RefBaixa,r.refvinda,');
  ZQRecebimento_bancario.SQL.Add('        r.contabil,r.empresa,r.custodaparcela,r.origem,r.adversa,r.recpag,r.numordem,r.idloteamento,r.venda_idvenda,r.quadralote,r.numboleto,r.Substituicao,r.sq,');
  ZQRecebimento_bancario.SQL.Add('        r.somar,r.nomeadversa,r.Reajustado,r.Data_reajuste,r.Proximo_Reajuste,r.Parcelas_fixas,r.juros,r.descontos,r.Data_Quitacao,');
  ZQRecebimento_bancario.SQL.Add('        il.incorporador_idincorporador,il.loteamento_idloteamento,il.percent_parte,il.codcontabancaria,il.digito_dif,');
  ZQRecebimento_bancario.SQL.Add('        cb.idconta_bancaria,cb.n_banco,cb.n_agencia,cb.n_agencia_v,cb.n_conta,cb.n_conta_v,cb.nomebanco,cb.n_no_banco,cb.nomeresposavel,cb.nometitular,cb.aberta,');
  ZQRecebimento_bancario.SQL.Add('        cb.senha_interna,cb.senha_conta,cb.ativa,cb.jurosemboleto,cb.taxadiaria,cb.valoroupercent,cb.carteira,cb.boleto,cb.idparticipante,cb.n_dif_empreed,cb.convenio,cb.doc_titular,cb.data_maxima_desconto,cb.Perc_descontos ');
  ZQRecebimento_bancario.SQL.Add(' FROM recebimento as r join incorporador_loteamento as il on il.loteamento_idloteamento=r.idloteamento join conta_bancaria as cb on cb.idconta_bancaria=il.codcontabancaria where r.adversa='+quotedstr(FrmPesqCobranca.ZQTempCliReceb.FieldByName('adversa').Text)+' and r.Dt_Vencimento< :dt and r.saldo>0 group by r.idloteamento order by r.idloteamento ');
  ZQRecebimento_bancario.ParamByName('dt').AsDate:=date;
  LStep := PerformanceStart;
  ZQRecebimento_bancario.open;
  PerformanceElapsed('Cobranca: abrir dados bancarios', LStep);

  FrmPesqCobranca.ZQTempCliReceb.close;
  ZQVenda_cobr.Close;
  ZQVenda_cobr.SQL.Clear;
  ZQVenda_cobr.SQL.Add('Select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao from venda where idvenda='+quotedstr(ZQRecebimento.FieldByName('venda_idvenda').Text));
  ZQVenda_cobr.open;
  RGSelecao.Visible:=true;

  FinalizarPesquisaCobranca;
  DBcobranca.SetFocus;
  PerformanceElapsed('Cobranca: preencher grid apos selecao (inclui subetapas)', LOperacao);
  FlushPerformanceLog;

end;

procedure TFrmCad_Recebimento.ZQRecebimento4CalcFields(DataSet: TDataSet);
var
  LDias: Integer;
  LMulta, LMoraCalculada, LCorrigido: Double;
  LStep: UInt64;
begin
  PrepararTaxasTituloCobranca;
  LStep := PerformanceStart;
  LDias := StrToInt(DifDias(
    ZQRecebimento4.FieldByName('Dt_Vencimento').AsDateTime, Date));
  CalcularValoresCobranca(
    ZQRecebimento4.FieldByName('saldo').AsFloat,
    FTaxasTotaisCobranca.FieldByName('Multa').AsFloat,
    FTaxasTotaisCobranca.FieldByName('Mora').AsFloat,
    FTaxasTotaisCobranca.FieldByName('tabela_price').AsFloat, LDias,
    Lmora.Caption <> 'N', LMulta, LMoraCalculada, LCorrigido);
  ZQRecebimento4.FieldByName('Dias').AsInteger := LDias;
  ZQRecebimento4.FieldByName('mora_dia').AsFloat := LMoraCalculada;
  ZQRecebimento4.FieldByName('Multa_Contrato').AsFloat := LMulta;
  ZQRecebimento4.FieldByName('parcela_corrigida').AsFloat := LCorrigido;
  PerformanceSample(smCalculoFinanceiro, LStep);
end;
procedure TFrmCad_Recebimento.DBcobrancaDblClick(Sender: TObject);
begin
  if cbmontante.Checked=true then
  begin
    DBcobranca.Options:=DBcobranca.Options+[dgMultiSelect];
  end
  else
  begin
    DBcobranca.Options:=DBcobranca.Options-[dgMultiSelect];
  end;
end;

procedure TFrmCad_Recebimento.dxButton7Click(Sender: TObject);
begin
  if FrmRelCobranca=nil then
     FrmRelCobranca:=TFrmRelCobranca.Create(Self);
  FrmRelCobranca.RLReport1.PreviewModal;
  FrmRelCobranca:=nil;
end;

procedure TFrmCad_Recebimento.dxButton9Click(Sender: TObject);
begin
  if FPaintBoxGrafico3 <> nil then
    PrintPaintBoxChart(FPaintBoxGrafico3,
      'Grafico de Inadimplencias ate a data de ' + DateToStr(Date - 1), True);
end;

procedure TFrmCad_Recebimento.JBProcessoItems4Click(Sender: TObject);
Var
  Varregidrec : integer;
begin
  if not Verif_senha('Financeiro','Substituição ','Cliente: '+DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString) then exit;
  Varregidrec := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
  if FrmRecebBaixa_subst=nil then
     FrmRecebBaixa_subst:=TFrmRecebBaixa_subst.Create(Self);
  FrmRecebBaixa_subst.Label9.Caption := 'Substituição de Títulos em Aberto do Contas a Pagar';
  FrmRecebBaixa_subst.Label9.Color := $0000BB00;
  if DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString = 'R' Then Begin
    FrmRecebBaixa_subst.Label9.Caption := 'Substituição de Títulos em Aberto do Contas a Receber';
    FrmRecebBaixa_subst.Label9.Color := ClRed;
  end;
  JBProcesso.Visible:=false;
  AbrirModal(Self, FrmRecebBaixa_subst);
  FreeAndNil(FrmRecebBaixa_subst);
  FrmRecebBaixa_subst:=nil;
  Panel1.Visible:=true;
  Panel2.Visible:=true;
  Application.ProcessMessages;

//  DM_tabelas.ZQRecebimento.Refresh;
  if DM_TAbelas.ZQRecebBxTemp.active=true then
     DM_TAbelas.zqrecebbxTemp.Refresh;
  if ZQGerou.active=true then
     ZQGerou.Refresh;
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento  r where not exists (select 1 from rescisao rs where r.venda_idvenda=rs.idvenda) and r.IdRecebimento='+quotedstr(Varregidrec.ToString)+'  order by r.DT_Vencimento');
  DM_tabelas.ZQRecebimento.open;
//  DM_tabelas.zqRecebimento.Refresh;
  DM_tabelas.zqRecebimento.Locate('IdRecebimento',Varregidrec,[]);
  DBGReceb.SetFocus;
  Panel1.Visible:=false;
  Panel2.Visible:=false;
  Application.ProcessMessages;
 // Botoes_setas;
end;

procedure TFrmCad_Recebimento.dxButton10Click(Sender: TObject);
var
  LStep: UInt64;
begin
  LStep := PerformanceStart;
  PrepararLookupsCobranca;
  AtualizarCamposCobranca;
  if FrmRelCobranca2=nil then
     FrmRelCobranca2:=TFrmRelCobranca2.Create(application);
  FrmRelCobranca2.RLReport1.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText2.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText2.DataField:='nome_loteamento';
  FrmRelCobranca2.RLDBText12.DataSource:=FrmCad_Recebimento.DS_ZQRecebimento_bancario;
  FrmRelCobranca2.RLDBText12.DataField:='n_banco';
  FrmRelCobranca2.RLDBText4.DataSource:=FrmCad_Recebimento.DS_ZQRecebimento_bancario;
  FrmRelCobranca2.RLDBText4.DataField:='n_agencia';
  FrmRelCobranca2.RLDBText6.DataSource:=FrmCad_Recebimento.DS_ZQRecebimento_bancario;
  FrmRelCobranca2.RLDBText6.DataField:='n_agencia_v';
  FrmRelCobranca2.RLDBText9.DataSource:=FrmCad_Recebimento.DS_ZQRecebimento_bancario;
  FrmRelCobranca2.RLDBText9.DataField:='n_conta';
  FrmRelCobranca2.RLDBText11.DataSource:=FrmCad_Recebimento.DS_ZQRecebimento_bancario;
  FrmRelCobranca2.RLDBText11.DataField:='n_conta_v';
  FrmRelCobranca2.RLDBText18.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText18.DataField:='adversanome';
  FrmRelCobranca2.RLDBText1.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText1.DataField:='CPF';
  FrmRelCobranca2.RLDBText20.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText20.DataField:='quadralote';
  FrmRelCobranca2.RLDBText8.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText8.DataField:='Dt_Vencimento';
  FrmRelCobranca2.RLDBText10.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText10.DataField:='saldo';
  FrmRelCobranca2.RLDBText5.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText5.DataField:='Dias';
  FrmRelCobranca2.RLDBText16.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText16.DataField:='Multa_Contrato';
  FrmRelCobranca2.RLDBText7.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText7.DataField:='mora_dia';
  FrmRelCobranca2.RLDBText3.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBText3.DataField:='parcela_corrigida';
  FrmRelCobranca2.RLDBResult1.DataSource:=FrmCad_Recebimento.DS_Recebimento;
  FrmRelCobranca2.RLDBResult1.DataField:='parcela_corrigida';


  PerformanceElapsed('Relatorio do grid: preparar dados e componentes', LStep);
  FlushPerformanceLog;
  ZQRecebimento.DisableControls;
  FrmRelCobranca2.RLReport1.Previewmodal;
  ZQRecebimento.First;
  ZQRecebimento.EnableControls;
  FrmRelCobranca2:=nil;
end;

procedure TFrmCad_Recebimento.Label5MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  //mover formulario sem bordas
  MovimentaObject(sender,button,shift,x,y,FrmCad_Recebimento);

end;

procedure TFrmCad_Recebimento.JBProcessoItems5Click(Sender: TObject);
begin
  if not Verif_senha('Trocar Loteamento','Tela Principal','') then exit;
  if DM_tabelas.ZQQuadras.Active=false then
     DM_tabelas.ZQQuadras.open;
  if DM_Tabelas.ZQincorp_loteame.Active=false then
     DM_Tabelas.ZQincorp_loteame.Open;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  if DM_tabelas.ZQNumOrdem.Active=false then
     DM_tabelas.ZQNumOrdem.Open;
  if DM_Tabelas.ZQImovel.Active=false then
     DM_Tabelas.ZQImovel.open;
  if DM_Tabelas.ZQVendedor.Active=false then
     DM_Tabelas.ZQVendedor.open;
  if DM_Tabelas.ZQVenda.Active=false then
     DM_Tabelas.ZQVenda.open;
  if DM_Tabelas.ZQComprador.Active=false then
     DM_Tabelas.ZQComprador.open;
  if Frm_Trocar=nil then
     Frm_Trocar:=Tfrm_Trocar.Create(Self);
  Frm_Trocar.Left := FrmCad_Recebimento.Left+7;
  Frm_Trocar.Top := FrmCad_Recebimento.Top+135;
  Frm_Trocar.show;
  JBProcesso.Visible:=false;

end;

procedure TFrmCad_Recebimento.JBProcessoItems6Click(Sender: TObject);
begin
  JBProcesso.Visible:=false;
  if not Verif_senha('Recebimento','Parcelas não Pagas no Mês','') then exit;
     if frmnaopaga=nil then
        frmnaopaga:=Tfrmnaopaga.Create(Self);
     AbrirModal(Self, frmnaopaga);
end;

procedure TFrmCad_Recebimento.JBProcessoItems7Click(Sender: TObject);
begin

//  FrmAditamento.Top := FrmCad_Recebimento.Top+135;
  if not Verif_senha('Balão','Tela Principal','') then exit;
  if Frmbalao=nil then
     Frmbalao:=TFrmbalao.Create(Self);
  Frmbalao.Left := FrmCad_Recebimento.Left+7;
  Frmbalao.show;
  JBProcesso.Visible:=false;

end;

procedure TFrmCad_Recebimento.JBProcessoItems8Click(Sender: TObject);
begin

  if not Verif_senha('Cobrança','Tela Principal','') then exit;
  if frmcobranca=nil then
     frmcobranca:=Tfrmcobranca.Create(Self);
  frmcobranca.Left := FrmCad_Recebimento.Left+7;
  frmcobranca.codcli.text:=DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  frmcobranca.nomecli.text:=DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  AbrirModal(Self, frmcobranca);
  frmcobranca:=nil;
  JBProcesso.Visible:=false;
end;

procedure TFrmCad_Recebimento.JBProcessoItems9Click(Sender: TObject);
begin
  if not Verif_senha('Trocar Empreendimento','Tela Principal','') then exit;
//  if DM_tabelas.ZQQuadras.Active=false then
//     DM_tabelas.ZQQuadras.open;
//  if DM_Tabelas.ZQincorp_loteame.Active=false then
//     DM_Tabelas.ZQincorp_loteame.Open;
//  if DM_Tabelas.ZQLoteamento.Active=false then
//     DM_Tabelas.ZQLoteamento.Open;
//  if DM_tabelas.ZQNumOrdem.Active=false then
//     DM_tabelas.ZQNumOrdem.Open;
//  if DM_Tabelas.ZQImovel.Active=false then
//     DM_Tabelas.ZQImovel.open;
//  if DM_Tabelas.ZQVendedor.Active=false then
//     DM_Tabelas.ZQVendedor.open;
//  if DM_Tabelas.ZQVenda.Active=false then
//     DM_Tabelas.ZQVenda.open;
//  if DM_Tabelas.ZQComprador.Active=false then
//     DM_Tabelas.ZQComprador.open;
  if Frm_Trocar_Empre=nil then
     Frm_Trocar_Empre:=TFrm_Trocar_Empre.Create(Self);
  Frm_Trocar_Empre.Left := FrmCad_Recebimento.Left+7;
  Frm_Trocar_Empre.Top := FrmCad_Recebimento.Top+135;
  Frm_Trocar_Empre.show;
  JBProcesso.Visible:=false;

end;

procedure TFrmCad_Recebimento.RGSelecaoClick(Sender: TObject);
begin
  RGSelecao.Visible:=false;
  if RGSelecao.ItemIndex=0 then
  begin
    dt_selecao.Date:=date;
    dt_selecao.Visible:=true;
    dt_selecao.SetFocus;
    fpg.Clear;
    fpg.Visible:=false;
    Label48.Visible:=false;
    exit;
  end
  else
  begin
    dt_selecao.Visible:=false;
    fpg.Visible:=true;
    Label48.Visible:=true;
    fpg.Clear;
    fpg.SetFocus;
    exit;
  end;
end;

procedure TFrmCad_Recebimento.dt_selecaoExit(Sender: TObject);
var
posi:integer;
begin
  posi:=0;
  fpg.Text:=DifDias(Date,dt_selecao.Date);
  posi:=pos(',',fpg.Text);
  if posi>0 then
  begin
    fpg.Text:=copy(fpg.Text,1,posi-1);
  end;
  FrmCad_Recebimento.fpgExit(sender);
end;


procedure TFrmCad_Recebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  { Estes eventos nao possuem manipuladores no DFM original. }
  ZQRecebimento.BeforeOpen := CobrancaAntesReabrir;
  ZQRecebimento.BeforeRefresh := CobrancaAntesReabrir;
  ZQRecebimento.AfterPost := CobrancaAntesReabrir;
  ZQRecebimento.AfterDelete := CobrancaAntesReabrir;
  { O DFM binario pode manter as aspas escapadas do texto-fonte dentro do
    SQL. Corrige somente a consulta local antes de preparar seus campos. }
  ZQRecebimento4.Close;
  ZQRecebimento4.SQL.Clear;
  ZQRecebimento4.SQL.Add('Select * from Cobranca where saldo > ''0''');
  EnsureRuntimeFields(Self);
  CriarGraficosRuntime;
end;

initialization
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQVIncRec', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecidRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRectipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecdt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVIncRec', 'ZQVIncRecrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQOrigem', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemidRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemdt_entrada', 'dt_entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemdt_vencimento', 'dt_vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemtipdoc', 'tipdoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQOrigem', 'ZQOrigemrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQTempDivid', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQTempDivid', 'ZQTempDivididdividido', 'iddividido', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQTempDivid', 'ZQTempDividparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQTempDivid', 'ZQTempDividpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQTempDivid', 'ZQTempDividplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQTempDivid', 'ZQTempDividtipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQTempDivid', 'ZQTempDividordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQGerou', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQGerou', 'ZQGerouidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQGerou', 'ZQGerouidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQGerou', 'ZQGerouquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQGerou', 'ZQGerourefbaixa', 'refbaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQGerou', 'ZQGerourefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQGerou', 'ZQGerourefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQResult1', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1idRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1Observ', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1marcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1origem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1soma', 'soma', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQResult1', 'ZQResult1contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQFeito', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQFeito', 'ZQFeitorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQFeito', 'ZQFeitosomafeito', 'somafeito', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQVenda', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQParcela', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQEntrada', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  { Estes tres campos sao calculados localmente para que a cobrança possa
    preservar o valor alternativo gravado no titulo quando o ID do lookup
    antigo nao possuir mais correspondencia. }
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkCalculated, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoDias', 'Dias', TIntegerField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoEndereco', 'Endereco', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'endereco', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoCPF', 'CPF', TWideStringField, fkCalculated, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentofone', 'fone', TWideStringField, fkLookup, 15, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Fone1', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentofone2', 'fone2', TWideStringField, fkLookup, 15, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'fone2', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentofone3', 'fone3', TWideStringField, fkLookup, 15, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'fone3', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoMulta_Contrato', 'Multa_Contrato', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentomora_dia', 'mora_dia', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoparcela_corrigida', 'parcela_corrigida', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoagencia', 'agencia', TWideStringField, fkLookup, 6, 0, False, '', '', '', '', 0, 'idloteamento', 'ZQRecebimento_bancario', 'idloteamento', 'n_agencia', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoagencia_v', 'agencia_v', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'idloteamento', 'ZQRecebimento_bancario', 'idloteamento', 'n_agencia_v', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoconta', 'conta', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'idloteamento', 'ZQRecebimento_bancario', 'idloteamento', 'n_conta', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoconta_v', 'conta_v', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'idloteamento', 'ZQRecebimento_bancario', 'idloteamento', 'n_conta_v', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentonomebanco', 'nomebanco', TWideStringField, fkLookup, 40, 0, False, '', '', '', '', 0, 'idloteamento', 'ZQRecebimento_bancario', 'idloteamento', 'n_banco', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentotip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentodt_nao_pagou_no_mes', 'dt_nao_pagou_no_mes', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoMulta', 'Multa', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoMora', 'Mora', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoCliente', 'Cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentomarca', 'marca', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoMulta_1', 'Multa_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentoMora_1', 'Mora_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento', 'ZQRecebimentotabela_price', 'tabela_price', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQRecebimento2', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2parcelas', 'parcelas', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2titulos', 'titulos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2Multa_Contrato', 'Multa_Contrato', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2mora_dia', 'mora_dia', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2parcela_corrigida', 'parcela_corrigida', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2Dias', 'Dias', TIntegerField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2multa', 'multa', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2mora', 'mora', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento2', 'ZQRecebimento2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQRecebimento4', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4documento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Observ', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4ordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4TipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4marcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4origem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4recpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4nomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4adversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4nome_loteamento', 'nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4quadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4numboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Substituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4sq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4nomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Reajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Data_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4somar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Proximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Parcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Dias', 'Dias', TIntegerField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Endereco', 'Endereco', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'endereco', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4CPF', 'CPF', TWideStringField, fkLookup, 18, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'doc1', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4fone', 'fone', TWideStringField, fkLookup, 15, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Fone1', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4fone2', 'fone2', TWideStringField, fkLookup, 15, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'fone2', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4fone3', 'fone3', TWideStringField, fkLookup, 15, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'fone3', True);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4Multa_Contrato', 'Multa_Contrato', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4mora_dia', 'mora_dia', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento4', 'ZQRecebimento4parcela_corrigida', 'parcela_corrigida', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQInadimplentes', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentescliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentessaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentessomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesdescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesinadimplente', 'inadimplente', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesnome_loteamento_sql', 'nome_loteamento_sql', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQInadimplentes', 'ZQInadimplentesloteamento', 'loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQRecebimento_bancario', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarionumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarionomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioincorporador_idincorporador', 'incorporador_idincorporador', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariopercent_parte', 'percent_parte', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariocodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariodigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_banco', 'n_banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_agencia', 'n_agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_agencia_v', 'n_agencia_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_conta', 'n_conta', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_conta_v', 'n_conta_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarionomebanco', 'nomebanco', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_no_banco', 'n_no_banco', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarionomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarionometitular', 'nometitular', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariosenha_interna', 'senha_interna', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariosenha_conta', 'senha_conta', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioativa', 'ativa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariojurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariotaxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariovaloroupercent', 'valoroupercent', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariocarteira', 'carteira', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioboleto', 'boleto', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarion_dif_empreed', 'n_dif_empreed', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioconvenio', 'convenio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariodoc_titular', 'doc_titular', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancariodata_maxima_desconto', 'data_maxima_desconto', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQRecebimento_bancario', 'ZQRecebimento_bancarioPerc_descontos', 'Perc_descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_Recebimento, 'ZQVenda_cobr', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda_cobr', 'ZQVenda_cobridvenda', 'idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda_cobr', 'ZQVenda_cobrMulta', 'Multa', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_Recebimento, 'ZQVenda_cobr', 'ZQVenda_cobrMora', 'Mora', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.



