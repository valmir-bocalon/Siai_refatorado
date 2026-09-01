unit Venda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XDBNum, XNum, StdCtrls, wwdbdatetimepicker, Mask, DBCtrls,
  Grids, DBGrids, ComCtrls, TabNotBk,   dxButton, ExtCtrls, XBanner,
  XDate, XDBEdit, DB, JvExControls, JvXPCore, JvXPBar;

type
  TFrm_Venda = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTExcluir: TdxButton;
    BtEditar: TdxButton;
    BtIncluir: TdxButton;
    BtCancelar: TdxButton;
    BtGravar: TdxButton;
    BTPesquisar: TdxButton;
    BtRelatorio: TdxButton;
    BtUltimo: TdxButton;
    BtProximo: TdxButton;
    BTAnterior: TdxButton;
    BtPrimeiro: TdxButton;
    BTFechar: TdxButton;
    BtProcessar: TdxButton;
    Pag_Venda: TTabbedNotebook;
    XBanner1: TXBanner;
    DBGVenda: TDBGrid;
    DBECod: TDBEdit;
    Label1: TLabel;
    Label5: TLabel;
    DBEDAta: TwwDBDateTimePicker;
    Comprador: TGroupBox;
    Corretor: TGroupBox;
    XBanner2: TXBanner;
    XBanner3: TXBanner;
    DBGrid2: TDBGrid;
    Label14: TLabel;
    DBEImovel: TDBEdit;
    GBEntrada: TGroupBox;
    XBanner7: TXBanner;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    XNEEntrada: TXNumEdit;
    XNQtdeEnt: TXNumEdit;
    DBEEntrada: TDBGrid;
    DBEVencimento: TwwDBDateTimePicker;
    DBEValor: TXDBNumEdit;
    XBanner9: TXBanner;
    Label23: TLabel;
    Label31: TLabel;
    XVencimentoEnt: TXDateEdit;
    EContabilEntrada: TEdit;
    Label6: TLabel;
    Label8: TLabel;
    DXBAchaLote: TdxButton;
    DBELote: TDBEdit;
    DBEdit6: TDBEdit;
    Label2: TLabel;
    Label4: TLabel;
    Label9: TLabel;
    DBEEmpreend: TDBEdit;
    EComprador: TEdit;
    Ecorretor: TEdit;
    DBGrid1: TDBGrid;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    Label3: TLabel;
    DBEdit3: TDBEdit;
    Label7: TLabel;
    Label10: TLabel;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    DBEdit10: TDBEdit;
    Labelx: TLabel;
    DBEOrdem: TDBEdit;
    Label17: TLabel;
    XDBNumEdit1: TXDBNumEdit;
    Label18: TLabel;
    CBPeriodoEnt: TComboBox;
    Label20: TLabel;
    Label21: TLabel;
    Label27: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    DBEBanco: TDBEdit;
    DBEAgencia: TDBEdit;
    DBEConta: TDBEdit;
    DBEDocumento: TDBEdit;
    DBENumeroEntrada: TDBEdit;
    DBCBDocVenda: TDBComboBox;
    GBParcelas: TGroupBox;
    XBanner4: TXBanner;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    XNParcela: TXNumEdit;
    XNQtdeParc: TXNumEdit;
    DBGPartelas: TDBGrid;
    DBEVencimentoParc: TwwDBDateTimePicker;
    DBEValorParc: TXDBNumEdit;
    XVencimentoParc: TXDateEdit;
    EContabilEntradaParc: TEdit;
    DBEOrdemParc: TDBEdit;
    DBEBancoParc: TDBEdit;
    DBEAgendaParc: TDBEdit;
    DBEContaParc: TDBEdit;
    DBEDocumentoParc: TDBEdit;
    DBENumeroParc: TDBEdit;
    DBCBDocVendaParc: TDBComboBox;
    CBPeriodoParc: TComboBox;
    EloteTemp: TEdit;
    Label22: TLabel;
    DBEVrvenda: TXDBNumEdit;
    Label50: TLabel;
    XDBNumEdit2: TXDBNumEdit;
    DBEDonoEntrada: TDBEdit;
    Label51: TLabel;
    Label52: TLabel;
    DBEDonoParcela: TDBEdit;
    JvXPBar1: TJvXPBar;
    procedure FormShow(Sender: TObject);
    procedure botoes;
    procedure Atualiza;
    procedure Atualiza_comprador;
    procedure DesativaCampos;
    procedure AtivaCampos;
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtProcessarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure BtRelatorioClick(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DXBAchaLoteClick(Sender: TObject);
    procedure ECompradorExit(Sender: TObject);
    procedure CompradorEnter(Sender: TObject);
    procedure DBGrid2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CorretorEnter(Sender: TObject);
    procedure EcorretorExit(Sender: TObject);
    procedure CorretorExit(Sender: TObject);
    procedure DBGVendaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGVendaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGVendaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGVendaEnter(Sender: TObject);
    procedure DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGrid2Enter(Sender: TObject);
    procedure DBEEntradaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBEOrdemChange(Sender: TObject);
    procedure GBEntradaEnter(Sender: TObject);
    procedure CBPeriodoEntExit(Sender: TObject);
    procedure atualiza_tela_venda_entrada;
    procedure atualiza_tela_venda_parcela;
    procedure atualizatebelaspagamento;
    procedure DBEEntradaEnter(Sender: TObject);
    procedure DBEEntradaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EContabilEntradaExit(Sender: TObject);
    procedure EContabilEntradaEnter(Sender: TObject);
    procedure DBEValorExit(Sender: TObject);
    procedure DBEValorEnter(Sender: TObject);
    procedure DBEDocumentoExit(Sender: TObject);
    procedure XVencimentoEntExit(Sender: TObject);
    procedure DBCBDocVendaExit(Sender: TObject);
    procedure XNEEntradaExit(Sender: TObject);
    procedure GBParcelasEnter(Sender: TObject);
    procedure CBPeriodoParcExit(Sender: TObject);
    procedure XVencimentoParcExit(Sender: TObject);
    procedure DBEDocumentoParcExit(Sender: TObject);
    procedure DBGPartelasDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBEVrvendaExit(Sender: TObject);
    procedure DBCBDocVendaParcExit(Sender: TObject);
    procedure EContabilEntradaParcEnter(Sender: TObject);
    procedure EContabilEntradaParcExit(Sender: TObject);
    procedure DBEOrdemParcChange(Sender: TObject);
    procedure DBEValorParcEnter(Sender: TObject);
    procedure DBEValorParcExit(Sender: TObject);
    procedure DBEEntradaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGPartelasKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGPartelasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGPartelasEnter(Sender: TObject);
    procedure DBEDonoEntradaExit(Sender: TObject);
    procedure DBEDonoParcelaExit(Sender: TObject);
    procedure JvXPBar1Items0Click(Sender: TObject);
    procedure JvXPBar1Items1Click(Sender: TObject);
    procedure JvXPBar1Exit(Sender: TObject);
    procedure BTPesquisarClick(Sender: TObject);
    procedure Pag_VendaChange(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Venda: TFrm_Venda;

implementation

uses tabelas, Funcoes, AchaLoteVenda, RelVenda, QuadroResumo, AchaVenda,
  Loteamento, uRuntimeFields;

{$R *.dfm}


procedure TFrm_Venda.FormShow(Sender: TObject);
begin
  Pag_Venda.PageIndex := 0;
  DBCBDocVenda.Clear;
  DBCBDocVendaParc.Clear;
  DM_Tabelas.ZQTipoDoc.First;
  While not DM_Tabelas.ZQTipoDoc.Eof do Begin
    DBCBDocVenda.Items.Add(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DBCBDocVendaParc.Items.Add(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_Tabelas.ZQTipoDoc.Next;
  end;
  DesativaCampos;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.botoes;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQVenda.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQVenda.RecNo=DM_Tabelas.ZQVenda.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  Atualiza;
End;

procedure TFrm_Venda.Atualiza;
Begin
  DM_TAbelas.CDSCompradorTemp.Close;
  DM_Tabelas.CDSCompradorTemp.CreateDataSet;
  DM_Tabelas.ZQComprador.First;
  while not DM_Tabelas.ZQComprador.Eof do begin
    DM_Tabelas.CDSCompradorTemp.Insert;
    DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').Value;
    DM_tabelas.CDSCompradorTemp.Post;
    dm_tabelas.ZQComprador.Next;
  end;
  DM_Tabelas.CDSVendedorTemp.Close;
  DM_Tabelas.CDSVendedorTemp.CreateDataSet;
  DM_Tabelas.ZQVendedor.First;
  While not DM_Tabelas.ZQVendedor.Eof do Begin
    DM_Tabelas.CDSVendedorTemp.Insert;
    DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').Value;
     DM_Tabelas.CDSVendedorTemp.Post;
    DM_Tabelas.ZQVendedor.Next;
  end;
  Atualiza_comprador;
  if Pag_Venda.PageIndex = 1 Then
    Atualizatebelaspagamento;
End;

procedure TFrm_Venda.Atualiza_comprador;
Begin
  Label10.Visible := True;
  if DM_Tabelas.CDSCompradorTemppessoa.Value = 'F' Then Begin
    Label7.Caption := 'C.P.F.';
    Label10.Caption := 'R.G.';
  end
  else if DM_Tabelas.CDSCompradorTemppessoa.Value = 'J' Then Begin
    Label7.Caption := 'C.N.P.J.';
    Label10.Caption := 'Insc. Est.';
  end
  else if DM_Tabelas.CDSCompradorTemppessoa.Value = 'R' Then Begin
    Label7.Caption := 'C.N.P.J.';
    Label10.Caption := 'Insc. Rural.';
  end
  else if DM_Tabelas.CDSCompradorTemppessoa.Value = 'P' Then Begin
    Label7.Caption := 'C.N.P.J.';
    Label10.Visible := False;
  end;
end;

procedure TFrm_Venda.DesativaCampos;
Begin
  DBEDAta.ReadOnly := True;
  Ecomprador.ReadOnly := True;
  Ecorretor.ReadOnly := True;
  DXBAchaLote.Enabled := False;
  DBEVrvenda.ReadOnly := True;
  XNEEntrada.ReadOnly := True;
  XNQtdeEnt.ReadOnly := True;
  XVencimentoEnt.ReadOnly := True;
  CBPeriodoEnt.Enabled := False;
  EContabilEntrada.ReadOnly := True;
  DBCBDocVenda.ReadOnly := True;
  DBEVencimento.ReadOnly := True;
  DBEValor.ReadOnly := True;
  DBEBanco.ReadOnly := True;
  DBEAgencia.ReadOnly := True;
  DBEConta.ReadOnly := True;
  DBENumeroEntrada.ReadOnly := True;
  DBEDocumento.ReadOnly := True;
  XNParcela.ReadOnly := True;
  XNQtdeParc.ReadOnly := True;
  XVencimentoParc.ReadOnly := True;
  CBPeriodoParc.Enabled := False;
  EContabilEntradaParc.ReadOnly := True;
  DBCBDocVendaParc.ReadOnly := True;
  DBEVencimentoParc.ReadOnly := True;
  DBEValorParc.ReadOnly := True;
  DBEBancoParc.ReadOnly := True;
  DBEAgendaParc.ReadOnly := True;
  DBEContaParc.ReadOnly := True;
  DBENumeroParc.ReadOnly := True;
  DBEDocumentoParc.ReadOnly := True;
  BTPrimeiro.Enabled := true;
  BTAnterior.Enabled := true;
  BTProximo.Enabled := true;
  BtUltimo.Enabled := true;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtProcessar.Enabled := true;
  BtIncluir.Enabled := true;
  BtEditar.Enabled := true;
  BtExcluir.Enabled := true;
  BtPesquisar.Enabled := true;
  BtRelatorio.Enabled := true;
  DBGVenda.Enabled := True;
end;

procedure TFrm_Venda.AtivaCampos;
Begin
  DBEDAta.ReadOnly := False;
  Ecomprador.ReadOnly := False;
  Ecorretor.ReadOnly := False;
  DXBAchaLote.Enabled := True;
  DBEVrvenda.ReadOnly := False;
  XNEEntrada.ReadOnly := False;
  XNQtdeEnt.ReadOnly := False;
  XVencimentoEnt.ReadOnly := False;
  CBPeriodoEnt.Enabled := True;
  EContabilEntrada.ReadOnly := False;
  DBCBDocVenda.ReadOnly := False;
  DBEVencimento.ReadOnly := False;
  DBEValor.ReadOnly := False;
  DBEBanco.ReadOnly := False;
  DBEAgencia.ReadOnly := False;
  DBEConta.ReadOnly := False;
  DBENumeroEntrada.ReadOnly := False;
  DBEDocumento.ReadOnly := False;
  DBEDonoEntrada.ReadOnly := False;
  DBEDonoParcela.ReadOnly := False;
  XNParcela.ReadOnly := False;
  XNQtdeParc.ReadOnly := False;
  XVencimentoParc.ReadOnly := False;
  CBPeriodoParc.Enabled := True;
  EContabilEntradaParc.ReadOnly := False;
  DBCBDocVendaParc.ReadOnly := False;
  DBEVencimentoParc.ReadOnly := False;
  DBEValorParc.ReadOnly := False;
  DBEBancoParc.ReadOnly := False;
  DBEAgendaParc.ReadOnly := False;
  DBEContaParc.ReadOnly := False;
  DBENumeroParc.ReadOnly := False;
  DBEDocumentoParc.ReadOnly := False;
  BTPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BTProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtProcessar.Enabled := False;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BtExcluir.Enabled := False;
  BtPesquisar.Enabled := False;
  BtRelatorio.Enabled := False;
  DBGVenda.Enabled := False;
end;

procedure TFrm_Venda.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.First;
  DBGVenda.SetFocus;
end;


procedure TFrm_Venda.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Prior;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Next;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Last;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtGravarClick(Sender: TObject);
var
  varnumordemE, varnumordemP, varvezes, varvenda : integer;
begin
  if DM_Tabelas.ZQVenda.FieldByName('imovel').Value=0 Then Begin
    showmessage('Deve ser indicado um lote para ser vendido...');
    Pag_Venda.PageIndex := 0;
    DXBAchaLote.SetFocus;
    exit;
  end;
  if DM_Tabelas.ZQVenda.FieldByName('valorvenda').Value<DM_Tabelas.ZQVenda.FieldByName('valorvenal').Value Then Begin
    Showmessage('O valor de VENDA n�o deve ser menor que o valor venal...');
    Pag_Venda.PageIndex := 1;
    DBEVrvenda.SetFocus;
    exit;
  end;
  if DM_Tabelas.CDSCompradorTemp.RecordCount=0 Then Begin
    Showmessage('Ao menos um COMPRADOR deve ser informado para que a venda possa ser concluida..');
    Pag_Venda.PageIndex := 0;
    EComprador.SetFocus;
    exit;
  end;
  if DM_Tabelas.CDSVendedorTemp.RecordCount=0 Then Begin
    Showmessage('Ao menos um VENDEDOR deve ser informado para que a venda possa ser concluida..');
    Pag_Venda.PageIndex := 0;
    Ecorretor.SetFocus;
    exit;
  end;
  DM_Tabelas.ZQVenda.Post;
  if not empty(EloteTemp.Text) Then Begin
    DM_Tabelas.ZQImovel.Locate('idimovel',strtoint(EloteTemp.Text),[]);
    DM_Tabelas.ZQImovel.Edit;
    DM_Tabelas.ZQimovel.FieldByName('disponivel').Value := 'SIM';
  end;
  DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').Value,[]);
  DM_Tabelas.ZQImovel.Edit;
  DM_Tabelas.ZQimovel.FieldByName('disponivel').Value := 'N�O';
  DM_Tabelas.ZQImovel.Post;
  DM_Tabelas.ZQComprador.First;
  while not DM_Tabelas.ZQComprador.Eof do
    DM_Tabelas.ZQComprador.Delete;
  DM_Tabelas.CDSCompradorTemp.First;
  while not DM_Tabelas.CDSCompradorTemp.Eof do begin
    DM_Tabelas.ZQComprador.Insert;
    DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
    DM_Tabelas.CDSCompradorTemp.Next;
  end;
  DM_Tabelas.ZQVendedor.First;
  while not DM_Tabelas.ZQVendedor.Eof do
    DM_Tabelas.ZQVendedor.Delete;
  DM_Tabelas.CDSVendedorTemp.First;
  while not DM_Tabelas.CDSVendedorTemp.Eof do begin
    DM_Tabelas.ZQVendedor.Insert;
    DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').Value := DM_Tabelas.CDSVendedorTempcodcorretor.Value;
    DM_Tabelas.CDSVendedorTemp.Next;
  end;
  if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text,DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-E-']),[]) Then
    varnumordemE := DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value
  else Begin
    DM_tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
    varnumordemE := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Value;
  end;
  if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text,DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-P-']),[]) Then
    varnumordemp := DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value
  else Begin
    DM_tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
    varnumordemp := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Value;
  end;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento');
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento');
  DM_Tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').Value,[]);
  while not DM_Tabelas.ZQincorp_loteame.Eof do begin
    DM_Tabelas.CDSCompradorTemp.First;
    DM_Tabelas.CDSPagEntrada.First;
    while not DM_Tabelas.CDSPagEntrada.Eof do Begin
      DM_Tabelas.ZQRecebimento.Insert;
      DM_Tabelas.ZQRecebimento.FieldByName('idloteamento').Value := DM_Tabelas.ZQVenda.FieldByName('codloteamento').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('documento').Value := DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-E- '+CBPeriodoEnt.Text;
      DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').Value := DM_Tabelas.CDSPagEntradatipodoc.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('cliente').Value := DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := date;
      DM_Tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := DM_Tabelas.CDSPagEntradavencimento.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('ordem').Value := inttostr(varnumordemE)+'-'+DM_Tabelas.CDSPagEntradaordem.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('saldo').Value := DM_Tabelas.CDSPagEntradavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').Value/100;
      DM_Tabelas.ZQRecebimento.FieldByName('Valor').Value := DM_Tabelas.CDSPagEntradavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').Value/100;
      DM_Tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNEEntrada.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('contabil').Value := DM_Tabelas.CDSPagEntradacontabil.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('origem').Value := 'V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text;
      DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := 'R';
      DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := varnumordemE;
      DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value := DM_Tabelas.ZQVenda.FieldByName('idvenda').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('quadralote').Value := DM_Tabelas.ZQVenda.FieldByName('quadra').Value+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value := 'Empreendimento: '+DM_Tabelas.ZQVenda.FieldByName('nometoeam').Value+chr(13);
      DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value := DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value+'Quadra: '+DM_Tabelas.ZQVenda.FieldByName('quadra').Value+chr(13);
      DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value := DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value+'Lote: '+DM_Tabelas.ZQVenda.FieldByName('lote').Value;
      DM_Tabelas.ZQRecebimento.Post;
      DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
      if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
        DM_tabelas.ZQCheque.Insert;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('Banco').Value := DM_Tabelas.CDSPagEntradabanco.Value;
        DM_Tabelas.ZQCheque.FieldByName('agencia').Value := DM_Tabelas.CDSPagEntradaagencia.Value;
        DM_Tabelas.ZQCheque.FieldByName('numero').Value := DM_Tabelas.CDSPagEntradanumero.Value;
        DM_Tabelas.ZQCheque.FieldByName('conta').Value := DM_Tabelas.CDSPagEntradaconta.Value;
        DM_Tabelas.ZQCheque.FieldByName('Dono').Value := DM_Tabelas.CDSPagEntradanomedono.Value;
        DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := DM_Tabelas.CDSPagEntradadocumento.Value;
        DM_Tabelas.ZQCheque.FieldByName('valor').Value := DM_Tabelas.CDSPagEntradavalor.Value;
        DM_Tabelas.ZQCheque.FieldByName('emissao').Value := Date;
        DM_Tabelas.ZQCheque.FieldByName('deposito').Value := DM_Tabelas.CDSPagEntradavencimento.Value;
        DM_Tabelas.ZQCheque.Post;
      end;
      DM_Tabelas.CDSPagEntrada.Next;
    end;
    DM_Tabelas.CDSPagParcela.First;
    while not DM_Tabelas.cdSPagParcela.Eof do Begin
      DM_Tabelas.ZQRecebimento.Insert;
      DM_Tabelas.ZQRecebimento.FieldByName('idloteamento').Value := DM_Tabelas.ZQVenda.FieldByName('codloteamento').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('documento').Value := DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-P- '+CBPeriodoParc.Text;
      DM_Tabelas.ZQRecebimento.FieldByName('TipDoc').Value := DM_Tabelas.cdSPagParcelatipodoc.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('cliente').Value := DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := date;
      DM_Tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := DM_Tabelas.cdSPagParcelavencimento.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('ordem').Value := inttostr(varnumordemP)+'-'+DM_Tabelas.cdSPagParcelaordem.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('saldo').Value := DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').Value/100;
      DM_Tabelas.ZQRecebimento.FieldByName('Valor').Value := DM_Tabelas.cdSPagParcelavalor.Value*DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').Value/100;
      DM_Tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNEEntrada.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('contabil').Value := DM_Tabelas.cdSPagParcelacontabil.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('origem').Value := 'V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text;
      DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := 'R';
      DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := varnumordemP;
      DM_Tabelas.ZQRecebimento.FieldByName('quadralote').Value := DM_Tabelas.ZQVenda.FieldByName('quadra').Value+'-'+DM_Tabelas.ZQVenda.FieldByName('lote').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value := DM_Tabelas.ZQVenda.FieldByName('idvenda').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value := 'Empreendimento: '+DM_Tabelas.ZQVenda.FieldByName('nometoeam').Value+chr(13);
      DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value := DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value+'Quadra: '+DM_Tabelas.ZQVenda.FieldByName('quadra').Value+chr(13);
      DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value := DM_Tabelas.ZQRecebimento.FieldByName('Observ').Value+'Lote: '+DM_Tabelas.ZQVenda.FieldByName('lote').Value;
      DM_Tabelas.ZQRecebimento.Post;
      DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagParcelatipodoc.Value,[]);
      if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
        DM_tabelas.ZQCheque.Insert;
        DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
        DM_Tabelas.ZQCheque.FieldByName('Banco').Value := DM_Tabelas.cdSPagParcelabanco.Value;
        DM_Tabelas.ZQCheque.FieldByName('agencia').Value := DM_Tabelas.cdSPagParcelaagencia.Value;
        DM_Tabelas.ZQCheque.FieldByName('numero').Value := DM_Tabelas.CDSPagParcelanumero.Value;
        DM_Tabelas.ZQCheque.FieldByName('Dono').Value := DM_Tabelas.cdSPagParcelanomedono.Value;
        DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := DM_Tabelas.cdSPagParceladocumento.Value;
        DM_Tabelas.ZQCheque.FieldByName('valor').Value := DM_Tabelas.cdSPagParcelavalor.Value;
        DM_Tabelas.ZQCheque.FieldByName('emissao').Value := Date;
        DM_Tabelas.ZQCheque.FieldByName('deposito').Value := DM_Tabelas.cdSPagParcelavencimento.Value;
        DM_Tabelas.ZQCheque.Post;
      end;
      DM_Tabelas.cdSPagParcela.Next;
    end;
     DM_Tabelas.ZQincorp_loteame.Next;
  end;
  DesativaCampos;
  DBGVenda.SetFocus;
  varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').Value;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  if Frm_QuadroResumo=nil then
    Frm_QuadroResumo:=TFrm_QuadroResumo.Create(Application);
  Frm_QuadroResumo.RLReport1.Preview;
  DM_Tabelas.ZQVenda.Filtered := False;
end;

procedure TFrm_Venda.BtCancelarClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Cancel;
  DesativaCampos;
  DBGVenda.SetFocus;
end;

procedure TFrm_Venda.BtProcessarClick(Sender: TObject);
begin
// porcesso
end;

procedure TFrm_Venda.BtIncluirClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Insert;
  EloteTemp.Text := ''; 
  AtivaCampos;
  Pag_Venda.PageIndex := 0;
  DM_TAbelas.CDSCompradorTemp.Close;
  DM_Tabelas.CDSCompradorTemp.CreateDataSet;
  DM_Tabelas.CDSVendedorTemp.Close;
  DM_Tabelas.CDSVendedorTemp.CreateDataSet;
  DM_Tabelas.CDSPagEntrada.Close;
  DM_Tabelas.CDSPagEntrada.CreateDataSet;
  DM_Tabelas.CDSPagParcela.Close;
  DM_Tabelas.CDSPagParcela.CreateDataSet;
  DM_Tabelas.ZQVenda.FieldByName('datavenda').Value := date;
  XNEEntrada.Value := 0;
  XNQtdeEnt.Value := 0;
  XNQtdeParc.Value := 0;
  XVencimentoEnt.DateValue := date;
  XVencimentoParc.DateValue := date;
  CBPeriodoEnt.ItemIndex := 3;
  CBPeriodoParc.ItemIndex := 3;
  DBEDAta.SetFocus;
end;

procedure TFrm_Venda.BtEditarClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Edit;
  EloteTemp.Text := DM_tabelas.ZQVenda.FieldByName('imovel').Text;
  AtivaCampos;
  if Pag_Venda.PageIndex = 0 Then
    DBEDAta.SetFocus;

end;

procedure TFrm_Venda.BTExcluirClick(Sender: TObject);
begin
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento where documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento ');
  DM_Tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQVenda.Delete;
end;

procedure TFrm_Venda.BtRelatorioClick(Sender: TObject);
begin
  JvXPBar1.Visible := True;
  JvXPBar1.SetFocus;
end;

procedure TFrm_Venda.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Venda.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Venda.DXBAchaLoteClick(Sender: TObject);
begin
  Frm_AchaLoteVenda.showmodal;
  Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Venda.ECompradorExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    if (DM_Tabelas.CDSCompradorTemp.RecordCount=0) or (not empty(Ecomprador.Text)) Then Begin
      if AchaCLiente(200, 10,Ecomprador.Text) Then Begin
        DM_Tabelas.CDSCompradorTemp.Insert;
        DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').Value;
        DM_Tabelas.CDSCompradorTempordem.Value := DM_Tabelas.CDSCompradorTemp.RecordCount;
        DM_Tabelas.CDSCompradorTemp.Post;
        Atualiza_comprador;
      end;
      Ecomprador.Text := '';
      Ecomprador.SetFocus;
    end;
end;

procedure TFrm_Venda.CompradorEnter(Sender: TObject);
begin
  Ecomprador.SetFocus;
end;

procedure TFrm_Venda.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key=46) and (DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit]) Then
    DM_Tabelas.CDSCompradorTemp.Delete;
  Atualiza_comprador;
end;

procedure TFrm_Venda.CorretorEnter(Sender: TObject);
begin
  Ecorretor.SetFocus;
end;

procedure TFrm_Venda.EcorretorExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    if (DM_Tabelas.CDSVendedorTemp.RecordCount=0) or (not empty(Ecorretor.Text)) Then Begin
      if AchaCorretor(200, 10,Ecorretor.Text) Then Begin
        DM_Tabelas.CDSVendedorTemp.Insert;
        DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_TAbelas.ZQAchaCorretor.FieldByName('idcorretor').Value;
        DM_TAbelas.CDSVendedorTempordem.Value := DM_Tabelas.CDSVendedorTemp.RecordCount;
        DM_Tabelas.CDSVendedorTemp.Post;
      end;
      Ecorretor.Text := '';
      Ecorretor.SetFocus;
    end;
end;

procedure TFrm_Venda.CorretorExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQVenda.State in [DsInsert]) Then Begin
    Pag_Venda.PageIndex := 1;
    DBEVrvenda.SetFocus;
  end;
end;

procedure TFrm_Venda.DBGVendaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQVenda.FieldByName('idvenda').Text Then Begin
    DBGVenda.Canvas.Brush.Color :=$006CFFFF;
    DBGVenda.Canvas.Font.Color := $00A80000;
    DBGVenda.Canvas.Font.Style := [FsBold];
  end;
  DBGVenda.DefaultDrawDataCell(Rect, DBGVenda.columns[datacol].field, State);

end;

procedure TFrm_Venda.DBGVendaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  Botoes;
end;

procedure TFrm_Venda.DBGVendaMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Botoes;
end;

procedure TFrm_Venda.DBGVendaEnter(Sender: TObject);
begin
  Botoes;
end;

procedure TFrm_Venda.DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Atualiza_comprador;
end;

procedure TFrm_Venda.DBGrid2Enter(Sender: TObject);
begin
  Atualiza_comprador;
end;

procedure TFrm_Venda.DBEEntradaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEOrdem.Text = DM_Tabelas.CDSPagEntradaordem.Text Then Begin
    DBEEntrada.Canvas.Brush.Color :=$006CFFFF;
    DBEEntrada.Canvas.Font.Color := $00A80000;
    DBEEntrada.Canvas.Font.Style := [FsBold];
  end;
  DBEEntrada.DefaultDrawDataCell(Rect, DBEEntrada.columns[datacol].field, State);

end;

procedure TFrm_Venda.DBEOrdemChange(Sender: TObject);
begin
  Label17.Caption := copy(DM_Tabelas.CDSPagEntradaordem.Value,1,3);
end;

procedure TFrm_Venda.GBEntradaEnter(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    XNEEntrada.SetFocus;
end;

procedure TFrm_Venda.CBPeriodoEntExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma : double;
  varvenci : Tdatetime;
  vartipo : string;
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then Begin
    if empty(CBPeriodoEnt.Text) Then Begin
      CBPeriodoEnt.SetFocus;
      Exit;
    end;
    varvenci :=  XVencimentoEnt.DateValue;
    DM_Tabelas.CDSPagEntrada.Close;
    DM_Tabelas.CDSPagEntrada.CreateDataSet;
    varfinal := strtoint(floattostr(XNQtdeEnt.Value));
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
        0 : vardias := 1;
        1 : vardias := 7;
        2 : vardias := 15;
        3 : vardias := 1;
        4 : vardias := 2;
        5 : vardias := 3;
        6 : vardias := 4;
        7 : vardias := 6;
        8 : vardias := 12;
        9 : vardias := 24;
      end;
      if CBPeriodoEnt.ItemIndex<=2 Then
        vartipo := 'N';
    end;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value,[]);
    varsoma := 0;
    for varvezes := 1 to varfinal do Begin
      DM_Tabelas.CDSPagEntrada.Insert;
      DM_Tabelas.CDSPagEntradaordem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
      DM_Tabelas.CDSPagEntradavalor.Value := truncar(XNEEntrada.Value/varfinal,2);
      DM_Tabelas.CDSPagEntradacontabil.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value;
      DM_Tabelas.CDSPagEntradatipodoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').Value;
      if varvezes = 1 Then
         DM_Tabelas.CDSPagEntradavencimento.Value := XVencimentoEnt.DateValue
      else Begin
        if vartipo = 'S' then
          DM_Tabelas.CDSPagEntradavencimento.Value := IncMonth(XVencimentoEnt.DateValue,(varvezes*vardias)-(vardias))
        else
          DM_Tabelas.CDSPagEntradavencimento.Value := varvenci+vardias;
        varvenci := DM_Tabelas.CDSPagEntradavencimento.Value;
      end;
      DM_Tabelas.CDSPagEntrada.Post;
      varsoma := varsoma + DM_Tabelas.CDSPagEntradavalor.Value;
    end;
    DM_Tabelas.CDSPagEntrada.First;
    DM_Tabelas.CDSPagEntrada.Edit;
    DM_Tabelas.CDSPagEntradavalor.Value := DM_Tabelas.CDSPagEntradavalor.Value+(XNEEntrada.Value-varsoma);
    DM_Tabelas.CDSPagEntrada.Post;
  end;
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.atualiza_tela_venda_entrada;
Begin
  EContabilEntrada.Text := DM_Tabelas.CDSPagEntradacontabildescri.Value;
  DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
  if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
    Label20.Visible := True;
    Label21.Visible := True;
    Label27.Visible := True;
    Label32.Visible := True;
    Label33.Visible := True;
    Label51.Visible := True;
    DBEBanco.Visible := True;
    DBEAgencia.Visible := True;
    DBEConta.Visible := True;
    DBEDocumento.Visible := True;
    DBENumeroEntrada.Visible := True;
    DBEDonoEntrada.Visible := True;
  end
  else Begin
    Label20.Visible := False;
    Label21.Visible := False;
    Label27.Visible := False;
    Label32.Visible := False;
    Label33.Visible := False;
    Label51.Visible := False;
    DBEBanco.Visible := False;
    DBEAgencia.Visible := False;
    DBEConta.Visible := False;
    DBEDocumento.Visible := False;
    DBENumeroEntrada.Visible := False;
    DBEDonoEntrada.Visible := False;
  end;
End;

procedure TFrm_Venda.atualiza_tela_venda_parcela;
Begin
  EContabilEntradaParc.Text := DM_Tabelas.CDSPagParcelacontabildescri.Value;
  DM_Tabelas.ZQTipoDoc.Locate('tipodoc',DM_Tabelas.CDSPagParcelatipodoc.Value,[]);
  if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
    Label45.Visible := True;
    Label46.Visible := True;
    Label47.Visible := True;
    Label48.Visible := True;
    Label49.Visible := True;
    Label52.Visible := True;
    DBEBancoParc.Visible := True;
    DBEAgendaParc.Visible := True;
    DBEContaParc.Visible := True;
    DBEDocumentoParc.Visible := True;
    DBENumeroParc.Visible := True;
    DBEDonoParcela.Visible := True;
  end
  else Begin
    Label45.Visible := False;
    Label46.Visible := False;
    Label47.Visible := False;
    Label48.Visible := False;
    Label49.Visible := False;
    Label52.Visible := False;
    DBEBancoParc.Visible := False;
    DBEAgendaParc.Visible := False;
    DBEContaParc.Visible := False;
    DBEDocumentoParc.Visible := False;
    DBENumeroParc.Visible := False;
    DBEDonoParcela.Visible := False;
  end;
End;

procedure TFrm_Venda.atualizatebelaspagamento;
var
  varsomaE, varsomaP, VE, VP : double;
Begin
  DM_Tabelas.ZQRecebimento_temp.SQL.Clear;
  DM_Tabelas.ZQRecebimento_temp.SQL.Add('Select *, sum(valor) as somavalor from recebimento');
  DM_Tabelas.ZQRecebimento_temp.SQL.Add(' Where origem='+quotedstr('V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text));
  DM_Tabelas.ZQRecebimento_temp.SQL.Add(' group by idrecebimento');
  DM_Tabelas.ZQRecebimento_temp.Open;
  DM_Tabelas.CDSPagEntrada.Close;
  DM_Tabelas.CDSPagEntrada.CreateDataSet;
  DM_Tabelas.CDSPagParcela.Close;
  DM_Tabelas.CDSPagParcela.CreateDataSet;
  varsomaE := 0;
  varsomaP := 0;
  Ve := 0;
  Vp := 0;
  XVencimentoEnt.DateText := '';
  XVencimentoParc.DateText := '';
  while not DM_Tabelas.ZQRecebimento_temp.Eof do begin
    IF pos('-E-',DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').Value)>0 Then Begin
      if Copy(XVencimentoEnt.DateText,1,2)='00'  Then Begin
        XVencimentoEnt.DateValue := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').Value;
        CBPeriodoEnt.Text := copy(DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').Value,pos(' ',DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').Value)+1,20);
      end;
      DM_Tabelas.CDSPagEntrada.Insert;
      DM_Tabelas.CDSPagEntradacontabil.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('contabil').Value;
      DM_Tabelas.CDSPagEntradavencimento.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').Value;
      DM_Tabelas.CDSPagEntradaordem.Value := copy(DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').Value,pos('-',DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').Value)+1,7);
      DM_Tabelas.CDSPagEntradavalor.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').Value;
      DM_Tabelas.CDSPagEntradatipodoc.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('TipDoc').Value;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
        DM_Tabelas.ZQRecebimento.Locate('idrecebimento',DM_Tabelas.ZQRecebimento_temp.FieldByName('idrecebimento').Value,[]);
        DM_Tabelas.CDSPagEntradabanco.Value := DM_Tabelas.ZQCheque.FieldByName('Banco').Value;
        DM_Tabelas.CDSPagEntradaagencia.Value := DM_Tabelas.ZQCheque.FieldByName('agencia').Value;
        DM_Tabelas.CDSPagEntradaconta.Value := DM_Tabelas.ZQCheque.FieldByName('conta').Value;
        DM_Tabelas.CDSPagEntradanumero.Value := DM_Tabelas.ZQCheque.FieldByName('numero').Value;
        DM_Tabelas.CDSPagEntradadocumento.Value := DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value;
        DM_Tabelas.CDSPagEntradanomedono.Value := DM_Tabelas.ZQCheque.FieldByName('Dono').Value;
      end;
      DM_Tabelas.CDSPagEntrada.Post;
      varsomaE := varsomaE + DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').Value;
      ve := ve+1;
    end
    else Begin
      if Copy(XVencimentoParc.DateText,1,2)='00'  Then Begin
        XVencimentoParc.DateValue := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').Value;
        CBPeriodoParc.Text := copy(DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').Value,pos(' ',DM_Tabelas.ZQRecebimento_temp.FieldByName('documento').Value)+1,20);
      end;
      DM_Tabelas.CDSPagParcela.Insert;
      DM_Tabelas.CDSPagParcelacontabil.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('contabil').Value;
      DM_Tabelas.CDSPagParcelavencimento.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('Dt_Vencimento').Value;
      DM_Tabelas.CDSPagParcelaordem.Value := copy(DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').Value,pos('-',DM_TAbelas.ZQRecebimento_temp.FieldByName('ordem').Value)+1,7);
      DM_Tabelas.CDSPagParcelavalor.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').Value;
      DM_Tabelas.CDSPagParcelatipodoc.Value := DM_Tabelas.ZQRecebimento_temp.FieldByName('TipDoc').Value;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',DM_Tabelas.CDSPagEntradatipodoc.Value,[]);
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
        DM_Tabelas.ZQRecebimento.Locate('idrecebimento',DM_Tabelas.ZQRecebimento_temp.FieldByName('idrecebimento').Value,[]);
        DM_Tabelas.cdSPagParcelabanco.Value := DM_Tabelas.ZQCheque.FieldByName('Banco').Value;
        DM_Tabelas.cdSPagParcelaagencia.Value := DM_Tabelas.ZQCheque.FieldByName('agencia').Value;
        DM_Tabelas.cdSPagParcelaconta.Value := DM_Tabelas.ZQCheque.FieldByName('conta').Value;
        DM_Tabelas.cdSPagParcelanumero.Value := DM_Tabelas.ZQCheque.FieldByName('numero').Value;
        DM_Tabelas.cdSPagParceladocumento.Value := DM_Tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value;
        DM_Tabelas.cdSPagParcelanomedono.Value := DM_Tabelas.ZQCheque.FieldByName('Dono').Value;
      end;
      DM_Tabelas.CDSPagParcela.Post;
      varsomaP := varsomaP + DM_Tabelas.ZQRecebimento_temp.FieldByName('somavalor').Value;
      vp := vp +1;
    end;
    DM_Tabelas.ZQRecebimento_temp.Next;
  end;
  DM_Tabelas.CDSPagEntrada.First;
  DM_Tabelas.CDSPagParcela.First;
  XNEEntrada.Value := varsomaE;
  XNParcela.Value := varsomaP;
  XNQtdeEnt.Value :=  ve;
  XNQtdeParc.Value := vp;
  atualiza_tela_venda_entrada;
  atualiza_tela_venda_parcela;
end;



procedure TFrm_Venda.DBEEntradaEnter(Sender: TObject);
begin
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.DBEEntradaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.EContabilEntradaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsEdit,DsInsert] Then Begin
    if (not empty(EContabilEntrada.Text)) and (AchaPlanoDeConta(450, 30, 'S', EContabilEntrada.Text, '')) Then Begin
      EContabilEntrada.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').Value;
      DM_Tabelas.CDSPagEntrada.Edit;
      DM_Tabelas.CDSPagEntradacontabil.Value :=  DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
      DM_tabelas.CDSPagEntradatipodoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').Value;
      DM_Tabelas.CDSPagEntrada.Post;
    end
    else
      EContabilEntrada.SetFocus;
  end;
end;

procedure TFrm_Venda.EContabilEntradaEnter(Sender: TObject);
begin
  EContabilEntrada.Text:= DM_Tabelas.CDSPagEntradacontabildescri.Value;
  atualiza_tela_venda_entrada;

end;

procedure TFrm_Venda.DBEValorExit(Sender: TObject);
var
  varreg, varconta : integer;
  varsoma, resto : double;
begin
  varreg := DM_Tabelas.CDSPagEntrada.RecNo;
  if varreg < XNQtdeEnt.Value Then Begin
    DM_Tabelas.CDSPagEntrada.First;
    varsoma := 0;
    for varconta := 1 to varreg do Begin
      DM_Tabelas.CDSPagEntrada.RecNo := varconta;
      varsoma := varsoma + DM_Tabelas.CDSPagEntradavalor.Value;
    end;
    resto := XNEEntrada.Value - varsoma;
    varsoma := 0;
    for varconta := varreg+1 to DM_Tabelas.CDSPagEntrada.RecordCount do begin
      DM_Tabelas.CDSPagEntrada.RecNo := varconta;
      DM_Tabelas.CDSPagEntrada.Edit;
      DM_Tabelas.CDSPagEntradavalor.Value := truncar(resto/(XNQtdeEnt.Value-varreg),2);
      varsoma:=varsoma+DM_Tabelas.CDSPagEntradavalor.Value;
    end;
    DM_Tabelas.CDSPagEntrada.RecNo := varreg+1;
    DM_Tabelas.CDSPagEntrada.Edit;
    DM_Tabelas.CDSPagEntradavalor.Value := DM_Tabelas.CDSPagEntradavalor.Value+(resto-varsoma);
    DM_Tabelas.CDSPagEntrada.Post;
    DM_Tabelas.CDSPagEntrada.RecNo := varreg;
    if (DM_Tabelas.CDSPagEntrada.RecNo<XNQtdeEnt.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'N') Then Begin
      DM_Tabelas.CDSPagEntrada.Next;
      EContabilEntrada.SetFocus;
    end
    else if (DM_Tabelas.CDSPagEntrada.RecNo<XNQtdeEnt.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S') Then Begin
      DBEBanco.SetFocus;
    end;
  end;
end;

procedure TFrm_Venda.DBEValorEnter(Sender: TObject);
begin
  if DM_Tabelas.CDSPagEntrada.RecNo=DM_Tabelas.CDSPagEntrada.RecordCount Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Venda.DBEDocumentoExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then 
    if not Verif_doc(DBEDocumento.text,True) Then Begin
      DBEDocumento.SetFocus;
      exit;
    end;
end;

procedure TFrm_Venda.XVencimentoEntExit(Sender: TObject);
begin
  if XVencimentoEnt.DateValue<=0 Then
    XVencimentoEnt.SetFocus;
end;

procedure TFrm_Venda.DBCBDocVendaExit(Sender: TObject);
begin
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.XNEEntradaExit(Sender: TObject);
begin
  XNParcela.Value := DBEVrvenda.Value - XNEEntrada.Value;
  if XNEEntrada.Value = 0 Then
    GBParcelas.SetFocus;
end;

procedure TFrm_Venda.GBParcelasEnter(Sender: TObject);
begin
  if XNParcela.Value>0 Then
    XNQtdeParc.SetFocus
  else
    BtGravar.SetFocus;
end;

procedure TFrm_Venda.CBPeriodoParcExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma : double;
  varvenci : Tdatetime;
  vartipo : string;
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then Begin
    if empty(CBPeriodoParc.Text) Then Begin
      CBPeriodoParc.SetFocus;
      Exit;
    end;
    varvenci :=  XVencimentoParc.DateValue;
    DM_Tabelas.CDSPagParcela.Close;
    DM_Tabelas.CDSPagParcela.CreateDataSet;
    varfinal := strtoint(floattostr(XNQtdeParc.Value));
    if StrIsNumero(CBPeriodoParc.Text) Then Begin
      vardias := strtoint(CBPeriodoParc.Text);
      vartipo := 'N';
    end
    else Begin
      if CBPeriodoParc.ItemIndex<0 Then Begin
        CBPeriodoParc.SetFocus;
        Exit;
      end;
      vartipo := 'S';
      case CBPeriodoParc.ItemIndex of
        0 : vardias := 1;
        1 : vardias := 7;
        2 : vardias := 15;
        3 : vardias := 1;
        4 : vardias := 2;
        5 : vardias := 3;
        6 : vardias := 4;
        7 : vardias := 6;
        8 : vardias := 12;
        9 : vardias := 24;
      end;
      if CBPeriodoParc.ItemIndex<=2 Then
        vartipo := 'N';
    end;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value,[]);
    varsoma := 0;
    for varvezes := 1 to varfinal do Begin
      DM_Tabelas.CDSPagParcela.Insert;
      DM_Tabelas.CDSPagParcelaordem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
      DM_Tabelas.CDSPagParcelavalor.Value := truncar(XNParcela.Value/varfinal,2);
      DM_Tabelas.CDSPagParcelacontabil.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value;
      DM_Tabelas.CDSPagParcelatipodoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').Value;
      if varvezes = 1 Then
         DM_Tabelas.CDSPagParcelavencimento.Value := XVencimentoParc.DateValue
      else Begin
        if vartipo = 'S' then
          DM_Tabelas.CDSPagParcelavencimento.Value := IncMonth(XVencimentoParc.DateValue,(varvezes*vardias)-(vardias))
        else
          DM_Tabelas.CDSPagParcelavencimento.Value := varvenci+vardias;
        varvenci := DM_Tabelas.CDSPagParcelavencimento.Value;
      end;
      DM_Tabelas.CDSPagParcela.Post;
      varsoma := varsoma + DM_Tabelas.CDSPagParcelavalor.Value;
    end;
    DM_Tabelas.CDSPagParcela.First;
    DM_Tabelas.CDSPagParcela.Edit;
    DM_Tabelas.CDSPagParcelavalor.Value := DM_Tabelas.CDSPagParcelavalor.Value+(XNParcela.Value-varsoma);
    DM_Tabelas.CDSPagParcela.Post;
  end;
  if Pag_Venda.PageIndex = 1 Then
    atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.XVencimentoParcExit(Sender: TObject);
begin
  if XVencimentoParc.DateValue<=0 Then
    XVencimentoParc.SetFocus;
end;

procedure TFrm_Venda.DBEDocumentoParcExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    if not Verif_doc(DBEDocumentoParc.text,True) Then Begin
      DBEDocumentoParc.SetFocus;
      exit;
    end;
end;

procedure TFrm_Venda.DBGPartelasDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEOrdemParc.Text = DM_Tabelas.CDSPagParcelaordem.Text Then Begin
    DBGPartelas.Canvas.Brush.Color :=$006CFFFF;
    DBGPartelas.Canvas.Font.Color := $00A80000;
    DBGPartelas.Canvas.Font.Style := [FsBold];
  end;
  DBGPartelas.DefaultDrawDataCell(Rect, DBGPartelas.columns[datacol].field, State);

end;

procedure TFrm_Venda.DBEVrvendaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.FieldByName('valorvenda').Value<DM_Tabelas.ZQVenda.FieldByName('valorvenal').Value Then
    Showmessage('O valor de venda � menor que o valor venal...');
end;

procedure TFrm_Venda.DBCBDocVendaParcExit(Sender: TObject);
begin
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.EContabilEntradaParcEnter(Sender: TObject);
begin
  EContabilEntradaParc.Text:= DM_Tabelas.CDSPagparcelacontabildescri.Value;
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.EContabilEntradaParcExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsEdit,DsInsert] Then Begin
    if (not empty(EContabilEntradaParc.Text)) and (AchaPlanoDeConta(450, 30, 'S', EContabilEntradaParc.Text, '')) Then Begin
      EContabilEntradaParc.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' | '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').Value;
      DM_Tabelas.cdSPagParcela.Edit;
      DM_Tabelas.cdSPagParcelacontabil.Value :=  DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
      DM_tabelas.cdSPagParcelatipodoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').Value;
      DM_Tabelas.cdSPagParcela.Post;
    end
    else
      EContabilEntradaParc.SetFocus;
  end;

end;

procedure TFrm_Venda.DBEOrdemParcChange(Sender: TObject);
begin
  Label44.Caption := copy(DM_Tabelas.CDSPagparcelaordem.Value,1,3);
end;

procedure TFrm_Venda.DBEValorParcEnter(Sender: TObject);
begin
  if DM_Tabelas.CDSPagparcela.RecNo=DM_Tabelas.CDSPagparcela.RecordCount Then
    Perform(WM_NEXTDLGCTL, 0, 0);

end;

procedure TFrm_Venda.DBEValorParcExit(Sender: TObject);
var
  varreg, varconta : integer;
  varsoma, resto : double;
begin
  varreg := DM_Tabelas.cdSPagParcela.RecNo;
  if varreg < XNQtdeParc.Value Then Begin
    DM_Tabelas.cdSPagParcela.First;
    varsoma := 0;
    for varconta := 1 to varreg do Begin
      DM_Tabelas.cdSPagParcela.RecNo := varconta;
      varsoma := varsoma + DM_Tabelas.cdSPagParcelavalor.Value;
    end;
    resto := XNParcela.Value - varsoma;
    varsoma := 0;
    for varconta := varreg+1 to DM_Tabelas.cdSPagParcela.RecordCount do begin
      DM_Tabelas.cdSPagParcela.RecNo := varconta;
      DM_Tabelas.cdSPagParcela.Edit;
      DM_Tabelas.cdSPagParcelavalor.Value := truncar(resto/(XNQtdeParc.Value-varreg),2);
      varsoma:=varsoma+DM_Tabelas.cdSPagParcelavalor.Value;
    end;
    DM_Tabelas.cdSPagParcela.RecNo := varreg+1;
    DM_Tabelas.cdSPagParcela.Edit;
    DM_Tabelas.cdSPagParcelavalor.Value := DM_Tabelas.cdSPagParcelavalor.Value+(resto-varsoma);
    DM_Tabelas.cdSPagParcela.Post;
    DM_Tabelas.cdSPagParcela.RecNo := varreg;
    if (DM_Tabelas.cdSPagParcela.RecNo<XNQtdeParc.Value) and (DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'N') Then Begin
      DM_Tabelas.cdSPagParcela.Next;
      EContabilEntradaParc.SetFocus;
    end;
  end;
end;

procedure TFrm_Venda.DBEEntradaMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  atualiza_tela_venda_entrada;
end;

procedure TFrm_Venda.DBGPartelasKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.DBGPartelasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.DBGPartelasEnter(Sender: TObject);
begin
  atualiza_tela_venda_parcela;
end;

procedure TFrm_Venda.DBEDonoEntradaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    if (DM_Tabelas.CDSPagEntrada.RecNo<XNQtdeEnt.Value) Then Begin
      DM_Tabelas.CDSPagEntrada.Next;
      EContabilEntrada.SetFocus;
    end;

end;

procedure TFrm_Venda.DBEDonoParcelaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] then
    if (DM_Tabelas.CDSPagParcela.RecNo<XNQtdeParc.Value) Then Begin
      DM_Tabelas.CDSPagParcela.Next;
      EContabilEntradaParc.SetFocus;
    end;
end;

procedure TFrm_Venda.JvXPBar1Items0Click(Sender: TObject);
var
  varvenda : integer;
begin
  varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').Value;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Open;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Open;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  DM_Tabelas.ZQCompr_Dados.First;
  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').Value,[]);
  Frm_Loteamento.atualiza_lotes;
  DM_TAbelas.CDSAss1.Close;
  DM_TAbelas.CDSAss1.CreateDataSet;
  DM_TAbelas.CDSAss1.Insert;
  DM_TAbelas.CDSAss1assinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('nomeloteamento').Value;
  DM_Tabelas.CDSAss1cargo.Value :=' Vendedor';
  DM_Tabelas.CDSAss1responsavel.Value := 'PP: '+DM_Tabelas.CDSIncorpnomeparti.Value;
  DM_TAbelas.CDSAss1.Post;
  DM_TAbelas.CDSAss.Close;
  DM_TAbelas.CDSAss.CreateDataSet;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').Value) then Begin
    DM_TAbelas.CDSAss.Insert;
    DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha1').Value;
    DM_Tabelas.CDSAsscargo.Value :='Testemunha';
    DM_TAbelas.CDSAss.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').Value) then Begin
    DM_TAbelas.CDSAss.Insert;
    DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha2').Value;
    DM_Tabelas.CDSAsscargo.Value :='Testemunha';
    DM_TAbelas.CDSAss.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').Value) then Begin
    DM_TAbelas.CDSAss.Insert;
    DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha3').Value;
    DM_Tabelas.CDSAsscargo.Value :='Testemunha';
    DM_TAbelas.CDSAss.Post;
  end;
  if not empty(DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').Value) then Begin
    DM_TAbelas.CDSAss.Insert;
    DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQLoteamento.FieldByName('testemunha4').Value;
    DM_Tabelas.CDSAsscargo.Value :='Testemunha';
    DM_TAbelas.CDSAss.Post;
  end;
  while not DM_Tabelas.ZQCompr_Dados.Eof do begin
    DM_TAbelas.CDSAss.Insert;
    DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Dados.FieldByName('nome_parte').Value;
    DM_Tabelas.CDSAsscargo.Value :='Comprador(a)';
    DM_TAbelas.CDSAss.Post;
    while not DM_TAbelas.ZQCompr_Resp_Dados.Eof do begin
      DM_TAbelas.CDSAss.Insert;
      DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_Resp_Dados.FieldByName('nomerepres').Value;
      DM_Tabelas.CDSAsscargo.Value :='Respons�vel';
      DM_TAbelas.CDSAss.Post;
      DM_TAbelas.ZQCompr_Resp_Dados.Next;
    end;
    while not DM_TAbelas.ZQCompr_Conjuge.Eof do begin
      if DM_Tabelas.ZQCompr_conjuge.FieldByName('assina').Value = 'S' Then Begin
        DM_TAbelas.CDSAss.Insert;
        DM_TAbelas.CDSAssassinante.Value :=DM_TAbelas.ZQCompr_conjuge.FieldByName('nome').Value;
        DM_Tabelas.CDSAsscargo.Value :='Conjuge';
        DM_TAbelas.CDSAss.Post;
      end;
      DM_TAbelas.ZQCompr_Conjuge.Next;
    end;
    DM_Tabelas.ZQCompr_Dados.Next;
  end;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_TAbelas.ZQCompr_Resp_Dados.Refresh;
  DM_TAbelas.ZQCompr_Conjuge.Refresh;
  if Frm_QuadroResumo=nil then
    Frm_QuadroResumo:=TFrm_QuadroResumo.Create(Application);
  Frm_QuadroResumo.RLReport1.Preview;
  DM_Tabelas.ZQVenda.Filtered := False;
end;

procedure TFrm_Venda.JvXPBar1Items1Click(Sender: TObject);
begin
  Frm_RelVenda.Top := Frm_Venda.Top+100;
  Frm_RelVenda.Left := Frm_Venda.Left-4+(Frm_Venda.Width-Frm_RelVenda.Width);
  Frm_RelVenda.Showmodal;
end;

procedure TFrm_Venda.JvXPBar1Exit(Sender: TObject);
begin
    JvXPBar1.Visible := False;

end;

procedure TFrm_Venda.BTPesquisarClick(Sender: TObject);
begin
  Frm_AchaVenda.Top := Frm_Venda.Top+100;
  Frm_AchaVenda.Left := Frm_Venda.Left+(Frm_Venda.Width-Frm_AchaVenda.Width)  ;
  Frm_AchaVenda.showmodal;
  DBGVenda.SetFocus;
{  Frm_AchaLoteVendido.showmodal;
  Perform(WM_NEXTDLGCTL, 0, 0);
  DBGVenda.SetFocus;
 }


end;

procedure TFrm_Venda.Pag_VendaChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
begin
 // Atualiza;
end;


procedure TFrm_Venda.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
