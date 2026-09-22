unit trocar_lote;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XDBNum, XNum, StdCtrls, wwdbdatetimepicker, Mask, DBCtrls,
  Grids, DBGrids, ComCtrls, TabNotBk, ExtCtrls, XBanner,
  XDate, XDBEdit, DB, JvExControls, JvXPCore, JvXPBar, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, JvExComCtrls, JvDateTimePicker, XDBDate,
  FnpNumericEdit,Comobj;

type
  TFrm_Trocar = class(TForm)
    XBanner14: TXBanner;
    Label19: TLabel;
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTExcluir: TdxButtonArround;
    BtEditar: TdxButtonArround;
    BtIncluir: TdxButtonArround;
    BtCancelar: TdxButtonArround;
    BTPesquisar: TdxButtonArround;
    BtUltimo: TdxButtonArround;
    BtProximo: TdxButtonArround;
    BTAnterior: TdxButtonArround;
    BtPrimeiro: TdxButtonArround;
    BTFechar: TdxButtonArround;
    BtGravar: TdxButtonArround;
    Pag_Venda: TTabbedNotebook;
    XBanner1: TXBanner;
    DBGVenda: TDBGrid;
    DBECod: TDBEdit;
    Label1: TLabel;
    Label5: TLabel;
    Comprador: TGroupBox;
    Corretor: TGroupBox;
    XBanner2: TXBanner;
    XBanner3: TXBanner;
    DBGrid2: TDBGrid;
    Label14: TLabel;
    DBEImovel: TDBEdit;
    DXBAchaLote: TdxButtonArround;
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
    XDBNumEdit1: TXDBNumEdit;
    Label18: TLabel;
    EloteTemp: TEdit;
    ZQcomprador_cessao: TZQuery;
    DS_comprador_cessao: TDataSource;








    Zpart: TZQuery;
    Datazpart: TDataSource;




































    ZQatualizar: TZQuery;









































    Dataatualizar: TDataSource;
    ZQRecebimento_historico_cessao: TZQuery;
    DS_ZQRecebimento_historico: TDataSource;




































    ZQNossoNumero: TZQuery;
    DataZQNossoNumero: TDataSource;


































    DS_Remes_Receb: TDataSource;
    ZQRemes_Receb: TZQuery;










    bar1: TProgressBar;
    recebimento_historico_excluidos: TZQuery;






























    Datarecebimento_historico_excluidos: TDataSource;



    DBEDAta: TXDBDateEdit;
    Label6: TLabel;
    JDEntrada: TXDateEdit;






    Label78: TLabel;
    Epercentual: TFnpNumericEdit;
    BtRelatorio: TdxButtonArround;
    imovelant: TEdit;
    Edvenda: TEdit;
    edloteamento: TEdit;
    Panel1: TPanel;
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
    procedure BtCancelarClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
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
    procedure BTPesquisarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid2CellClick(Column: TColumn);
    procedure BtRelatorioClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Trocar: TFrm_Trocar;
  ncompr: integer;
implementation

uses tabelas, Funcoes, AchaLoteVenda, RelVenda, QuadroResumo, AchaVenda,
  Loteamento, uRuntimeFields;

{$R *.dfm}

type
  { A tela reutiliza a ZQVenda do data module. Recalcule os lookups no
    registro atual para que o primeiro e o ultimo registro ja sejam exibidos. }
  TDataSetCalcFieldsAccess = class(TDataSet);

procedure RecalcularVendaAtual;
begin
  if (DM_Tabelas = nil) or
     (not DM_Tabelas.ZQVenda.Active) or
     DM_Tabelas.ZQVenda.IsEmpty then
    Exit;
  TDataSetCalcFieldsAccess(DM_Tabelas.ZQVenda).GetCalcFields(
    DM_Tabelas.ZQVenda.ActiveBuffer);
end;


procedure TFrm_Trocar.FormShow(Sender: TObject);
begin
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
  RecalcularVendaAtual;
  ncompr:=0;
  
  JDEntrada.DateText:=datetostr(date);
  Pag_Venda.PageIndex := 0;
  DesativaCampos;
  DBGVenda.SetFocus;
end;

procedure TFrm_Trocar.botoes;
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

procedure TFrm_Trocar.Atualiza;
Begin
  DM_TAbelas.CDSCompradorTemp.Close;
  DM_Tabelas.CDSCompradorTemp.CreateDataSet;
  DM_Tabelas.ZQComprador.First;
  DM_Tabelas.ZQComprador.DisableControls;
  while not DM_Tabelas.ZQComprador.Eof do begin
    DM_Tabelas.CDSCompradorTemp.Insert;
    DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt;

    // 01/07/2013
    DM_Tabelas.CDSCompradorTempnomeparticipante.Value := DM_Tabelas.ZQComprador.FieldByName('nomecomprador').AsString;
    DM_Tabelas.CDSCompradorTemppercentual.Value := DM_Tabelas.ZQcomprador.FieldByName('percentual').AsFloat;


    if DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger=0 then
       DM_Tabelas.CDSCompradorTempMarcar.Value := False
    else
      DM_Tabelas.CDSCompradorTempMarcar.Value := True;

    DM_tabelas.CDSCompradorTemp.Post;
    dm_tabelas.ZQComprador.Next;
  end;
  DM_Tabelas.ZQComprador.EnableControls;
  ncompr:=DM_tabelas.CDSCompradorTemp.recordcount;
  DM_Tabelas.CDSVendedorTemp.Close;
  DM_Tabelas.CDSVendedorTemp.CreateDataSet;
  DM_Tabelas.ZQVendedor.First;
  DM_Tabelas.ZQVendedor.DisableControls;
  While not DM_Tabelas.ZQVendedor.Eof do Begin
    DM_Tabelas.CDSVendedorTemp.Insert;
    DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').AsLargeInt;
     DM_Tabelas.CDSVendedorTemp.Post;
    DM_Tabelas.ZQVendedor.Next;
  end;
  DM_Tabelas.ZQVendedor.EnableControls;
  Atualiza_comprador;
End;

procedure TFrm_Trocar.Atualiza_comprador;
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

procedure TFrm_Trocar.DesativaCampos;
Begin
  JDEntrada.Enabled:=false;
  DBEDAta.ReadOnly := True;
  Ecomprador.ReadOnly := True;
  Ecorretor.ReadOnly := True;
  DXBAchaLote.Enabled := False;
  BTPrimeiro.Enabled := true;
  BTAnterior.Enabled := true;
  BTProximo.Enabled := true;
  BtUltimo.Enabled := true;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtIncluir.Enabled := true;
  BtEditar.Enabled := true;
  BtExcluir.Enabled := true;
  BtPesquisar.Enabled := true;
  DBGVenda.Enabled := True;
end;

procedure TFrm_Trocar.AtivaCampos;
Begin
  JDEntrada.Enabled:=true;
  DBEDAta.ReadOnly := False;
  Ecomprador.ReadOnly := False;
  Ecorretor.ReadOnly := False;
  DXBAchaLote.Enabled := True;
  BTPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BTProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BtExcluir.Enabled := False;
  BtPesquisar.Enabled := False;
  DBGVenda.Enabled := False;
end;

procedure TFrm_Trocar.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.First;
  RecalcularVendaAtual;
  Botoes;
  DBGVenda.SetFocus;
end;


procedure TFrm_Trocar.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Prior;
  RecalcularVendaAtual;
  Botoes;
  DBGVenda.SetFocus;
end;

procedure TFrm_Trocar.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Next;
  RecalcularVendaAtual;
  Botoes;
  DBGVenda.SetFocus;
end;

procedure TFrm_Trocar.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Last;
  RecalcularVendaAtual;
  Botoes;
  DBGVenda.SetFocus;
end;

procedure TFrm_Trocar.BtCancelarClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Cancel;
  DesativaCampos;
  DBGVenda.SetFocus;
end;

procedure TFrm_Trocar.BtGravarClick(Sender: TObject);
var
  tm,tm2,varnumordemE, varnumordemP, varvezes, varvenda : integer;
  traco,dts,historico,historico2,entrada:String;
  bmLocal : TBookmark;

begin
  if simnao('Confirma a Operação ?','SIM')=false then
  begin
    mensagem('O Processo não será executado.');
    JDEntrada.SetFocus;
    exit;
  end;
  Panel1.Visible:=true;
  application.ProcessMessages;
  tm:=0;
  tm2:=0;
  entrada:='N';
  historico:='';
  historico2:='';
  if DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt=0 Then Begin
    mensagem('Deve ser indicado um lote para ser vendido...');
    Pag_Venda.PageIndex := 0;
    DXBAchaLote.SetFocus;
    exit;
  end;
//  if DM_Tabelas.ZQVendavalorvenda.Value<DM_Tabelas.ZQVendavalorvenal.Value Then Begin
//    Showmessage('O valor de VENDA não deve ser menor que o valor venal...');
//    Pag_Venda.PageIndex := 1;
//    DBEVrvenda.SetFocus;
//    exit;
//  end;
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
{  DM_Tabelas.ZQVenda.Post;
  bmLocal := DM_Tabelas.ZQVenda.GetBookMark;
  DM_Tabelas.ZQVenda.refresh;
  DM_Tabelas.ZQVenda.close;
  DM_Tabelas.ZQVenda.Open;
  DM_Tabelas.ZQVenda.last;
  DM_Tabelas.ZQVenda.GotoBookMark(bmLocal);
  DM_Tabelas.ZQVenda.FreeBookMark(bmLocal);
  DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVendaimovel.Value,[]);
  DM_Tabelas.ZQComprador.First;

  // grava aqui o nome do antigo comprador
  ZQcomprador_cessao.Open;


  if DM_Tabelas.CDSCompradorTemp.RecordCount=1 then
  begin
    DM_Tabelas.ZQComprador.First;
    while not DM_Tabelas.ZQComprador.Eof do
    begin
      Zpart.close;
      Zpart.SQL.clear;
      Zpart.SQL.Add('select * from participante where idpaticipante='+quotedstr(DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Text));
      Zpart.open;
      if (Zpartdoc1.Value<>DM_Tabelas.CDSCompradorTempdoc1.Value) then
      begin
        ZQcomprador_cessao.Insert;
        ZQcomprador_cessaoidvenda.Value:=DM_Tabelas.ZQVendaidvenda.Value;
        ZQcomprador_cessaodatavenda.Value:=DM_Tabelas.ZQVendadatavenda.Value;
        ZQcomprador_cessaodatacessao.Value:=strtodate(JDEntrada.DateText);
        ZQcomprador_cessaoimovel.Value:=DM_Tabelas.ZQVendaimovel.Value;
        ZQcomprador_cessaovalorvenda.Value:=DM_Tabelas.ZQVendavalorvenda.Value;
        ZQcomprador_cessaoidcomprador.Value:=DM_Tabelas.ZQCompradoridcomprador.Value;
        ZQcomprador_cessaoidparticipante.Value:=DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Value;
        ZQcomprador_cessaonome_parte.Value:=Zpartnome_parte.Value;
        ZQcomprador_cessaohistorico.Value:='Cessão de direito de compra do imóvel pelos cedente(s):'+Zpartnome_parte.Value+', '+Zpartnacionalidade.Value+', '+Zpartprofissao.Value+', '+Zpartestadocivil.Value+', portador do CPF/CNPJ nº '+
                                            Zpartdoc1.Value+', para o(s) cessionário(s): '+DM_Tabelas.CDSCompradorTempnomeparticipante.Value+', '+DM_Tabelas.CDSCompradorTempprofissao.Value+', '+DM_Tabelas.CDSCompradorTempestadocivil.Value+', portador do CPF/CNPJ nº '+DM_Tabelas.CDSCompradorTempdoc1.Value;
        ZQcomprador_cessaohora_cessao.value:=time;
        ZQcomprador_cessao.post;
      end;
      DM_Tabelas.ZQComprador.Delete;
    end;
    Zpart.close;
    ZQcomprador_cessao.Refresh;
    ZQcomprador_cessao.last;
  end

  else if DM_Tabelas.CDSCompradorTemp.RecordCount>1 then
  begin
    DM_Tabelas.CDSCompradorTemp.First;
    DM_Tabelas.ZQComprador.First;
    historico:='Cessão de direito de compra do imóvel pelos cedente(s):';
    while not DM_Tabelas.ZQComprador.Eof do
    begin
      Zpart.close;
      Zpart.SQL.clear;
      Zpart.SQL.Add('select * from participante where idpaticipante='+quotedstr(DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Text));
      Zpart.open;
      historico:=historico+Zpartnome_parte.Value+', '+Zpartnacionalidade.Value+', '+Zpartprofissao.Value+', '+Zpartestadocivil.Value+', portador do CPF/CNPJ nº '+Zpartdoc1.Value+' e ';
      DM_Tabelas.ZQComprador.Delete;
    end;
    historico2:=', para o(s) cessionário(s): ';
    while not DM_Tabelas.CDSCompradorTemp.eof do
    begin
       historico2:=historico2+' Com '+DM_Tabelas.CDSCompradorTemppercentual.Text+' %, ' +DM_Tabelas.CDSCompradorTempnomeparticipante.Value+', '+DM_Tabelas.CDSCompradorTempprofissao.Value+', '+DM_Tabelas.CDSCompradorTempestadocivil.Value+', portador do CPF/CNPJ nº '+DM_Tabelas.CDSCompradorTempdoc1.Value+' e ';
       DM_Tabelas.CDSCompradorTemp.Next;
    end;
    tm:=length(historico);
    tm:=tm-3;
    tm2:=length(historico2);
    tm2:=tm2-3;

    Zpart.First;
    DM_Tabelas.CDSCompradorTemp.First;
    if (Zpartdoc1.Value<>DM_Tabelas.CDSCompradorTempdoc1.Value) then
    begin
      ZQcomprador_cessao.Insert;
      ZQcomprador_cessaoidvenda.Value:=DM_Tabelas.ZQVendaidvenda.Value;
      ZQcomprador_cessaodatavenda.Value:=DM_Tabelas.ZQVendadatavenda.Value;
      ZQcomprador_cessaodatacessao.Value:=strtodate(JDEntrada.DateText);
      ZQcomprador_cessaoimovel.Value:=DM_Tabelas.ZQVendaimovel.Value;
      ZQcomprador_cessaovalorvenda.Value:=DM_Tabelas.ZQVendavalorvenda.Value;
      ZQcomprador_cessaoidcomprador.Value:=DM_Tabelas.ZQCompradoridcomprador.Value;
      ZQcomprador_cessaoidparticipante.Value:=DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Value;
      ZQcomprador_cessaonome_parte.Value:=Zpartnome_parte.Value;
      ZQcomprador_cessaohistorico.Value:=copy(historico,1,tm)+copy(historico2,1,tm2);
      ZQcomprador_cessaohora_cessao.value:=time;
      ZQcomprador_cessao.post;
    end;
    tm:=0;
    tm2:=0;
    historico:='';
    historico2:='';
    Zpart.close;
    ZQcomprador_cessao.Refresh;
    ZQcomprador_cessao.last;
  end;}


  // até aqui

{  DM_Tabelas.CDSCompradorTemp.First;
  while not DM_Tabelas.CDSCompradorTemp.Eof do begin
    DM_Tabelas.ZQComprador.Insert;
    DM_Tabelas.ZQCompradorpaticipante_idpaticipante.Value := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
    DM_Tabelas.ZQCompradorpercentual.Value:=DM_Tabelas.CDSCompradorTemppercentual.Value;
    if DM_Tabelas.CDSCompradorTempmarcar.Value=False then
       DM_Tabelas.ZQCompradormarcar.Value:=0
    else
       DM_Tabelas.ZQCompradormarcar.Value:=1;
    DM_Tabelas.CDSCompradorTemp.Next;
  end;
  DM_Tabelas.ZQVendedor.First;
  while not DM_Tabelas.ZQVendedor.Eof do
    DM_Tabelas.ZQVendedor.Delete;
  DM_Tabelas.CDSVendedorTemp.First;
  while not DM_Tabelas.CDSVendedorTemp.Eof do begin
    DM_Tabelas.ZQVendedor.Insert;
    DM_Tabelas.ZQVendedorcorretor_idcorretor.Value := DM_Tabelas.CDSVendedorTempcodcorretor.Value;
    DM_Tabelas.CDSVendedorTemp.Next;
  end;
  if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVendaidvenda.Text,DM_Tabelas.ZQVendaidvenda.Text+'-E-']),[]) Then
    varnumordemE := DM_Tabelas.ZQRecebimentonumordem.Value
  else Begin
    DM_tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
    varnumordemE := DM_Tabelas.ZQNumOrdemidnumordem.Value;
  end;
  if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVendaidvenda.Text,DM_Tabelas.ZQVendaidvenda.Text+'-P-']),[]) Then
    varnumordemp := DM_Tabelas.ZQRecebimentonumordem.Value
  else Begin
    DM_tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
    varnumordemp := DM_Tabelas.ZQNumOrdemidnumordem.Value;
  end;

  DM_Tabelas.CDSCompradorTemp.First;

  //copia a dados de uma tabela para outra
{  ZQRecebimento_historico_cessao.SQL.Clear;
  ZQRecebimento_historico_cessao.SQL.Add('Insert Into recebimento_historico_cessao (idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
  ZQRecebimento_historico_cessao.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
  ZQRecebimento_historico_cessao.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,nomeadversa,Reajustado,' );
  ZQRecebimento_historico_cessao.SQL.Add(' Data_reajuste,Proximo_Reajuste,Parcelas_fixas)' );

  ZQRecebimento_historico_cessao.SQL.Add(' Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
  ZQRecebimento_historico_cessao.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
  ZQRecebimento_historico_cessao.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,nomeadversa,Reajustado,' );
  ZQRecebimento_historico_cessao.SQL.Add(' Data_reajuste,Proximo_Reajuste,Parcelas_fixas from recebimento ' );
  ZQRecebimento_historico_cessao.SQL.Add(' Where  DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  ZQRecebimento_historico_cessao.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
//  ZQRecebimento_historico_cessao.SQL.Add(' Where  documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  ZQRecebimento_historico_cessao.ExecSQL;
  ZQRecebimento_historico_cessao.Close;


  recebimento_historico_excluidos.SQL.Clear;
  recebimento_historico_excluidos.SQL.Add('Insert Into recebimento_historico_excluidos (idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
  recebimento_historico_excluidos.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
  recebimento_historico_excluidos.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar)' );

  recebimento_historico_excluidos.SQL.Add(' Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
  recebimento_historico_excluidos.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
  recebimento_historico_excluidos.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar from recebimento ' );
//  recebimento_historico_excluidos.SQL.Add(' Where  DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  recebimento_historico_excluidos.SQL.Add(' Where  documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
//  recebimento_historico_excluidos.ParamByName('dt').AsDate:=JDEntrada.Date;
  recebimento_historico_excluidos.ExecSQL;
  recebimento_historico_excluidos.Close;

  // seleciona o comprador que ira as parcelas
  DM_Tabelas.CDSCompradorTemp.First;
  if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
  begin
    DM_Tabelas.CDSCompradorTemp.Filtered:=false;
    DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
    DM_Tabelas.CDSCompradorTemp.Filtered:=true;
  end;                                  }

{  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set adversa='+quotedstr(DM_Tabelas.CDSCompradorTempcodparticipante.Text));
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;

  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set nomeadversa='+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value));
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;
 }
 // Tony pediu no dia 11/04/2012 para ser como segue, por isso foi inibido como era acima


  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set idloteamento='+quotedstr(edloteamento.Text));
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(Edvenda.Text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;

  traco:='-';
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set quadralote='+quotedstr(DBEdit6.Text+traco+DBELote.text));
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(Edvenda.Text+'-%'));
 // mensagem(DM_Tabelas.ZQRecebimento.SQL.Text);
  DM_Tabelas.ZQRecebimento.ExecSQL;

  // ate aqui

  {DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('SELECT * FROM recebimento ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  DM_Tabelas.ZQRecebimento.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
  DM_Tabelas.ZQRecebimento.open;
  if DM_Tabelas.ZQRecebimento.recordcount>0 then
  begin

    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set Substituicao=''C''');
    DM_Tabelas.ZQRecebimento.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
    DM_Tabelas.ZQRecebimento.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
    DM_Tabelas.ZQRecebimento.ExecSQL;

    dts:=datetostr(date);
    dts:=copy(dts,7,7)+'-'+copy(dts,4,2)+'-'+copy(dts,1,2);
    ZQRecebimento_historico_cessao.SQL.Clear;
    ZQRecebimento_historico_cessao.SQL.Add('update recebimento_historico_cessao set Substituicao=''C'',data_cessao='+quotedstr(dts)+',hora_cessao='+quotedstr(timetostr(time)));
    ZQRecebimento_historico_cessao.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
    ZQRecebimento_historico_cessao.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
    ZQRecebimento_historico_cessao.ExecSQL;
    ZQRecebimento_historico_cessao.close;

  end
  else
  begin

    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set Substituicao=''C''');
    DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
    DM_Tabelas.ZQRecebimento.ExecSQL;

    dts:=datetostr(date);
    dts:=copy(dts,7,7)+'-'+copy(dts,4,2)+'-'+copy(dts,1,2);

    ZQRecebimento_historico_cessao.SQL.Clear;
    ZQRecebimento_historico_cessao.SQL.Add('update recebimento_historico_cessao set Substituicao=''C'',data_cessao='+quotedstr(dts)+',hora_cessao='+quotedstr(timetostr(time)));
    ZQRecebimento_historico_cessao.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
    ZQRecebimento_historico_cessao.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
    ZQRecebimento_historico_cessao.ExecSQL;
    ZQRecebimento_historico_cessao.close;

  end;

  DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_Tabelas.ZQReceb_Baixa.SQL.Add('update Receb_Baixa set Substituicao=''C''');
  DM_Tabelas.ZQReceb_Baixa.SQL.Add(' Where docum like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  DM_Tabelas.ZQReceb_Baixa.ExecSQL;


  DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_Tabelas.ZQReceb_Baixa.SQL.Add('select * from Receb_Baixa');
  DM_Tabelas.ZQReceb_Baixa.Open;

  DM_Tabelas.ZQReceb_Baixa.Refresh;
  DM_Tabelas.ZQincorp_loteame.first;}
  DesativaCampos;
//  DBGVenda.SetFocus;

  ZQcomprador_cessao.close;

{  varvenda := DM_Tabelas.ZQVendaidvenda.Value;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_Tabelas.ZQVenda.Refresh;
//  if Frm_QuadroResumo=nil then
//    Frm_QuadroResumo:=TFrm_QuadroResumo.Create(Application);
//  Frm_QuadroResumo.RLReport1.Preview;
  DM_Tabelas.ZQVenda.Filtered := False;}

 // atualizar recebimento parcelas_fixas 15/12/2011
{  ZQatualizar.close;
  ZQatualizar.SQL.Clear;
  ZQatualizar.SQL.Add('update recebimento as r join venda as v on r.venda_idvenda=v.idvenda set Parcelas_fixas=''S'' where forma_reajuste=''PARCELAS FIXAS''');
  ZQatualizar.ExecSQL;
  ZQatualizar.close;
  // ate aqui

 // zerar nosso numero das parcelas em abertos 14/02/2012
  ZQNossoNumero.close;
  ZQNossoNumero.SQL.Clear;
  ZQNossoNumero.SQL.Add('update recebimento set numboleto='''' where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%')+' and saldo>0 and recpag=''R''');
  ZQNossoNumero.ExecSQL;
  ZQNossoNumero.close;


  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('select * from recebimento where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
  DM_Tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQRecebimento.First;
  if DM_Tabelas.ZQRecebimento.RecordCount>0 then
  begin
    bar1.Max:=DM_Tabelas.ZQRecebimento.RecordCount;
    bar1.Position:=0;
    bar1.Visible:=true;
    application.ProcessMessages;
    while not DM_Tabelas.ZQRecebimento.Eof do
    begin
      bar1.Position:=DM_Tabelas.ZQRecebimento.RecNo;
      ZQRemes_Receb.close;
      ZQRemes_Receb.SQL.Clear;
      ZQRemes_Receb.SQL.Add('delete FROM remessa_receb where idrec='+quotedstr(DM_Tabelas.ZQRecebimentoidrecebimento.Text));
      ZQRemes_Receb.ExecSQL;
      DM_Tabelas.ZQRecebimento.Next;
    end;
    bar1.Position:=0;
    bar1.Visible:=false;
    application.ProcessMessages;
  end;}
  // ate aqui


  DM_Tabelas.ZQImovel.Locate('idimovel',imovelant.text,[]);
  DM_Tabelas.ZQImovel.Edit;
  DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'SIM';
  DM_Tabelas.ZQImovel.Post;
  imovelant.text:='';
  edloteamento.Text:='';
  edvenda.Text:='';
  DM_Tabelas.ZQImovel.Locate('idimovel',DBEImovel.Text,[]);
  DM_Tabelas.ZQImovel.Edit;
  DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString := 'NºO';
  DM_Tabelas.ZQImovel.Post;


  DM_Tabelas.ZQVenda.Post;
  bmLocal := DM_Tabelas.ZQVenda.GetBookMark;
//  DM_Tabelas.ZQVenda.refresh;
  DM_Tabelas.ZQVenda.close;
  DM_Tabelas.ZQVenda.Open;
 // DM_Tabelas.ZQVenda.last;
  DM_Tabelas.ZQVenda.GotoBookMark(bmLocal);
  DM_Tabelas.ZQVenda.FreeBookMark(bmLocal);
  DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
  DM_Tabelas.ZQComprador.First;


  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento where idloteamento='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').text));
  DM_Tabelas.ZQRecebimento.Open;

  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger,[]);
  DM_Tabelas.CDSCompradorTemp.Filtered:=false;
  Panel1.Visible:=false;
  application.ProcessMessages;
end;

procedure TFrm_Trocar.BtIncluirClick(Sender: TObject);
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
  DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime := strtodate(JDEntrada.DateText);
  JDEntrada.SetFocus;
end;

procedure TFrm_Trocar.BtEditarClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Edit;
  EloteTemp.Text := DM_tabelas.ZQVenda.FieldByName('imovel').Text;
  AtivaCampos;
  if Pag_Venda.PageIndex = 0 Then
    DXBAchaLote.SetFocus;

end;

procedure TFrm_Trocar.BTExcluirClick(Sender: TObject);
begin
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento where documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento where idloteamento='+quotedstr(DM_Tabelas.ZQVenda.FieldByName('codloteamento').text));
  DM_Tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQVenda.Delete;
end;

procedure TFrm_Trocar.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Trocar.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Trocar.DXBAchaLoteClick(Sender: TObject);
begin
  imovelant.text:=DM_Tabelas.ZQVenda.FieldByName('imovel').text;
  Edvenda.Text:=DM_Tabelas.ZQVenda.FieldByName('idvenda').Text;
  AbrirModal(Self, Frm_AchaLoteVenda);
  edloteamento.Text:=DM_Tabelas.ZQImovelTemp.FieldByName('loteamento_idloteamento').Text;

//  JDEntrada.Enabled:=true;
 // Perform(WM_NEXTDLGCTL, 0, 0);
//  JDEntrada.SetFocus;
  BtGravar.SetFocus;
end;

procedure TFrm_Trocar.ECompradorExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    if (DM_Tabelas.CDSCompradorTemp.RecordCount=0) or (not empty(Ecomprador.Text)) Then Begin
      if AchaCLiente(200, 10,Ecomprador.Text) Then
      Begin
        DM_Tabelas.CDSCompradorTemp.Insert;
        DM_Tabelas.CDSCompradorTempcodparticipante.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
        DM_Tabelas.CDSCompradorTempnomeparticipante.Value := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        DM_Tabelas.CDSCompradorTempordem.Value := DM_Tabelas.CDSCompradorTemp.RecordCount;
        DM_Tabelas.CDSCompradorTemppercentual.Value:=epercentual.Value;
        DM_Tabelas.CDSCompradorTemp.Post;
        Atualiza_comprador;
      end;
      Ecomprador.Text := '';
      Ecomprador.SetFocus;
    end;
    epercentual.Value:=100;
end;

procedure TFrm_Trocar.CompradorEnter(Sender: TObject);
begin
  Ecomprador.SetFocus;
end;

procedure TFrm_Trocar.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key=46) and (DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit]) Then
    DM_Tabelas.CDSCompradorTemp.Delete;
  Atualiza_comprador;
end;

procedure TFrm_Trocar.CorretorEnter(Sender: TObject);
begin
  Ecorretor.SetFocus;
end;

procedure TFrm_Trocar.EcorretorExit(Sender: TObject);
begin
  if DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit] Then
    if (DM_Tabelas.CDSVendedorTemp.RecordCount=0) or (not empty(Ecorretor.Text)) Then Begin
      if AchaCorretor(200, 10,Ecorretor.Text) Then Begin
        DM_Tabelas.CDSVendedorTemp.Insert;
        DM_Tabelas.CDSVendedorTempcodcorretor.Value := DM_TAbelas.ZQAchaCorretor.FieldByName('idcorretor').AsLargeInt;
        DM_TAbelas.CDSVendedorTempordem.Value := DM_Tabelas.CDSVendedorTemp.RecordCount;
        DM_Tabelas.CDSVendedorTemp.Post;
      end;
      Ecorretor.Text := '';
      Ecorretor.SetFocus;
    end;
end;

procedure TFrm_Trocar.CorretorExit(Sender: TObject);
begin
  BtGravar.SetFocus;
end;

procedure TFrm_Trocar.DBGVendaDrawColumnCell(Sender: TObject;
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

procedure TFrm_Trocar.DBGVendaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  RecalcularVendaAtual;
  Botoes;
end;

procedure TFrm_Trocar.DBGVendaMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  RecalcularVendaAtual;
  Botoes;
end;

procedure TFrm_Trocar.DBGVendaEnter(Sender: TObject);
begin
  RecalcularVendaAtual;
  Botoes;
end;

procedure TFrm_Trocar.DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Atualiza_comprador;
end;

procedure TFrm_Trocar.DBGrid2Enter(Sender: TObject);
begin
  Atualiza_comprador;
end;

procedure TFrm_Trocar.BTPesquisarClick(Sender: TObject);
begin
  Frm_AchaVenda.Top := Frm_Trocar.Top+100;
  Frm_AchaVenda.Left := Frm_Trocar.Left+(Frm_Trocar.Width-Frm_AchaVenda.Width)  ;
  AbrirModal(Self, Frm_AchaVenda);
  DBGVenda.SetFocus;
end;

procedure TFrm_Trocar.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQincorp_loteame.close;
  DM_Tabelas.ZQLoteamento.close;
  DM_tabelas.ZQNumOrdem.close;
  DM_Tabelas.ZQImovel.close;
  DM_Tabelas.ZQVendedor.close;
  DM_Tabelas.ZQVenda.close;
  DM_Tabelas.ZQComprador.close;
  DM_tabelas.ZQQuadras.close;
  Frm_Trocar:=nil;
  Action:=CaFree;
end;

procedure TFrm_Trocar.DBGrid2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  if ((Shift = [ssCtrl]) and (key = vk_delete)) THEN
     Abort;

end;

procedure TFrm_Trocar.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  if ((Shift = [ssCtrl]) and (key = vk_delete)) THEN
     Abort;

end;

procedure TFrm_Trocar.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  Check: Integer;
  R: TRect;
begin
  if Column.FieldName = 'Marcar' then //COLOQUE O NOME DA COLUNA, NO MEU CASO 'CONCLUIDO'
  begin
    DBGrid2.Canvas.FillRect(Rect);
    Check := 0;
    if DM_TAbelas.CDSCompradorTempmarcar.AsBoolean = true then
      Check := DFCS_CHECKED
    else
      Check := 0;
    R:=Rect;
    InflateRect(R,-2,-2); {Diminue o tamanho do CheckBox}
    DrawFrameControl(DBGrid2.Canvas.Handle,R,DFC_BUTTON, DFCS_BUTTONCHECK or Check);
  end;
end;

procedure TFrm_Trocar.DBGrid2CellClick(Column: TColumn);
begin
//  if Column.FieldName = 'Marcar' then
//  begin
    DM_TAbelas.CDSCompradorTemp.Edit;
    if DM_TAbelas.CDSCompradorTempmarcar.AsBoolean = true then
      DM_TAbelas.CDSCompradorTempmarcar.AsBoolean := false
    else
      DM_TAbelas.CDSCompradorTempmarcar.AsBoolean := true;
    DM_TAbelas.CDSCompradorTemp.Post
//  end;

end;

procedure TFrm_Trocar.BtRelatorioClick(Sender: TObject);
var
 WinWord,Docs, Doc: Variant;
begin
  WinWord := CreateOleObject('Word.Application');

  // Mostra o Word
   WinWord.Visible := False;
 // Pega uma interface para o objeto que manipula documentos
   Docs := WinWord.Documents;
   Doc := Docs.Open(ExtractFilePath(Application.ExeName)+'\Contratos\CESSAO_DIREITOS.DOC');

   //        Doc.Content.Find.Execute(FindText := '&CONTRATANTE', ReplaceWith := Manutencao_loca2.MENSALNOME.AsString);
{        Doc.Content.Find.Execute(FindText := '&CONTRATANTE', ReplaceWith := Manutencao_loca2.ClientesNOME.AsString);
        Doc.Content.Find.Execute(FindText := '&CPF', ReplaceWith := Manutencao_loca2.ClientesCPF.AsString);
        Doc.Content.Find.Execute(FindText := '&RG', ReplaceWith := Manutencao_loca2.ClientesRG.AsSTring);
        Doc.Content.Find.Execute(FindText := '&NASC', ReplaceWith := Manutencao_loca2.ClientesDATA_NASCIMENTO.AsString);
        Doc.Content.Find.Execute(FindText := '&ENDE', ReplaceWith := Manutencao_loca2.ClientesENDERECO.AsString);
        Doc.Content.Find.Execute(FindText := '&CID', ReplaceWith := Manutencao_loca2.ClientesCIDADE.AsString);
        Doc.Content.Find.Execute(FindText := '&UF', ReplaceWith := Manutencao_loca2.ClientesUF.AsString);
        //     Doc.Content.Find.Execute(FindText := '&FONE', ReplaceWith := Manutencao_loca2.ClientesFONE1.AsString);
        Doc.Content.Find.Execute(FindText := 'VALOR_IMO', ReplaceWith := Transform(Manutencao_loca2.mensalvalor.Value,'###,##0.00'));
        ValorExtenso(VExt,Manutencao_loca2.mensalvalor.Value,'','',02,60,' ',3);
        Doc.Content.Find.Execute(FindText := 'IMOREAIS', ReplaceWith := VExt[0]);
        Doc.Content.Find.Execute(FindText := '&VALIDADE', ReplaceWith := Transform(Manutencao_loca2.validade.Value,'##0'));
        Doc.Content.Find.Execute(FindText := '&VALORMENS', ReplaceWith := Transform(Manutencao_loca2.MENSALVALOR_MENS.Value,'###,##0.00'));
        dit:=DATETOSTR(Manutencao_loca2.MENSALDT_EMI.VALUE);
        dit:=copy(dit,1,2);
        Doc.Content.Find.Execute(FindText := '&DTEMI', ReplaceWith := dit);
        Doc.Content.Find.Execute(FindText := '&VCTO', ReplaceWith := DATETOSTR(Manutencao_loca2.MENSALVCTO.VALUE));
        Doc.Content.Find.Execute(FindText := 'Data_Contrato',  ReplaceWith :=  dia(DATE)+ ' de '+extmes(Date)+ ' de '+ano(date));
     // Imprime
     Doc.PrintOut(false);

     // Grava documento
     //  if Manutencao_loca2.VendasTIPO.Value='CASA' then
     Doc.SaveAs(ExtractFilePath(Application.ExeName)+'\Temp\'+DBEEmpreend.text+'_CESSAO_DIREITOS.DOC');

     // Fecha o Word
     WinWord.Quit;

     sleep(10000);

     // 2ª via

     // Cria Objeto principal de Controle
     WinWord := CreateOleObject('Word.Application');


     // Mostra o Word
     WinWord.Visible := False;

     // Pega uma interface para o objeto que manipula documentos
     Docs := WinWord.Documents;
     if Manutencao_loca2=nil then
        Manutencao_loca2:=TManutencao_loca2.Create(Application);
     // Abre um Documento
     //  if Manutencao_loca2.vendasTIPO.Value='CASA' then
         Doc := Docs.Open(ExtractFilePath(Application.ExeName)+'\Contratos\CONTRATO_SERVICOS.DOC');
     // Substitui texto via "name parameters"
     //  if Manutencao_loca2.VendasTIPO.Value='CASA' then
     begin
//        Doc.Content.Find.Execute(FindText := '&CONTRATANTE', ReplaceWith := Manutencao_loca2.MENSALNOME.AsString);
        Doc.Content.Find.Execute(FindText := '&CONTRATANTE', ReplaceWith := Manutencao_loca2.ClientesNOME.AsString);
        Doc.Content.Find.Execute(FindText := '&CPF', ReplaceWith := Manutencao_loca2.ClientesCPF.AsString);
        Doc.Content.Find.Execute(FindText := '&RG', ReplaceWith := Manutencao_loca2.ClientesRG.AsSTring);
        Doc.Content.Find.Execute(FindText := '&NASC', ReplaceWith := Manutencao_loca2.ClientesDATA_NASCIMENTO.AsString);
        Doc.Content.Find.Execute(FindText := '&ENDE', ReplaceWith := Manutencao_loca2.ClientesENDERECO.AsString);
        Doc.Content.Find.Execute(FindText := '&CID', ReplaceWith := Manutencao_loca2.ClientesCIDADE.AsString);
        Doc.Content.Find.Execute(FindText := '&UF', ReplaceWith := Manutencao_loca2.ClientesUF.AsString);
        //     Doc.Content.Find.Execute(FindText := '&FONE', ReplaceWith := Manutencao_loca2.ClientesFONE1.AsString);
        Doc.Content.Find.Execute(FindText := 'VALOR_IMO', ReplaceWith := Transform(Manutencao_loca2.mensalvalor.Value,'###,##0.00'));
        ValorExtenso(VExt,Manutencao_loca2.mensalvalor.Value,'','',02,60,' ',3);
        Doc.Content.Find.Execute(FindText := 'IMOREAIS', ReplaceWith := VExt[0]);
        Doc.Content.Find.Execute(FindText := '&VALIDADE', ReplaceWith := Transform(Manutencao_loca2.validade.Value,'##0'));
        Doc.Content.Find.Execute(FindText := '&VALORMENS', ReplaceWith := Transform(Manutencao_loca2.MENSALVALOR_MENS.Value,'###,##0.00'));
        dit:=DATETOSTR(Manutencao_loca2.MENSALDT_EMI.VALUE);
        dit:=copy(dit,1,2);
        Doc.Content.Find.Execute(FindText := '&DTEMI', ReplaceWith := dit);
        Doc.Content.Find.Execute(FindText := '&VCTO', ReplaceWith := DATETOSTR(Manutencao_loca2.MENSALVCTO.VALUE));
        Doc.Content.Find.Execute(FindText := 'Data_Contrato',  ReplaceWith :=  dia(DATE)+ ' de '+extmes(Date)+ ' de '+ano(date));
     end;
     inc(contar);
     // Imprime
     Doc.PrintOut(false);

     // Grava documento
     //  if Manutencao_loca2.VendasTIPO.Value='CASA' then
     Doc.SaveAs(ExtractFilePath(Application.ExeName)+'\Temp\Temp_CONTRATO_SERVICOS.DOC');

     // Fecha o Word
     WinWord.Quit;


 }

end;


procedure TFrm_Trocar.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Trocar, 'ZQcomprador_cessao', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQcomprador_cessao', 'ZQcomprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Trocar, 'Zpart', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'ZpartFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zparttipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartlocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'ZpartcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'ZpartcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'Zpartcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'Zpart', 'ZpartNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Trocar, 'ZQatualizar', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizaridrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizardocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizaradversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizaridloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizaridvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizardatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizartabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizardescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQatualizar', 'ZQatualizarData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Trocar, 'ZQRecebimento_historico_cessao', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaodata_cessao', 'data_cessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoSequencial', 'Sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Trocar, 'ZQNossoNumero', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeromarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeronumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeronumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeronomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumerodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQNossoNumero', 'ZQNossoNumeroData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Trocar, 'ZQRemes_Receb', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebretorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebacao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_RecebNossonumero', 'Nossonumero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebdesc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebcredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'ZQRemes_Receb', 'ZQRemes_Recebdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Trocar, 'recebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmotivo', 'motivo', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Trocar, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossequencial', 'sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
