unit Cessao;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XDBNum, XNum, StdCtrls, wwdbdatetimepicker, Mask, DBCtrls,
  Grids, DBGrids, ComCtrls, TabNotBk, ExtCtrls, XBanner,
  XDate, XDBEdit, DB, JvExControls, JvXPCore, JvXPBar, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, JvExComCtrls, JvDateTimePicker, XDBDate,
  FnpNumericEdit,Comobj;

type
  TFrm_Cessao = class(TForm)
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
  Frm_Cessao: TFrm_Cessao;
  ncompr: integer;
implementation

uses tabelas, Funcoes, AchaLoteVenda, RelVenda, QuadroResumo, AchaVenda,
  Loteamento, uRuntimeFields, CessaoRecebimentos;

type
  { GetCalcFields e o ciclo nativo que recalcula os lookups no buffer atual
    sem colocar o dataset em modo de edicao. }
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

{$R *.dfm}


procedure TFrm_Cessao.FormShow(Sender: TObject);
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
  ncompr:=0;
  
  JDEntrada.DateText:=datetostr(date);
  Pag_Venda.PageIndex := 0;
  DesativaCampos;
  RecalcularVendaAtual;
  DBGVenda.SetFocus;
end;

procedure TFrm_Cessao.botoes;
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
  RecalcularVendaAtual;
  Atualiza;
End;

procedure TFrm_Cessao.Atualiza;
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

procedure TFrm_Cessao.Atualiza_comprador;
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

procedure TFrm_Cessao.DesativaCampos;
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

procedure TFrm_Cessao.AtivaCampos;
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

procedure TFrm_Cessao.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.First;
  RecalcularVendaAtual;
  DBGVenda.SetFocus;
end;


procedure TFrm_Cessao.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Prior;
  RecalcularVendaAtual;
  DBGVenda.SetFocus;
end;

procedure TFrm_Cessao.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Next;
  RecalcularVendaAtual;
  DBGVenda.SetFocus;
end;

procedure TFrm_Cessao.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Last;
  RecalcularVendaAtual;
  DBGVenda.SetFocus;
end;

procedure TFrm_Cessao.BtCancelarClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Cancel;
  DesativaCampos;
  DBGVenda.SetFocus;
end;

procedure TFrm_Cessao.BtGravarClick(Sender: TObject);
var
  tm,tm2,varnumordemE, varnumordemP, varvezes, varvenda : integer;
  dts,historico,historico2,entrada:String;
  bmLocal : TBookmark;

begin
  if simnao('A data da Cessão está Correta ?','SIM')=false then
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

  // preparar cessao novos ids    11/09/2026
  PrepararCessaoRecebimentos(DM_Tabelas.zconeccao);

  DM_Tabelas.ZQVenda.Post;
  bmLocal := DM_Tabelas.ZQVenda.GetBookMark;
  DM_Tabelas.ZQVenda.refresh;
  DM_Tabelas.ZQVenda.close;
  DM_Tabelas.ZQVenda.Open;
  DM_Tabelas.ZQVenda.last;
  DM_Tabelas.ZQVenda.GotoBookMark(bmLocal);
  DM_Tabelas.ZQVenda.FreeBookMark(bmLocal);
  DM_Tabelas.ZQImovel.Locate('idimovel',DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt,[]);
  DM_Tabelas.ZQComprador.First;

  // grava aqui o nome do antigo comprador
  ZQcomprador_cessao.Open;

  
  if DM_Tabelas.CDSCompradorTemp.RecordCount=1 then
  begin
    DM_Tabelas.ZQComprador.First;
    DM_Tabelas.ZQComprador.DisableControls;
    while not DM_Tabelas.ZQComprador.Eof do
    begin
      Zpart.close;
      Zpart.SQL.clear;
      Zpart.SQL.Add('select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
      Zpart.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
      Zpart.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
      Zpart.SQL.Add(' from participante where idpaticipante='+quotedstr(DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').Text));
      Zpart.open;
      if (Zpart.FieldByName('doc1').AsString<>DM_Tabelas.CDSCompradorTempdoc1.Value) then
      begin
        ZQcomprador_cessao.Insert;
        ZQcomprador_cessao.FieldByName('idvenda').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
        ZQcomprador_cessao.FieldByName('datavenda').AsDateTime:=DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime;
        ZQcomprador_cessao.FieldByName('datacessao').AsDateTime:=strtodate(JDEntrada.DateText);
        ZQcomprador_cessao.FieldByName('imovel').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt;
        ZQcomprador_cessao.FieldByName('valorvenda').AsFloat:=DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat;
        ZQcomprador_cessao.FieldByName('idcomprador').AsLargeInt:=DM_Tabelas.ZQcomprador.FieldByName('idcomprador').AsLargeInt;
        ZQcomprador_cessao.FieldByName('idparticipante').AsLargeInt:=DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt;
        ZQcomprador_cessao.FieldByName('nome_parte').AsString:=Zpart.FieldByName('nome_parte').AsString;
        ZQcomprador_cessao.FieldByName('historico').AsString:='Cessão de direito de compra do imóvel pelos cedente(s):'+Zpart.FieldByName('nome_parte').AsString+', '+Zpart.FieldByName('nacionalidade').AsString+', '+Zpart.FieldByName('profissao').AsString+', '+Zpart.FieldByName('estadocivil').AsString+', portador do CPF/CNPJ nº '+
                                            Zpart.FieldByName('doc1').AsString+', para o(s) cessionário(s): '+DM_Tabelas.CDSCompradorTempnomeparticipante.Value+', '+DM_Tabelas.CDSCompradorTempprofissao.Value+', '+DM_Tabelas.CDSCompradorTempestadocivil.Value+', portador do CPF/CNPJ nº '+DM_Tabelas.CDSCompradorTempdoc1.Value;
        ZQcomprador_cessao.FieldByName('hora_cessao').AsDateTime:=time;
        ZQcomprador_cessao.post;
      end;
      DM_Tabelas.ZQComprador.Delete;
    end;
    DM_Tabelas.ZQComprador.EnableControls;
    Zpart.close;
    ZQcomprador_cessao.Refresh;
    ZQcomprador_cessao.last;
  end

  else if DM_Tabelas.CDSCompradorTemp.RecordCount>1 then
  begin
    DM_Tabelas.CDSCompradorTemp.First;
    DM_Tabelas.ZQComprador.First;
    historico:='Cessão de direito de compra do imóvel pelos cedente(s):';
    DM_Tabelas.ZQComprador.DisableControls;
    while not DM_Tabelas.ZQComprador.Eof do
    begin
      Zpart.close;
      Zpart.SQL.clear;
      Zpart.SQL.Add('select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
      Zpart.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
      Zpart.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico ');
      Zpart.SQL.Add(' from participante where idpaticipante='+quotedstr(DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').Text));
      Zpart.open;
      historico:=historico+Zpart.FieldByName('nome_parte').AsString+', '+Zpart.FieldByName('nacionalidade').AsString+', '+Zpart.FieldByName('profissao').AsString+', '+Zpart.FieldByName('estadocivil').AsString+', portador do CPF/CNPJ nº '+Zpart.FieldByName('doc1').AsString+' e ';
      DM_Tabelas.ZQComprador.Delete;
    end;
    DM_Tabelas.ZQComprador.EnableControls;
    historico2:=', para o(s) cessionário(s): ';
    DM_Tabelas.CDSCompradorTemp.DisableControls;
    while not DM_Tabelas.CDSCompradorTemp.eof do
    begin
       historico2:=historico2+' Com '+DM_Tabelas.CDSCompradorTemppercentual.Text+' %, ' +DM_Tabelas.CDSCompradorTempnomeparticipante.Value+', '+DM_Tabelas.CDSCompradorTempprofissao.Value+', '+DM_Tabelas.CDSCompradorTempestadocivil.Value+', portador do CPF/CNPJ nº '+DM_Tabelas.CDSCompradorTempdoc1.Value+' e ';
       DM_Tabelas.CDSCompradorTemp.Next;
    end;
    DM_Tabelas.CDSCompradorTemp.EnableControls;
    tm:=length(historico);
    tm:=tm-3;
    tm2:=length(historico2);
    tm2:=tm2-3;

    Zpart.First;
    DM_Tabelas.CDSCompradorTemp.First;
    if (Zpart.FieldByName('doc1').AsString<>DM_Tabelas.CDSCompradorTempdoc1.Value) then
    begin
      ZQcomprador_cessao.Insert;
      ZQcomprador_cessao.FieldByName('idvenda').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
      ZQcomprador_cessao.FieldByName('datavenda').AsDateTime:=DM_Tabelas.ZQVenda.FieldByName('datavenda').AsDateTime;
      ZQcomprador_cessao.FieldByName('datacessao').AsDateTime:=strtodate(JDEntrada.DateText);
      ZQcomprador_cessao.FieldByName('imovel').AsLargeInt:=DM_Tabelas.ZQVenda.FieldByName('imovel').AsLargeInt;
      ZQcomprador_cessao.FieldByName('valorvenda').AsFloat:=DM_Tabelas.ZQVenda.FieldByName('valorvenda').AsFloat;
      ZQcomprador_cessao.FieldByName('idcomprador').AsLargeInt:=DM_Tabelas.ZQcomprador.FieldByName('idcomprador').AsLargeInt;
      ZQcomprador_cessao.FieldByName('idparticipante').AsLargeInt:=DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt;
      ZQcomprador_cessao.FieldByName('nome_parte').AsString:=Zpart.FieldByName('nome_parte').AsString;
      ZQcomprador_cessao.FieldByName('historico').AsString:=copy(historico,1,tm)+copy(historico2,1,tm2);
      ZQcomprador_cessao.FieldByName('hora_cessao').AsDateTime:=time;
      ZQcomprador_cessao.post;
    end;
    tm:=0;
    tm2:=0;
    historico:='';
    historico2:='';
    Zpart.close;
    ZQcomprador_cessao.Refresh;
    ZQcomprador_cessao.last;
  end;


  // até aqui

  DM_Tabelas.CDSCompradorTemp.First;
  DM_Tabelas.CDSCompradorTemp.DisableControls;
  while not DM_Tabelas.CDSCompradorTemp.Eof do begin
    DM_Tabelas.ZQComprador.Insert;
    DM_Tabelas.ZQcomprador.FieldByName('paticipante_idpaticipante').AsLargeInt := DM_Tabelas.CDSCompradorTempcodparticipante.Value;
    DM_Tabelas.ZQcomprador.FieldByName('percentual').AsFloat:=DM_Tabelas.CDSCompradorTemppercentual.Value;
    if DM_Tabelas.CDSCompradorTempmarcar.Value=False then
       DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger:=0
    else
       DM_Tabelas.ZQcomprador.FieldByName('Marcar').AsInteger:=1;
    DM_Tabelas.CDSCompradorTemp.Next;
  end;
  DM_Tabelas.CDSCompradorTemp.EnableControls;
  DM_Tabelas.ZQVendedor.First;
  DM_Tabelas.ZQVendedor.DisableControls;
  while not DM_Tabelas.ZQVendedor.Eof do
    DM_Tabelas.ZQVendedor.Delete;

  DM_Tabelas.ZQVendedor.EnableControls;
  DM_Tabelas.CDSVendedorTemp.First;
  DM_Tabelas.CDSVendedorTemp.DisableControls;
  while not DM_Tabelas.CDSVendedorTemp.Eof do begin
    DM_Tabelas.ZQVendedor.Insert;
    DM_Tabelas.ZQvendedor.FieldByName('corretor_idcorretor').AsLargeInt := DM_Tabelas.CDSVendedorTempcodcorretor.Value;
    DM_Tabelas.CDSVendedorTemp.Next;
  end;
  DM_Tabelas.CDSVendedorTemp.EnableControls;
  if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text,DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-E-']),[]) Then
    varnumordemE := DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger
  else Begin
    DM_tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
    varnumordemE := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
  end;
  if DM_Tabelas.ZQRecebimento.Locate('origem,documento',VarArrayOf(['V-'+DM_Tabelas.ZQVenda.FieldByName('idvenda').Text,DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-P-']),[]) Then
    varnumordemp := DM_Tabelas.ZQRecebimento.FieldByName('numordem').AsInteger
  else Begin
    DM_tabelas.ZQNumOrdem.Insert;
    DM_Tabelas.ZQNumOrdem.Post;
    varnumordemp := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
  end;

  DM_Tabelas.CDSCompradorTemp.First;

  DM_Tabelas.ZQRecebimento.Close;
  //cessao novos ids    11/09/2026
  DM_Tabelas.zconeccao.StartTransaction;
  try
    //copia a dados de uma tabela para outra
    ZQRecebimento_historico_cessao.SQL.Clear;
    ZQRecebimento_historico_cessao.SQL.Add('Insert Into recebimento_historico_cessao (idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
    ZQRecebimento_historico_cessao.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
    ZQRecebimento_historico_cessao.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,nomeadversa,Reajustado,' );
    ZQRecebimento_historico_cessao.SQL.Add(' Data_reajuste,Proximo_Reajuste,Parcelas_fixas)' );

    ZQRecebimento_historico_cessao.SQL.Add(' Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
    ZQRecebimento_historico_cessao.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
    ZQRecebimento_historico_cessao.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,nomeadversa,Reajustado,' );
    ZQRecebimento_historico_cessao.SQL.Add(' Data_reajuste,Proximo_Reajuste,Parcelas_fixas from recebimento ' );
    ZQRecebimento_historico_cessao.SQL.Add(' Where  DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt.ToString+'-%'));
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
    recebimento_historico_excluidos.SQL.Add(' Where  documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt.ToString+'-%'));
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
    end;

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
    DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set adversa='+quotedstr(DM_Tabelas.CDSCompradorTempcodparticipante.Text));
    DM_Tabelas.ZQRecebimento.SQL.Add(' Where saldo>0 and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt.ToString+'-%'));
    DM_Tabelas.ZQRecebimento.ExecSQL;

    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set nomeadversa='+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value));
    DM_Tabelas.ZQRecebimento.SQL.Add(' Where saldo>0 and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt.ToString+'-%'));
    DM_Tabelas.ZQRecebimento.ExecSQL;

    // As parcelas copiadas ja possuem o comprador novo. cessao novos ids    11/09/2026
    RenumerarRecebimentosCessao(DM_Tabelas.zconeccao, DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt, StrToDate(JDEntrada.DateText));

    DM_Tabelas.zconeccao.Commit;
  except
    if DM_Tabelas.zconeccao.InTransaction then
      DM_Tabelas.zconeccao.Rollback;
    DM_Tabelas.CDSCompradorTemp.Filtered := False;
    Panel1.Visible := False;
    raise;
  end;


//
//  //copia a dados de uma tabela para outra
//  ZQRecebimento_historico_cessao.SQL.Clear;
//  ZQRecebimento_historico_cessao.SQL.Add('Insert Into recebimento_historico_cessao (idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
//  ZQRecebimento_historico_cessao.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
//  ZQRecebimento_historico_cessao.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,nomeadversa,Reajustado,' );
//  ZQRecebimento_historico_cessao.SQL.Add(' Data_reajuste,Proximo_Reajuste,Parcelas_fixas)' );
//
//  ZQRecebimento_historico_cessao.SQL.Add(' Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
//  ZQRecebimento_historico_cessao.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
//  ZQRecebimento_historico_cessao.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar,nomeadversa,Reajustado,' );
//  ZQRecebimento_historico_cessao.SQL.Add(' Data_reajuste,Proximo_Reajuste,Parcelas_fixas from recebimento ' );
//  ZQRecebimento_historico_cessao.SQL.Add(' Where  DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
//  ZQRecebimento_historico_cessao.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
////  ZQRecebimento_historico_cessao.SQL.Add(' Where  documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
//  ZQRecebimento_historico_cessao.ExecSQL;
//  ZQRecebimento_historico_cessao.Close;
//
//
//  recebimento_historico_excluidos.SQL.Clear;
//  recebimento_historico_excluidos.SQL.Add('Insert Into recebimento_historico_excluidos (idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
//  recebimento_historico_excluidos.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
//  recebimento_historico_excluidos.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar)' );
//
//  recebimento_historico_excluidos.SQL.Add(' Select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,' );
//  recebimento_historico_excluidos.SQL.Add(' Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,custodaparcela,origem,' );
//  recebimento_historico_excluidos.SQL.Add(' adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,somar from recebimento ' );
////  recebimento_historico_excluidos.SQL.Add(' Where  DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
//  recebimento_historico_excluidos.SQL.Add(' Where  documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
////  recebimento_historico_excluidos.ParamByName('dt').AsDate:=JDEntrada.Date;
//  recebimento_historico_excluidos.ExecSQL;
//  recebimento_historico_excluidos.Close;
//
//  // seleciona o comprador que ira as parcelas
//  DM_Tabelas.CDSCompradorTemp.First;
//  if DM_Tabelas.CDSCompradorTemp.recordcount>1 then
//  begin
//    DM_Tabelas.CDSCompradorTemp.Filtered:=false;
//    DM_Tabelas.CDSCompradorTemp.Filter:='marcar=''true''';
//    DM_Tabelas.CDSCompradorTemp.Filtered:=true;
//  end;
//
//{  DM_Tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set adversa='+quotedstr(DM_Tabelas.CDSCompradorTempcodparticipante.Text));
//  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
//  DM_Tabelas.ZQRecebimento.ExecSQL;
//
//  DM_Tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set nomeadversa='+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value));
//  DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVendaidvenda.text+'-%'));
//  DM_Tabelas.ZQRecebimento.ExecSQL;
// }
// // Tony pediu no dia 11/04/2012 para ser como segue, por isso foi inibido como era acima
//
//
//  DM_Tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set adversa='+quotedstr(DM_Tabelas.CDSCompradorTempcodparticipante.Text));
//  DM_Tabelas.ZQRecebimento.SQL.Add(' Where saldo>0 and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
//  DM_Tabelas.ZQRecebimento.ExecSQL;
//
//  DM_Tabelas.ZQRecebimento.SQL.Clear;
//  DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set nomeadversa='+quotedstr(DM_Tabelas.CDSCompradorTempnomeparticipante.Value));
//  DM_Tabelas.ZQRecebimento.SQL.Add(' Where saldo>0 and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
//  DM_Tabelas.ZQRecebimento.ExecSQL;

  // ate aqui

  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' FROM recebimento ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
  DM_Tabelas.ZQRecebimento.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
  DM_Tabelas.ZQRecebimento.open;
  if DM_Tabelas.ZQRecebimento.recordcount>0 then
  begin

    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set Substituicao=''C''');
    DM_Tabelas.ZQRecebimento.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
    DM_Tabelas.ZQRecebimento.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
    DM_Tabelas.ZQRecebimento.ExecSQL;

    dts:=datetostr(date);
    dts:=copy(dts,7,7)+'-'+copy(dts,4,2)+'-'+copy(dts,1,2);
    ZQRecebimento_historico_cessao.SQL.Clear;
    ZQRecebimento_historico_cessao.SQL.Add('update recebimento_historico_cessao set Substituicao=''C'',data_cessao='+quotedstr(dts)+',hora_cessao='+quotedstr(timetostr(time)));
    ZQRecebimento_historico_cessao.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
    ZQRecebimento_historico_cessao.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
    ZQRecebimento_historico_cessao.ExecSQL;
    ZQRecebimento_historico_cessao.close;

  end
  else
  begin

    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('update recebimento set Substituicao=''C''');
    DM_Tabelas.ZQRecebimento.SQL.Add(' Where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
    DM_Tabelas.ZQRecebimento.ExecSQL;

    dts:=datetostr(date);
    dts:=copy(dts,7,7)+'-'+copy(dts,4,2)+'-'+copy(dts,1,2);

    ZQRecebimento_historico_cessao.SQL.Clear;
    ZQRecebimento_historico_cessao.SQL.Add('update recebimento_historico_cessao set Substituicao=''C'',data_cessao='+quotedstr(dts)+',hora_cessao='+quotedstr(timetostr(time)));
    ZQRecebimento_historico_cessao.SQL.Add(' Where DT_Vencimento>= :dt and documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
    ZQRecebimento_historico_cessao.ParamByName('dt').AsDate:=strtodate(JDEntrada.DateText);
    ZQRecebimento_historico_cessao.ExecSQL;
    ZQRecebimento_historico_cessao.close;

  end;

  DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_Tabelas.ZQReceb_Baixa.SQL.Add('update Receb_Baixa set Substituicao=''C''');
  DM_Tabelas.ZQReceb_Baixa.SQL.Add(' Where docum like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
  DM_Tabelas.ZQReceb_Baixa.ExecSQL;


  DM_Tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_Tabelas.ZQReceb_Baixa.SQL.Add('select * from Receb_Baixa');
  DM_Tabelas.ZQReceb_Baixa.Open;

  DM_Tabelas.ZQReceb_Baixa.Refresh;
  DM_Tabelas.ZQincorp_loteame.first;
  DesativaCampos;
  DBGVenda.SetFocus;

  ZQcomprador_cessao.close;

  varvenda := DM_Tabelas.ZQVenda.FieldByName('idvenda').AsLargeInt;
  DM_TAbelas.ZQVenda.Filter := 'idvenda='+inttostr(varvenda);
  DM_Tabelas.ZQVenda.Filtered := True;
  DM_TAbelas.ZQCompr_Dados.Open;
  DM_TAbelas.ZQCompr_Dados.Refresh;
  DM_Tabelas.ZQVenda.Refresh;
//  if Frm_QuadroResumo=nil then
//    Frm_QuadroResumo:=TFrm_QuadroResumo.Create(Application);
//  Frm_QuadroResumo.RLReport1.Preview;
  DM_Tabelas.ZQVenda.Filtered := False;

 // atualizar recebimento parcelas_fixas 15/12/2011
  ZQatualizar.close;
  ZQatualizar.SQL.Clear;
  ZQatualizar.SQL.Add('update recebimento as r join venda as v on r.venda_idvenda=v.idvenda set Parcelas_fixas=''S'' where forma_reajuste=''PARCELAS FIXAS''');
  ZQatualizar.ExecSQL;
  ZQatualizar.close;
  // ate aqui

 // zerar nosso numero das parcelas em abertos 14/02/2012
  if simnao('Zerar Nosso Nº e Remessas Geradas até Agora ?','SIM') then
  begin
    ZQNossoNumero.close;
    ZQNossoNumero.SQL.Clear;
    ZQNossoNumero.SQL.Add('update recebimento set numboleto='''' where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%')+' and saldo>0 and recpag=''R''');
    ZQNossoNumero.ExecSQL;
    ZQNossoNumero.close;


    DM_Tabelas.ZQRecebimento.close;
    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%')+' and saldo>0 and recpag=''R''');
    DM_Tabelas.ZQRecebimento.Open;
    DM_Tabelas.ZQRecebimento.First;
    if DM_Tabelas.ZQRecebimento.RecordCount>0 then
    begin
      bar1.Max:=DM_Tabelas.ZQRecebimento.RecordCount;
      bar1.Position:=0;
      bar1.Visible:=true;
      application.ProcessMessages;
      DM_Tabelas.ZQRecebimento.DisableControls;
      while not DM_Tabelas.ZQRecebimento.Eof do
      begin
        bar1.Position:=DM_Tabelas.ZQRecebimento.RecNo;
//        ZQRemes_Receb.close;
//        ZQRemes_Receb.SQL.Clear;
//        ZQRemes_Receb.SQL.Add('delete FROM remessa_receb where idrec='+quotedstr(DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Text));
//        ZQRemes_Receb.ExecSQL;

        // A resposta SIM remove remessas atuais e anteriores, mas preserva o mapa.  cessao novos ids    11/09/2026
        ExcluirRemessasDaParcela(DM_Tabelas.zconeccao, DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt);

        DM_Tabelas.ZQRecebimento.Next;
      end;
      DM_Tabelas.ZQRecebimento.EnableControls;
      bar1.Position:=0;
      bar1.Visible:=false;
      application.ProcessMessages;
    end;
  end;
  // ate aqui

  DM_Tabelas.ZQRecebimento.close;
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
  DM_Tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQRecebimento.Refresh;
  DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQVenda.FieldByName('codloteamento').AsInteger,[]);
  DM_Tabelas.CDSCompradorTemp.Filtered:=false;
  Panel1.Visible:=false;
  application.ProcessMessages;

end;

procedure TFrm_Cessao.BtIncluirClick(Sender: TObject);
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

procedure TFrm_Cessao.BtEditarClick(Sender: TObject);
begin
  DM_Tabelas.ZQVenda.Edit;
  EloteTemp.Text := DM_tabelas.ZQVenda.FieldByName('imovel').Text;
  AtivaCampos;
  if Pag_Venda.PageIndex = 0 Then
    JDEntrada.SetFocus;

end;

procedure TFrm_Cessao.BTExcluirClick(Sender: TObject);
begin
  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('delete from recebimento where documento like '+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').Text+'-%'));
  DM_Tabelas.ZQRecebimento.ExecSQL;

  DM_Tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  DM_Tabelas.ZQRecebimento.SQL.Add(' from recebimento where documento like'+quotedstr(DM_Tabelas.ZQVenda.FieldByName('idvenda').text+'-%'));
  DM_Tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQVenda.Delete;
end;

procedure TFrm_Cessao.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Cessao.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Cessao.DXBAchaLoteClick(Sender: TObject);
begin
  AbrirModal(Self, Frm_AchaLoteVenda);
  Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Cessao.ECompradorExit(Sender: TObject);
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

procedure TFrm_Cessao.CompradorEnter(Sender: TObject);
begin
  Ecomprador.SetFocus;
end;

procedure TFrm_Cessao.DBGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key=46) and (DM_Tabelas.ZQVenda.State in [DsInsert, DsEdit]) Then
    DM_Tabelas.CDSCompradorTemp.Delete;
  Atualiza_comprador;
end;

procedure TFrm_Cessao.CorretorEnter(Sender: TObject);
begin
  Ecorretor.SetFocus;
end;

procedure TFrm_Cessao.EcorretorExit(Sender: TObject);
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

procedure TFrm_Cessao.CorretorExit(Sender: TObject);
begin
  BtGravar.SetFocus;
end;

procedure TFrm_Cessao.DBGVendaDrawColumnCell(Sender: TObject;
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

procedure TFrm_Cessao.DBGVendaKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  Botoes;
end;

procedure TFrm_Cessao.DBGVendaMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Botoes;
end;

procedure TFrm_Cessao.DBGVendaEnter(Sender: TObject);
begin
  Botoes;
end;

procedure TFrm_Cessao.DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Atualiza_comprador;
end;

procedure TFrm_Cessao.DBGrid2Enter(Sender: TObject);
begin
  Atualiza_comprador;
end;

procedure TFrm_Cessao.BTPesquisarClick(Sender: TObject);
begin
  Frm_AchaVenda.Top := Frm_Cessao.Top+100;
  Frm_AchaVenda.Left := Frm_Cessao.Left+(Frm_Cessao.Width-Frm_AchaVenda.Width)  ;
  AbrirModal(Self, Frm_AchaVenda);
  DBGVenda.SetFocus;
end;

procedure TFrm_Cessao.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.ZQincorp_loteame.close;
  DM_Tabelas.ZQLoteamento.close;
  DM_tabelas.ZQNumOrdem.close;
  DM_Tabelas.ZQImovel.close;
  DM_Tabelas.ZQVendedor.close;
  DM_Tabelas.ZQVenda.close;
  DM_Tabelas.ZQComprador.close;
  DM_tabelas.ZQQuadras.close;
  Frm_Cessao:=nil;
  Action:=CaFree;
end;

procedure TFrm_Cessao.DBGrid2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  if ((Shift = [ssCtrl]) and (key = vk_delete)) THEN
     Abort;

end;

procedure TFrm_Cessao.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //bloqueia a tecla ctrl+del
  if ((Shift = [ssCtrl]) and (key = vk_delete)) THEN
     Abort;

end;

procedure TFrm_Cessao.DBGrid2DrawColumnCell(Sender: TObject;
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

procedure TFrm_Cessao.DBGrid2CellClick(Column: TColumn);
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

procedure TFrm_Cessao.BtRelatorioClick(Sender: TObject);
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


procedure TFrm_Cessao.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Cessao, 'ZQcomprador_cessao', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaodatacessao', 'datacessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaoimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaovalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaoidparticipante', 'idparticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaonome_parte', 'nome_parte', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaohistorico', 'historico', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQcomprador_cessao', 'ZQcomprador_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Cessao, 'Zpart', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'ZpartFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zparttipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartlocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'ZpartcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'ZpartcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'Zpartcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'Zpart', 'ZpartNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Cessao, 'ZQatualizar', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizaridrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizardocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizaradversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizaridloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizaridvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizardatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizartabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizardescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQatualizar', 'ZQatualizarData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Cessao, 'ZQRecebimento_historico_cessao', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaodata_cessao', 'data_cessao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaohora_cessao', 'hora_cessao', TTimeField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRecebimento_historico_cessao', 'ZQRecebimento_historico_cessaoSequencial', 'Sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Cessao, 'ZQNossoNumero', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeromarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeronumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeronumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeronomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumerodescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQNossoNumero', 'ZQNossoNumeroData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Cessao, 'ZQRemes_Receb', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebretorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebacao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_RecebNossonumero', 'Nossonumero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebdesc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebcredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'ZQRemes_Receb', 'ZQRemes_Recebdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Cessao, 'recebimento_historico_excluidos', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidoscustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosmotivo', 'motivo', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidosProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Cessao, 'recebimento_historico_excluidos', 'recebimento_historico_excluidossequencial', 'sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
