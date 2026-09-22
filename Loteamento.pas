unit Loteamento;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, ExtCtrls, Grids, DBGrids,
  DBCtrls, Mask, wwdbdatetimepicker, ComCtrls, TabNotBk, DB, ExtDlgs, XNum,
  XDBNum;

type
  TFrm_Loteamento = class(TForm)
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTExcluir: TdxButtonArround;
    BtEditar: TdxButtonArround;
    BtIncluir: TdxButtonArround;
    BtCancelar: TdxButtonArround;
    BtGravar: TdxButtonArround;
    BTPesquisar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    BtUltimo: TdxButtonArround;
    BtProximo: TdxButtonArround;
    BTAnterior: TdxButtonArround;
    BtPrimeiro: TdxButtonArround;
    BTFechar: TdxButtonArround;
    BtProcessar: TdxButtonArround;
    XBanner4: TXBanner;
    Label2: TLabel;
    DBGLotemamento: TDBGrid;
    Pag_Loteamento: TTabbedNotebook;
    XBanner10: TXBanner;
    XBanner11: TXBanner;
    DBEMatricula: TDBEdit;
    GBTestemunha: TGroupBox;
    DBETestemunha4: TDBEdit;
    Label13: TLabel;
    DBETestemunha3: TDBEdit;
    Label12: TLabel;
    DBETestemunha2: TDBEdit;
    Label11: TLabel;
    DBETestemunha1: TDBEdit;
    Label10: TLabel;
    XBanner2: TXBanner;
    GBIncorp: TGroupBox;
    XBanner3: TXBanner;
    Label8: TLabel;
    DBECod: TDBEdit;
    Label9: TLabel;
    Label14: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    DBEBAirro: TDBEdit;
    DBECadastrado: TwwDBDateTimePicker;
    DBEInalgurado: TwwDBDateTimePicker;
    DBEConcluido: TwwDBDateTimePicker;
    EUsuCidade: TEdit;
    Label24: TLabel;
    EEstado: TEdit;
    Label15: TLabel;
    DBELogo: TDBImage;
    Label16: TLabel;
    EscolhePict: TOpenPictureDialog;
    Image1: TImage;
    DBEdit1: TDBEdit;
    Label1: TLabel;
    DBEApelido: TDBEdit;
    DBGrid1: TDBGrid;
    Label3: TLabel;
    Eparticipante: TEdit;
    Label4: TLabel;
    DBEPercent: TXDBNumEdit;
    DBCBTipo: TDBComboBox;
    Label5: TLabel;
    GroupBox1: TGroupBox;
    XBanner1: TXBanner;
    DBGQruadra: TDBGrid;
    EContaBancaria: TEdit;
    Label6: TLabel;
    DBENomeloteamento: TDBEdit;
    Label7: TLabel;
    XBanner5: TXBanner;
    GroupBox2: TGroupBox;
    XBanner7: TXBanner;
    DBGrid3: TDBGrid;
    edprocuradores: TEdit;
    Label22: TLabel;
    CBAssina: TComboBox;
    Label21: TLabel;
    Label23: TLabel;
    DBEdit2: TDBEdit;
    procedure EUsuCidadeExit(Sender: TObject);
    procedure botoes;
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure BTPesquisarClick(Sender: TObject);
    procedure BtRelatorioClick(Sender: TObject);
    procedure BtProcessarClick(Sender: TObject);
    procedure ativar_campos;
    procedure desativar_campos;
    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBEMapaClick(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure DBEMapaExit(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure DBECodChange(Sender: TObject);
    procedure DBELogoClick(Sender: TObject);
    procedure EparticipanteExit(Sender: TObject);
    Procedure atualiza_lotes_incorp;
    procedure atualiza_lotes;
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit1Exit(Sender: TObject);
    procedure DBGLotemamentoEnter(Sender: TObject);
    procedure DBGLotemamentoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGLotemamentoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGLotemamentoDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBGrid1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GroupBox1Exit(Sender: TObject);
    procedure EContaBancariaExit(Sender: TObject);
    procedure DBGQruadraKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid3KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CBAssinaExit(Sender: TObject);
    procedure DBGrid3DblClick(Sender: TObject);
    procedure DBGQruadraKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
  end;

var
  Frm_Loteamento: TFrm_Loteamento;

implementation

uses tabelas, Funcoes, Mapa_loteamento, RelLOTEAMENTO, uRuntimeFields;

{$R *.dfm}

constructor TFrm_Loteamento.Create(AOwner: TComponent);
begin
  { O DBGrid resolve os nomes das colunas durante o carregamento do DFM.
    Garanta o lookup antes que o inherited Create carregue os componentes. }
  EnsureLoteamentoNomeCidadeField(DM_Tabelas);
  inherited Create(AOwner);
end;

procedure TFrm_Loteamento.EUsuCidadeExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQLoteamento.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(EUsuCidade.text) Then
      EUsuCidade.SetFocus
    else Begin
      EUsuCidade.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EEstado.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQLoteamento.FieldByName('cidade_idcidade').AsLargeInt := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;

end;

procedure TFrm_Loteamento.botoes;
Begin
  if DM_Tabelas = nil then
  begin
    BtPrimeiro.Enabled := False;
    BtAnterior.Enabled := False;
    BtProximo.Enabled := False;
    BtUltimo.Enabled := False;
    Exit;
  end;
  if not DM_Tabelas.ZQLoteamento.Active then
  begin
    BtPrimeiro.Enabled := False;
    BtAnterior.Enabled := False;
    BtProximo.Enabled := False;
    BtUltimo.Enabled := False;
    atualiza_lotes;
    Exit;
  end;
  if DM_Tabelas.ZQLoteamento.IsEmpty then
  begin
    BtPrimeiro.Enabled := False;
    BtAnterior.Enabled := False;
    BtProximo.Enabled := False;
    BtUltimo.Enabled := False;
    atualiza_lotes;
    Exit;
  end;

  BtPrimeiro.Enabled := True;
  BtAnterior.Enabled := True;
  BtProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQLoteamento.Recno <=1 Then Begin
    BtPrimeiro.Enabled := False;
    BtAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQLoteamento.RecNo=DM_Tabelas.ZQLoteamento.RecordCount Then Begin
    BtProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  atualiza_lotes;
End;

procedure TFrm_Loteamento.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQLoteamento.First;
  Botoes;
end;

procedure TFrm_Loteamento.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQLoteamento.Prior;
  Botoes;
end;

procedure TFrm_Loteamento.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQLoteamento.Next;
  Botoes;
end;

procedure TFrm_Loteamento.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQLoteamento.Last;
  Botoes;
end;

procedure TFrm_Loteamento.BtGravarClick(Sender: TObject);
Var
  varx : integer;
  vary : Double;
begin
//  if not Verif_senha('Loteamento','Gravar inclusão ou edição','Loteamento: '+DM_Tabelas.ZQLoteamentoidloteamento.Text+' - '+DM_Tabelas.ZQLoteamentomatriculaloteamento.Value) Then Exit;
//  if empty(DM_Tabelas.ZQLoteamentomatriculaloteamento.Value) Then Begin
//    Showmessage('A matricula do loteamento não pode ficar em branco!!!...');
//    Pag_Loteamento.PageIndex := 0;
//    DBEMatricula.SetFocus;
//    exit;
//  end;
  if DM_Tabelas.ZQLoteamento.FieldByName('cidade_idcidade').AsLargeInt=0 Then Begin
    Showmessage('A cidade do loteamento não pode ficar em branco!!!...');
    Pag_Loteamento.PageIndex := 0;
    EUsuCidade.SetFocus;
    exit;
  end;
  vary := 0;
  DM_Tabelas.CDSIncorp.First;
  while not DM_Tabelas.CDSIncorp.Eof do Begin
    vary := vary+ DM_Tabelas.CDSIncorppercent.Value;
    DM_Tabelas.CDSIncorp.Next;
  end;
  if vary<100 Then Begin
    Showmessage('A soma do % de participação deve ser igual a 100!!!...');
    Pag_Loteamento.PageIndex := 1;
    DBEPercent.SetFocus;
    exit;
  end;
  DM_Tabelas.ZQLoteamento.Post;
  DM_tabelas.ZQQuadras.First;
  While not DM_TAbelas.ZQQuadras.Eof do
    DM_Tabelas.ZQQuadras.Delete;
  DM_tabelas.CDSQuadrasTemp.First;
  while not DM_Tabelas.CDSQuadrasTemp.Eof do begin
    DM_tabelas.ZQQuadras.Insert;
    DM_tabelas.ZQquadras.FieldByName('descricao').AsString := DM_tabelas.CDSQuadrasTempquadra.Value;
    DM_tabelas.ZQquadras.FieldByName('apedlido').AsString := DM_Tabelas.CDSQuadrasTempapelido.Value;
    DM_Tabelas.CDSQuadrasTemp.Next;
    DM_Tabelas.ZQQuadras.Post;
  end;
  DM_Tabelas.ZQincorp_loteame.First;
  While not DM_Tabelas.ZQincorp_loteame.Eof do
    DM_Tabelas.ZQincorp_loteame.Delete;
  DM_Tabelas.CDSIncorp.First;
  While not DM_Tabelas.CDSIncorp.Eof do begin
    DM_tabelas.ZQincorp_loteame.Insert;
    DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt := DM_Tabelas.CDSIncorpcodparticipante.Value;
    DM_Tabelas.ZQincorp_loteame.FieldByName('loteamento_idloteamento').AsLargeInt := DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt;
    DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').AsFloat := DM_Tabelas.CDSIncorppercent.Value;
    DM_Tabelas.ZQincorp_loteame.FieldByName('codcontabancaria').AsInteger := DM_Tabelas.CDSIncorpcodcontabancaria.Value;
    DM_Tabelas.ZQincorp_loteame.FieldByName('digito_dif').AsString := DM_Tabelas.CDSIncorpdigito_dif.Value;
    DM_Tabelas.ZQincorp_loteame.Post;
    DM_Tabelas.CDSIncorp.Next;
  end;
  desativar_campos;
end;

procedure TFrm_Loteamento.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Cancelar inclusão ou edição','Loteamento: '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text+' - '+DM_Tabelas.ZQLoteamento.FieldByName('matriculaloteamento').AsString) Then Exit;
  DM_Tabelas.ZQLoteamento.Cancel;
  desativar_campos;
end;

procedure TFrm_Loteamento.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Inclusão','') Then Exit;
  DM_Tabelas.CDSQuadrasTemp.Close;
  DM_Tabelas.CDSQuadrasTemp.CreateDataSet;
  DM_Tabelas.CDSIncorp.Close;
  DM_Tabelas.CDSIncorp.CreateDataSet;
  Eparticipante.Text := '';
  Pag_Loteamento.PageIndex := 0;
  DBENomeloteamento.SetFocus;
  ativar_campos;
  DM_Tabelas.ZQLoteamento.Insert;
end;

procedure TFrm_Loteamento.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Edição','Loteamento: '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text+' - '+DM_Tabelas.ZQLoteamento.FieldByName('matriculaloteamento').AsString) Then Exit;
  ativar_campos;
  if Pag_Loteamento.PageIndex = 0 then
    DBENomeloteamento.SetFocus
  else
    DBETestemunha1.SetFocus;
  DM_Tabelas.ZQLoteamento.Edit;
end;

procedure TFrm_Loteamento.BTExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Exclusão','Loteamento: '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text+' - '+DM_Tabelas.ZQLoteamento.FieldByName('matriculaloteamento').AsString) Then Exit;
  if simnao('Confirma a exclusão do empreendimento ?','SIM') then
  begin
    DM_tabelas.ZQQuadras.First;
    While not DM_TAbelas.ZQQuadras.Eof do
      DM_Tabelas.ZQQuadras.Delete;
    DM_Tabelas.ZQLoteamento.Delete;
    atualiza_lotes;
  end;
end;

procedure TFrm_Loteamento.BTPesquisarClick(Sender: TObject);
begin
// pesquisa
end;

procedure TFrm_Loteamento.BtRelatorioClick(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Relatório','') Then Exit;
  Frm_Loteamento.FormStyle:=fsNormal;
  if Frm_RelLoteamento=nil then
     Frm_RelLoteamento   := TFrm_RelLoteamento.Create(Application);
  Frm_RelLoteamento.Top  := Frm_RelLoteamento.Top+105;
  Frm_RelLoteamento.Left := Frm_RelLoteamento.Left-5+(Frm_RelLoteamento.Width-Frm_RelLoteamento.Width);
  AbrirModal(Self, Frm_RelLoteamento);
  freeandnil(Frm_RelLoteamento);
  Frm_Loteamento.FormStyle:=fsStayOnTop;
end;

procedure TFrm_Loteamento.BtProcessarClick(Sender: TObject);
begin
//  if not Verif_senha('Loteamento','Processar','') Then Exit;
end;

procedure TFrm_Loteamento.ativar_campos;
Begin
  DBCBTipo.ReadOnly          := False;
  DBEMatricula.ReadOnly      := False;
  DBECadastrado.ReadOnly     := False;
  DBEInalgurado.ReadOnly     := False;
  DBEConcluido.ReadOnly      := False;
  DBENomeloteamento.ReadOnly := False;
  DBEApelido.ReadOnly        := False;
  EUSUCIDADE.ReadOnly        := False;
  DBEBAirro.ReadOnly         := False;
  DBELogo.ReadOnly           := False;
  DBGQruadra.ReadOnly        := False;
  GBTestemunha.Enabled       := True;
  DBEPercent.ReadOnly        := False;
  Eparticipante.ReadOnly     := False;
  BTPrimeiro.Enabled         := False;
  BTAnterior.Enabled         := False;
  BTProximo.Enabled          := False;
  BtUltimo.Enabled           := False;
  BtGravar.Enabled           := True;
  BtCancelar.Enabled         := True;
  BtProcessar.Enabled        := False;
  BtFechar.Enabled           := False;
  BtIncluir.Enabled          := False;
  BtEditar.Enabled           := False;
  BtExcluir.Enabled          := False;
  BtPesquisar.Enabled        := False;
  BtRelatorio.Enabled        := False;
  DBGLotemamento.Enabled     := False;
end;

procedure TFrm_Loteamento.desativar_campos;
Begin
  DBCBTipo.ReadOnly          := True;
  DBEMatricula.ReadOnly      := True;
  DBENomeloteamento.ReadOnly := True;
  DBEApelido.ReadOnly        := True;
  DBECadastrado.ReadOnly     := True;
  DBEInalgurado.ReadOnly     := True;
  DBEConcluido.ReadOnly      := True;
  EUSUCIDADE.ReadOnly        := True;
  DBEBAirro.ReadOnly         := True;
  DBELogo.ReadOnly           := True;
  GBTestemunha.Enabled       := False;
  DBGQruadra.ReadOnly        := True;
  DBEPercent.ReadOnly        := True;
  Eparticipante.ReadOnly     := True;
  BTPrimeiro.Enabled         := True;
  BTAnterior.Enabled         := True;
  BTProximo.Enabled          := True;
  BtUltimo.Enabled           := True;
  BtGravar.Enabled           := False;
  BtCancelar.Enabled         := False;
  BtProcessar.Enabled        := False;
  BtFechar.Enabled           := True;
  BtIncluir.Enabled          := True;
  BtEditar.Enabled           := True;
  BtExcluir.Enabled          := True;
  BtPesquisar.Enabled        := True;
  BtRelatorio.Enabled        := True;
  DBGLotemamento.Enabled     := True;
end;

procedure TFrm_Loteamento.FormShow(Sender: TObject);
begin
  EnsureLoteamentoNomeCidadeField(DM_Tabelas);
  Pag_Loteamento.PageIndex := 0;
  if not DM_Tabelas.ZQLoteamento.Active then
    DM_Tabelas.ZQLoteamento.Open;
  if not DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Open;
  DM_Tabelas.ZQincorp_loteame.open;
  DM_Tabelas.ZQAchaContaBanc.Open;
  DM_tabelas.ZQQuadras.Open;
  DM_tabelas.ZQLoteamento.Open;
  DM_TAbelas.ZQContaBancaria.open;
  DM_Tabelas.zqprocuradores.open;
  desativar_campos;
  { A abertura nao passa pelos botoes de navegacao. Atualize o primeiro
    registro somente depois que todos os datasets dependentes estiverem ativos. }
  botoes;
  DBGLotemamento.SetFocus;
end;

procedure TFrm_Loteamento.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Loteamento.DBEMapaClick(Sender: TObject);
begin
  if DM_Tabelas.ZQLoteamento.State in [dsEdit, dsInsert] then Begin
//    EscolhePict.InitialDir := 'D:\sistemas\projeto Tabelas';
    if EscolhePict.Execute Then
      TBlobField(DM_Tabelas.ZQLoteamento.FieldByName('mapa')).LoadFromFile(EscolhePict.FileName);
  end;

end;

procedure TFrm_Loteamento.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Loteamento.DBEMapaExit(Sender: TObject);
begin
  if DM_Tabelas.ZQLoteamento.State in [DsInsert] Then Begin
    Pag_Loteamento.PageIndex := 1;
    DBETestemunha1.SetFocus;
  end;
end;

procedure TFrm_Loteamento.Image1Click(Sender: TObject);
begin
  if DM_Tabelas.ZQLoteamento.State in [dsEdit, dsInsert] then Begin
    if EscolhePict.Execute Then Begin
      DM_Tabelas.ZQLoteamento.FieldByName('pasta_mapa').AsString := EscolhePict.FileName;
      Image1.Picture.LoadFromFile(DM_Tabelas.ZQLoteamento.FieldByName('pasta_mapa').AsString);
    end;
  end
  else Begin
    if Frm_Mapa_loteamento=nil then
      Frm_Mapa_loteamento:=tFrm_Mapa_loteamento.Create(Application);
    Frm_Mapa_loteamento.RLImage1.Picture.LoadFromFile(DM_Tabelas.ZQLoteamento.FieldByName('pasta_mapa').AsString);
    Frm_Mapa_loteamento.RLReport1.Previewmodal;
    Frm_Mapa_loteamento:=nil;
  end;

end;

procedure TFrm_Loteamento.DBECodChange(Sender: TObject);
begin
  If fileexists(DM_Tabelas.ZQLoteamento.FieldByName('pasta_mapa').AsString) then Begin
    Image1.Visible := True;
    Image1.Picture.LoadFromFile(DM_Tabelas.ZQLoteamento.FieldByName('pasta_mapa').AsString)
  end
  else
    Image1.Visible := False;
end;

procedure TFrm_Loteamento.DBELogoClick(Sender: TObject);
begin
  if DM_Tabelas.ZQLoteamento.State in [dsEdit, dsInsert] then Begin
    if EscolhePict.Execute Then
      TBlobField(DM_Tabelas.ZQLoteamento.FieldByName('logotipo')).LoadFromFile(EscolhePict.FileName);
//      TBlobField(DM_Tabelas.ZQLoteamento.FieldByName('logotipo')).SaveToFile(DM_Tabelas.ZQLoteamentopasta_mapa.Value);
  end;


end;

procedure TFrm_Loteamento.EparticipanteExit(Sender: TObject);
begin
  if DM_Tabelas.ZQLoteamento.State in [DsInsert, DsEdit] Then Begin
    if not AchaCLiente(200, 10,Eparticipante.Text) Then Begin
      Eparticipante.SetFocus;
      exit;
    end
    else Begin
      DM_Tabelas.CDSIncorp.Insert;
      DM_Tabelas.CDSIncorpcodparticipante.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
      DM_Tabelas.CDSIncorp.Post;
    end;
  end;
end;

Procedure TFrm_Loteamento.atualiza_lotes_incorp;
Begin
  if DM_Tabelas = nil then
  begin
    Eparticipante.Clear;
    EContaBancaria.Clear;
    Exit;
  end;
  if not DM_Tabelas.CDSIncorp.Active then
  begin
    Eparticipante.Clear;
    EContaBancaria.Clear;
    Exit;
  end;
  if DM_Tabelas.CDSIncorp.IsEmpty then
  begin
    Eparticipante.Clear;
    EContaBancaria.Clear;
    Exit;
  end;

  Eparticipante.Text := DM_Tabelas.CDSIncorpnomeparti.Value;
  EContaBancaria.Text := DM_Tabelas.CDSIncorpcontabancaria.Value+'|'+DM_Tabelas.CDSIncorpdigito_dif.Value;
//  if DM_TAbelas.ZQContaBancaria.Locate('idconta_bancaria',DM_TAbelas.CDSIncorpcodcontabancaria.Value,[]) Then
//    EContaBancaria.Text := DM_Tabelas.ZQAchaContaBancn_conta.Value+'-'+DM_Tabelas.ZQAchaContaBancn_conta_v.Value+' - '+DM_Tabelas.ZQAchaContaBancn_no_banco.Value+' - '+DM_Tabelas.ZQAchaContaBancnometitular.Value;

End;

Procedure TFrm_Loteamento.atualiza_lotes;
var
  FCidade: TField;
  VCidade: Variant;
  VEstado: Variant;
Begin
  EUsuCidade.Clear;
  EEstado.Clear;
  if DM_Tabelas = nil then
  begin
    GroupBox1.Caption := 'Quadras  (0)';
    Eparticipante.Clear;
    EContaBancaria.Clear;
    Exit;
  end;
  if not DM_Tabelas.ZQLoteamento.Active then
  begin
    GroupBox1.Caption := 'Quadras  (0)';
    DM_Tabelas.CDSQuadrasTemp.Close;
    DM_Tabelas.CDSIncorp.Close;
    atualiza_lotes_incorp;
    Exit;
  end;
  if DM_Tabelas.ZQLoteamento.IsEmpty then
  begin
    GroupBox1.Caption := 'Quadras  (0)';
    DM_Tabelas.CDSQuadrasTemp.Close;
    DM_Tabelas.CDSIncorp.Close;
    atualiza_lotes_incorp;
    Exit;
  end;

  FCidade := DM_Tabelas.ZQLoteamento.FindField('cidade_idcidade');
  if (FCidade <> nil) and not FCidade.IsNull then
  begin
    if not DM_Tabelas.ZQCidade.Active then
      DM_Tabelas.ZQCidade.Open;
    VCidade := DM_Tabelas.ZQCidade.Lookup('idcidade', FCidade.Value, 'nomecid');
    VEstado := DM_Tabelas.ZQCidade.Lookup('idcidade', FCidade.Value, 'estado');
    if not VarIsNull(VCidade) then
      EUsuCidade.Text := VarToStr(VCidade);
    if not VarIsNull(VEstado) then
      EEstado.Text := VarToStr(VEstado);
  end;
  DM_Tabelas.CDSQuadrasTemp.Close;
  DM_Tabelas.CDSQuadrasTemp.CreateDataSet;
  DM_tabelas.ZQQuadras.First;
  While not DM_TAbelas.ZQQuadras.Eof do begin
    DM_Tabelas.CDSQuadrasTemp.Insert;
    DM_Tabelas.CDSQuadrasTempquadra.Value := DM_Tabelas.ZQquadras.FieldByName('descricao').AsString;
    DM_Tabelas.CDSQuadrasTempapelido.Value := DM_tabelas.ZQquadras.FieldByName('apedlido').AsString;
    DM_Tabelas.CDSQuadrasTemp.Post;
    DM_Tabelas.ZQQuadras.Next;
  end;
  GroupBox1.Caption := 'Quadras  ('+inttostr(DM_Tabelas.ZQQuadras.RecordCount)+')';
  DM_Tabelas.CDSIncorp.Close;
  DM_Tabelas.CDSIncorp.CreateDataSet;
  DM_Tabelas.ZQincorp_loteame.First;
  while not DM_Tabelas.ZQincorp_loteame.Eof do
  Begin
    DM_Tabelas.CDSIncorp.Insert;
    DM_Tabelas.CDSIncorpcodparticipante.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').AsLargeInt;
    DM_Tabelas.CDSIncorppercent.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('percent_parte').AsFloat;
    DM_Tabelas.CDSIncorpcodcontabancaria.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('codcontabancaria').AsInteger;
    DM_Tabelas.CDSIncorpdigito_dif.Value :=DM_Tabelas.ZQincorp_loteame.FieldByName('digito_dif').AsString;
    DM_Tabelas.CDSIncorp.Post;
    DM_Tabelas.ZQincorp_loteame.Next;
  end;
  atualiza_lotes_incorp;
End;

procedure TFrm_Loteamento.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (DM_Tabelas.ZQLoteamento.State in [DsInsert, DsEdit]) and (IntToStr(key)='46') and (DM_Tabelas.CDSIncorp.RecordCount>0) Then Begin
    DM_Tabelas.CDSIncorp.Delete;
  end;
  atualiza_lotes_incorp;
end;

procedure TFrm_Loteamento.Edit1Exit(Sender: TObject);
begin
  if DM_Tabelas.ZQLoteamento.State in [DsInsert, DsEdit] Then Begin
    if not AchaCLiente(200, 10,Eparticipante.Text) Then Begin
      Eparticipante.SetFocus;
      exit;
    end
    else Begin
      DM_Tabelas.CDSIncorp.Insert;
      DM_Tabelas.CDSIncorpcodparticipante.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
      DM_Tabelas.CDSIncorp.Post;
    end;
  end;
end;

procedure TFrm_Loteamento.DBGLotemamentoEnter(Sender: TObject);
begin
  Botoes;
end;

procedure TFrm_Loteamento.DBGLotemamentoKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  Botoes;
end;

procedure TFrm_Loteamento.DBGLotemamentoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Botoes;
end;

procedure TFrm_Loteamento.DBGLotemamentoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  FCidadeId: TField;
  TextoCidade: Variant;
  EstiloPincel: TBrushStyle;
begin
  if DM_Tabelas = nil then
    Exit;
  if not DM_Tabelas.ZQLoteamento.Active then
    Exit;
  if DM_Tabelas.ZQLoteamento.IsEmpty then
    Exit;
  if (Column = nil) or (Column.Field = nil) then
    Exit;

  if DBECod.Text = DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text Then Begin
    DBGLotemamento.Canvas.Brush.Color :=$006CFFFF;
    DBGLotemamento.Canvas.Font.Color := $00A80000;
    DBGLotemamento.Canvas.Font.Style := [FsBold];
  end;
  DBGLotemamento.DefaultDrawDataCell(Rect, Column.Field, State);

  { Estas duas colunas nao podem depender do valor calculado pelo dataset:
    na primeira linha ele ainda pode estar vazio. O desenho usa a mesma chave
    de cidade que ja alimenta o detalhe da tela, sem alterar o registro. }
  if (Column.Field = nil) or
     (not SameText(Column.Field.FieldName, 'nomecidade')) and
     (not SameText(Column.Field.FieldName, 'estado')) then
    Exit;
  if not DM_Tabelas.ZQCidade.Active then
    Exit;

  FCidadeId := DM_Tabelas.ZQLoteamento.FindField('cidade_idcidade');
  if (FCidadeId = nil) or FCidadeId.IsNull then
    Exit;

  if SameText(Column.Field.FieldName, 'nomecidade') then
    TextoCidade := DM_Tabelas.ZQCidade.Lookup('idcidade', FCidadeId.Value,
      'nomecid')
  else
    TextoCidade := DM_Tabelas.ZQCidade.Lookup('idcidade', FCidadeId.Value,
      'estado');

  if VarIsNull(TextoCidade) or VarIsEmpty(TextoCidade) then
    Exit;

  EstiloPincel := DBGLotemamento.Canvas.Brush.Style;
  DBGLotemamento.Canvas.Brush.Style := bsClear;
  try
    DBGLotemamento.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      VarToStr(TextoCidade));
  finally
    DBGLotemamento.Canvas.Brush.Style := EstiloPincel;
  end;
end;

procedure TFrm_Loteamento.DBGrid1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  atualiza_lotes_incorp;
end;

procedure TFrm_Loteamento.GroupBox1Exit(Sender: TObject);
begin
  if DM_Tabelas.CDSQuadrasTemp.state in [DsInsert, DsEdit] Then
    DM_Tabelas.CDSQuadrasTemp.Post;
end;

procedure TFrm_Loteamento.EContaBancariaExit(Sender: TObject);
begin
  if Acha_Contabancaria(400,40,EContaBancaria.Text) Then Begin
    EContaBancaria.Text := DM_Tabelas.ZQAchaContaBanc.FieldByName('n_conta').AsString+'-'+DM_Tabelas.ZQAchaContaBanc.FieldByName('n_conta_v').AsString+' - '+DM_Tabelas.ZQAchaContaBanc.FieldByName('n_no_banco').AsString+' - '+DM_Tabelas.ZQAchaContaBanc.FieldByName('nometitular').AsString+'|'+DM_Tabelas.ZQAchaContaBanc.FieldByName('n_dif_empreed').AsString;
    DM_Tabelas.CDSIncorp.Edit;
    DM_Tabelas.CDSIncorpcodcontabancaria.Value := DM_Tabelas.ZQAchaContaBanc.FieldByName('idconta_bancaria').AsLargeInt;
    DM_Tabelas.CDSIncorpdigito_dif.Value :=DM_Tabelas.ZQAchaContaBanc.FieldByName('n_dif_empreed').AsString; 
    DM_Tabelas.CDSIncorp.Post;
  end;
end;

procedure TFrm_Loteamento.DBGQruadraKeyPress(Sender: TObject; var Key: Char);
begin
  Key:= Upcase(Key);
end;

procedure TFrm_Loteamento.DBGQruadraKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (DM_Tabelas.ZQLoteamento.State in [DsEdit,DsInsert]) and (IntToStr(key)='46') and (DM_Tabelas.CDSQuadrasTemp.RecordCount>0) Then Begin
    DM_Tabelas.CDSQuadrasTemp.Delete;
  end;
end;

procedure TFrm_Loteamento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  // Shared data module datasets stay active. Vendas and Financeiro can still
  // receive grid notifications while this form is being destroyed.
  Frm_Loteamento := nil;
  Action := caFree;
end;

procedure TFrm_Loteamento.DBGrid3KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (DM_Tabelas.ZQLoteamento.State in [DsInsert, DsEdit]) and (IntToStr(key)='46') and (DM_Tabelas.zqprocuradores.RecordCount>0) Then
  Begin
    if simnao('Deseja Excluir esse Procurador ?','S') then
       DM_Tabelas.zqprocuradores.Delete;
  end;
end;

procedure TFrm_Loteamento.CBAssinaExit(Sender: TObject);
begin
  if simnao('Deseja Incluir esse Procurador ?','S') then
  begin
    if (DM_Tabelas.ZQLoteamento.State in [DsInsert]) and ( not empty(edprocuradores.Text)) and ( not empty(CBAssina.text)) Then
    Begin
      DM_Tabelas.zqprocuradores.insert;
      DM_Tabelas.ZQProcuradores.FieldByName('idloteamento').AsLargeInt:= DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt;
      DM_Tabelas.ZQProcuradores.FieldByName('idparti').AsLargeInt:= DM_Tabelas.CDSIncorpcodparticipante.Value;
      DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString:= edprocuradores.Text;
      DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString:= CBAssina.Text;
      DM_Tabelas.zqprocuradores.post;
      edprocuradores.Clear;
    end;

    if (DM_Tabelas.ZQLoteamento.State in [DsEdit]) and ( not empty(edprocuradores.Text)) and ( not empty(CBAssina.text)) Then
    Begin
      DM_Tabelas.zqprocuradores.insert;
      DM_Tabelas.ZQProcuradores.FieldByName('idloteamento').AsLargeInt:= DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').AsLargeInt;
      DM_Tabelas.ZQProcuradores.FieldByName('idparti').AsLargeInt:= DM_Tabelas.CDSIncorpcodparticipante.Value;
      DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString:= edprocuradores.Text;
      DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString:= CBAssina.Text;
      DM_Tabelas.zqprocuradores.post;
      edprocuradores.Clear;
    end;

  end
  else
  begin
    edprocuradores.Clear;
  end;
end;

procedure TFrm_Loteamento.DBGrid3DblClick(Sender: TObject);
begin
  if (DM_Tabelas.ZQLoteamento.State in [DsEdit]) then
  begin
    edprocuradores.Text:=DM_Tabelas.ZQProcuradores.FieldByName('nome').AsString;
    CBAssina.Text:=DM_Tabelas.ZQProcuradores.FieldByName('assina').AsString;
  end;
end;


procedure TFrm_Loteamento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureLoteamentoNomeCidadeField(DM_Tabelas);
  EnsureRuntimeFields(Self);
end;

end.
