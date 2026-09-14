unit Usuario;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, wwdbdatetimepicker, Mask, DBCtrls,
  ComCtrls, TabNotBk,   dxButton, ExtCtrls, XBanner, DB, dxCore2;

type
  TFrm_Usuario = class(TForm)
    XBanner3: TXBanner;
    Label1: TLabel;
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
    Pag_Usuario: TTabbedNotebook;
    XBanner10: TXBanner;
    Label5: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label24: TLabel;
    Label30: TLabel;
    DBEUsuCod: TDBEdit;
    DBEUsuNome: TDBEdit;
    DBEUsuCargo: TDBEdit;
    DBEUsuEnde: TDBEdit;
    DBEUsuBairro: TDBEdit;
    DBEUsutel1: TDBEdit;
    DBEUsutel2: TDBEdit;
    DBEUsuAdm: TwwDBDateTimePicker;
    DBEUsuDemit: TwwDBDateTimePicker;
    EUsuCidade: TEdit;
    EUsuUF: TEdit;
    DBEUsuApelido: TDBEdit;
    XBanner11: TXBanner;
    Label31: TLabel;
    Label8: TLabel;
    DBEUsuSenha: TDBEdit;
    Esenha2: TEdit;
    DBGAgrupaGrupo: TDBGrid;
    DBGPermis: TDBGrid;
    DBEGru: TDBEdit;
    DBGUsuario: TDBGrid;
    procedure EUsuCidadeExit(Sender: TObject);
    procedure DBEUsuDemitExit(Sender: TObject);
    procedure DBEUsuSenhaExit(Sender: TObject);
    procedure Esenha2Exit(Sender: TObject);
    procedure DBGAgrupaGrupoDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBGAgrupaGrupoEnter(Sender: TObject);
    procedure DBGAgrupaGrupoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGAgrupaGrupoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGPermisDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGPermisKeyPress(Sender: TObject; var Key: Char);
    procedure DBGPermisMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure ativausuario;
    procedure desativausuario;
    procedure atualiza_usuario;
    procedure botoes_usuario;
    procedure BTFecharClick(Sender: TObject);
    procedure marca_permissoes;
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBGUsuarioEnter(Sender: TObject);
    procedure DBGUsuarioMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGUsuarioKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure DBGUsuarioDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Pag_UsuarioClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Usuario: TFrm_Usuario;

implementation

uses tabelas, Funcoes, Principal, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Usuario.EUsuCidadeExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQUsuario.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(EUsuCidade.text) Then
      EUsuCidade.SetFocus
    else Begin
      EUsuCidade.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EUsuUF.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQUsuario.FieldByName('cidade').AsLargeInt := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
end;

procedure TFrm_Usuario.DBEUsuDemitExit(Sender: TObject);
begin
  if DM_Tabelas.ZQUsuario.State in [DsInsert] Then Begin
    Pag_Usuario.PageIndex := 1;
    DBEUsuSenha.SetFocus;
  end;
end;

procedure TFrm_Usuario.DBEUsuSenhaExit(Sender: TObject);
begin
  Esenha2.SetFocus;
end;

procedure TFrm_Usuario.Esenha2Exit(Sender: TObject);
begin
  if Esenha2.Text <> DM_Tabelas.ZQUsuario.FieldByName('senha').AsString Then Begin
    showmessage('Senha não confere...  Confirme a senha!!');
    DBEUsuSenha.SetFocus;
  end;
end;

procedure TFrm_Usuario.DBGAgrupaGrupoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEGru.Text = DM_Tabelas.ZQAgrupaGrupo.FieldByName('grupo').Text Then Begin
    DBGAgrupaGrupo.Canvas.Brush.Color :=$006CFFFF;
    DBGAgrupaGrupo.Canvas.Font.Color := $00A80000;
    DBGAgrupaGrupo.Canvas.Font.Style := [FsBold];
  end;
  DBGAgrupaGrupo.DefaultDrawDataCell(Rect, DBGAgrupaGrupo.columns[datacol].field, State);
end;

procedure TFrm_Usuario.DBGAgrupaGrupoEnter(Sender: TObject);
begin
  DM_Tabelas.CDSLocaisSenha.Filter := 'grupo = '+quotedstr(DM_Tabelas.ZQAgrupaGrupo.FieldByName('grupo').AsString);
  DM_Tabelas.CDSLocaisSenha.Filtered := True;
end;

procedure TFrm_Usuario.DBGAgrupaGrupoKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  DM_Tabelas.CDSLocaisSenha.Filter := 'grupo = '+quotedstr(DM_Tabelas.ZQAgrupaGrupo.FieldByName('grupo').AsString);
  DM_Tabelas.CDSLocaisSenha.Filtered := True;
end;

procedure TFrm_Usuario.DBGAgrupaGrupoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  DM_Tabelas.CDSLocaisSenha.Filter := 'grupo = '+quotedstr(DM_Tabelas.ZQAgrupaGrupo.FieldByName('grupo').AsString);
  DM_Tabelas.CDSLocaisSenha.Filtered := True;
end;

procedure TFrm_Usuario.DBGPermisDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if Column.Field = DM_Tabelas.CDSLocaisSenhapedir then begin
    DBGPermis.Canvas.FillRect(Rect);
//    Frm_principal.ImageList1.Draw(DBGPermis.Canvas, Rect.Left + 15, Rect.Top + 1, 0);
    if DM_Tabelas.CDSLocaisSenhapedir.Value = 'S' then
      Frm_principal.ImageList1.Draw(DBGPermis.Canvas, Rect.Left + 15, Rect.Top + 1, 1)
    else
      Frm_principal.ImageList1.Draw(DBGPermis.Canvas, Rect.Left + 15, Rect.Top + 1, 0);
  end;
  if Column.Field = DM_Tabelas.CDSLocaisSenhapermitir then begin
    DBGPermis.Canvas.FillRect(Rect);
//    Frm_principal.ImageList1.Draw(DBGPermis.Canvas, Rect.Left + 20, Rect.Top + 1, 0);
    if DM_Tabelas.CDSLocaisSenhapermitir.Value = 'S' then
      Frm_principal.ImageList1.Draw(DBGPermis.Canvas, Rect.Left + 20, Rect.Top + 1, 0)
    else
      Frm_principal.ImageList1.Draw(DBGPermis.Canvas, Rect.Left + 20, Rect.Top + 1, 1);
  end;

end;

procedure TFrm_Usuario.DBGPermisKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then Begin
    marca_permissoes;
  end;

end;

procedure TFrm_Usuario.DBGPermisMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  marca_permissoes;
end;

procedure TFrm_Usuario.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQUsuario.Next;

end;

procedure TFrm_Usuario.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQUsuario.Last;

end;

procedure TFrm_Usuario.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('Usuário','Gravar inclusão ou edição','Usuario: '+DM_Tabelas.ZQUsuario.FieldByName('nome').AsString) Then Exit;
    DM_Tabelas.ZQUsuario.Post;
  desativausuario;
  DBGUsuario.SetFocus;

end;

procedure TFrm_Usuario.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Usuário','Cancelar inclusão ou edição','Usuario: '+DM_Tabelas.ZQUsuario.FieldByName('nome').AsString) Then Exit;
  DM_Tabelas.ZQUsuario.Cancel;
  desativausuario;
  DBGUsuario.SetFocus;
end;

procedure TFrm_Usuario.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Usuário','inclusão','') Then Exit;
  Pag_Usuario.PageIndex := 0;
  DM_Tabelas.ZQUsuario.Insert;
  ativausuario;
  DBEUsuNome.SetFocus;
end;

procedure TFrm_Usuario.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Usuário','Edição','Usuario: '+DM_Tabelas.ZQUsuario.FieldByName('idusuario').Text+' - '+DM_Tabelas.ZQUsuario.FieldByName('nome').AsString) Then Exit;
  DM_Tabelas.ZQUsuario.Edit;
  ativausuario;
  if Pag_Usuario.PageIndex = 0 then
    DBEUsuNome.SetFocus
  else
    DBEUsuSenha.SetFocus;
end;

procedure TFrm_Usuario.BTExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Usuário','Exclusão','Usuario: '+DM_Tabelas.ZQUsuario.FieldByName('idusuario').Text+' - '+DM_Tabelas.ZQUsuario.FieldByName('nome').AsString) Then Exit;
  if simnao('Confirma a exclusão do usuario ?','SIM') then
     DM_Tabelas.ZQUsuario.Delete;
end;

procedure TFrm_Usuario.ativausuario;
Begin
  DBEUsuNome.ReadOnly    := False;
  DBEUsuCargo.ReadOnly   := False;
  DBEUsuEnde.ReadOnly    := False;
  DBEUsuBairro.ReadOnly  := False;
  DBEUsutel1.ReadOnly    := False;
  DBEUsutel2.ReadOnly    := False;
  DBEUsuAdm.ReadOnly     := False;
  DBEUsuDemit.ReadOnly   := False;
  DBEUsuSenha.ReadOnly   := False;
  DBEUsuApelido.ReadOnly := False;
  BTPrimeiro.Enabled     := False;
  BTAnterior.Enabled     := False;
  BTProximo.Enabled      := False;
  BtUltimo.Enabled       := False;
  BtGravar.Enabled       := True;
  BtCancelar.Enabled     := True;
  BtFechar.Enabled       := False;
  BtIncluir.Enabled      := False;
  BtEditar.Enabled       := False;
  BtExcluir.Enabled      := False;
  BtPesquisar.Enabled    := False;
  BtRelatorio.Enabled    := False;
  DBGAgrupaGrupo.Enabled := False;
  DBGPermis.Enabled      := False;
End;

procedure TFrm_Usuario.desativausuario;
Begin
  DBEUsuNome.ReadOnly    := True;
  DBEUsuCargo.ReadOnly   := True;
  DBEUsuEnde.ReadOnly    := True;
  DBEUsuBairro.ReadOnly  := True;
  DBEUsutel1.ReadOnly    := True;
  DBEUsutel2.ReadOnly    := True;
  DBEUsuAdm.ReadOnly     := True;
  DBEUsuDemit.ReadOnly   := True;
  DBEUsuSenha.ReadOnly   := True;
  DBEUsuApelido.ReadOnly := True;
  BTPrimeiro.Enabled     := true;
  BTAnterior.Enabled     := true;
  BTProximo.Enabled      := true;
  BtUltimo.Enabled       := true;
  BtGravar.Enabled       := False;
  BtCancelar.Enabled     := False;
  BtFechar.Enabled       := true;
  BtIncluir.Enabled      := true;
  BtEditar.Enabled       := true;
  BtExcluir.Enabled      := true;
  BtPesquisar.Enabled    := true;
  BtRelatorio.Enabled    := true;
  DBGAgrupaGrupo.Enabled := True;
  DBGPermis.Enabled      := True;
End;

procedure TFrm_Usuario.atualiza_usuario;
var
  FCidade: TField;
  CidadeNome: Variant;
  CidadeUF: Variant;
Begin
  EUsuCidade.Clear;
  EUsuUF.Clear;

  { cidnome e estadocid sao apresentacoes de lookup, nao colunas de usuario. }
  FCidade := DM_Tabelas.ZQUsuario.FindField('cidade');
  if (FCidade <> nil) and not FCidade.IsNull then
  begin
    if not DM_Tabelas.ZQCidade.Active then
      DM_Tabelas.ZQCidade.Open;
    CidadeNome := DM_Tabelas.ZQCidade.Lookup('idcidade', FCidade.Value, 'nomecid');
    CidadeUF := DM_Tabelas.ZQCidade.Lookup('idcidade', FCidade.Value, 'estado');
    if not VarIsNull(CidadeNome) then
      EUsuCidade.Text := VarToStr(CidadeNome);
    if not VarIsNull(CidadeUF) then
      EUsuUF.Text := VarToStr(CidadeUF);
  end;
  if Pag_Usuario.PageIndex = 1 Then Begin
    DM_Tabelas.CDSLocaisSenha.Close;
    DM_Tabelas.CDSLocaisSenha.CreateDataSet;
    DM_Tabelas.ZQLoc_Senha.First;
    DM_Tabelas.ZQLoc_Senha.DisableControls;
    while not DM_Tabelas.ZQLoc_Senha.Eof do begin
      DM_Tabelas.CDSLocaisSenha.Insert;
      DM_Tabelas.CDSLocaisSenhaidlocais.Value := DM_Tabelas.ZQLoc_Senha.FieldByName('idlocais_com_senha').AsLargeInt;
      DM_Tabelas.CDSLocaisSenhaGrupo.Value := DM_Tabelas.ZQLoc_Senha.FieldByName('grupo').AsString;
      DM_Tabelas.CDSLocaisSenhadescricao.Value := DM_Tabelas.ZQLoc_Senha.FieldByName('descricao').AsString;
      DM_Tabelas.CDSLocaisSenhapedir.Value := DM_Tabelas.ZQLoc_Senha.FieldByName('pedesenha').AsString;
      if DM_Tabelas.ZQPermissoes.Locate('local_senha,usuario',VarArrayOf([DM_Tabelas.ZQLoc_Senha.FieldByName('idlocais_com_senha').AsLargeInt,DM_Tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt]),[]) Then
        DM_Tabelas.CDSLocaisSenhapermitir.Value := 'S'
      else
        DM_Tabelas.CDSLocaisSenhapermitir.Value := 'N';
      DM_Tabelas.CDSLocaisSenha.Post;
      DM_Tabelas.ZQLoc_Senha.Next;
    end;
    DM_Tabelas.ZQLoc_Senha.EnableControls;
    DM_Tabelas.CDSLocaisSenha.IndexFieldNames := 'descricao';
    DM_Tabelas.CDSLocaisSenha.Filter := 'grupo = '+quotedstr(DM_Tabelas.ZQAgrupaGrupo.FieldByName('grupo').AsString);
    DM_Tabelas.CDSLocaisSenha.Filtered := True;
  end;
end;

procedure TFrm_Usuario.botoes_usuario;
Begin
  BtPrimeiro.Enabled := True;
  BtAnterior.Enabled := True;
  BtProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQUsuario.Recno <=1 Then Begin
    BtPrimeiro.Enabled := False;
    BtAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQUsuario.RecNo=DM_Tabelas.ZQUsuario.RecordCount Then Begin
    BtProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  atualiza_usuario;
End;

procedure TFrm_Usuario.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Usuario.marca_permissoes;
Begin
  if DBGPermis.SelectedField.fieldname = 'pedir' Then Begin
    DM_Tabelas.CDSLocaisSenha.Edit;
    if DM_Tabelas.CDSLocaisSenhapedir.Value = 'N' then
      DM_Tabelas.CDSLocaisSenhapedir.Value := 'S'
    else
      DM_Tabelas.CDSLocaisSenhapedir.Value := 'N';
    DM_Tabelas.CDSLocaisSenha.Post;
    DM_Tabelas.ZQLoc_Senha.Locate('idlocais_com_senha',DM_Tabelas.CDSLocaisSenhaidlocais.Value,[]);
    DM_Tabelas.ZQLoc_Senha.Edit;
    DM_Tabelas.ZQLoc_Senha.FieldByName('pedesenha').AsString := DM_Tabelas.CDSLocaisSenhapedir.Value;
    DM_tabelas.ZQLoc_Senha.Post;
  end;
  if DBGPermis.SelectedField.fieldname = 'permitir' Then Begin
    DM_Tabelas.CDSLocaisSenha.Edit;
    if (DM_Tabelas.CDSLocaisSenhapermitir.Value = 'N') or (empty(DM_Tabelas.CDSLocaisSenhapermitir.Value)) then
      DM_Tabelas.CDSLocaisSenhapermitir.Value := 'S'
    else
      DM_Tabelas.CDSLocaisSenhapermitir.Value := 'N';
    DM_Tabelas.CDSLocaisSenha.Post;
    if DM_Tabelas.CDSLocaisSenhapermitir.Value = 'N' Then Begin
      if DM_Tabelas.ZQPermissoes.Locate('local_senha,usuario',VarArrayOf([DM_Tabelas.CDSLocaisSenhaidlocais.Value,DM_Tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt]),[]) Then
        DM_Tabelas.ZQPermissoes.Delete;
    end
    else Begin
      DM_Tabelas.ZQPermissoes.Insert;
      DM_Tabelas.ZQPermissoes.FieldByName('usuario').AsLargeInt := DM_Tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt;
      DM_Tabelas.ZQPermissoes.FieldByName('local_senha').AsLargeInt := DM_Tabelas.CDSLocaisSenhaidlocais.Value;
      DM_Tabelas.ZQPermissoes.Post;
    end;
  end;
end;

procedure TFrm_Usuario.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Usuario.DBGUsuarioEnter(Sender: TObject);
begin
  botoes_usuario;
end;

procedure TFrm_Usuario.DBGUsuarioMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes_usuario;
end;

procedure TFrm_Usuario.DBGUsuarioKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes_usuario;
end;

procedure TFrm_Usuario.FormShow(Sender: TObject);
begin
  DM_Tabelas.ZQAgrupaGrupo.Open;
  DM_Tabelas.CDSLocaisSenha.close;
  DM_Tabelas.CDSLocaisSenha.createdataset;
  
  Pag_Usuario.PageIndex := 0;
  desativausuario;
  DBGUsuario.SetFocus;
end;

procedure TFrm_Usuario.DBGUsuarioDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEUsuCod.Text = DM_Tabelas.ZQUsuario.FieldByName('idusuario').Text Then Begin
    DBGUsuario.Canvas.Brush.Color :=$006CFFFF;
    DBGUsuario.Canvas.Font.Color := $00A80000;
    DBGUsuario.Canvas.Font.Style := [FsBold];
  end;
  DBGUsuario.DefaultDrawDataCell(Rect, DBGUsuario.columns[datacol].field, State);
end;

procedure TFrm_Usuario.Pag_UsuarioClick(Sender: TObject);
begin
  atualiza_usuario;
end;

procedure TFrm_Usuario.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQAgrupaGrupo.close;
  DM_Tabelas.CDSLocaisSenha.close;
  Frm_Usuario:=nil;
  Action:=CaFree;

end;


procedure TFrm_Usuario.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
