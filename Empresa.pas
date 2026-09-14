unit Empresa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls,   dxButton, XBanner, ExtCtrls,
  GradBtn, ExtDlgs, DB, dxCore2;

type
  TFrmCad_Empresa = class(TForm)
    PPainelBotoes: TPanel;
    DXBExcluir: TdxButton;
    DXBEditar: TdxButton;
    DXBIncluir: TdxButton;
    DXBCancelar: TdxButton;
    DXBGravar: TdxButton;
    DXBPesquisar: TdxButton;
    DXBRelatorios: TdxButton;
    DXBFechar: TdxButton;
    DXBUltimo: TdxButton;
    DXBProximo: TdxButton;
    DXBAnterior: TdxButton;
    DXBPrimeiro: TdxButton;
    Label1: TLabel;
    Label3: TLabel;
    DBEFantazia: TDBEdit;
    Label4: TLabel;
    DBEEndereco: TDBEdit;
    Label5: TLabel;
    DBEBairro: TDBEdit;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    DBECep: TDBEdit;
    Label9: TLabel;
    DBEIfalgura: TDBEdit;
    Label10: TLabel;
    DBEIe: TDBEdit;
    Label11: TLabel;
    CBECnpj: TDBEdit;
    Label12: TLabel;
    DBEFone1: TDBEdit;
    Label13: TLabel;
    DBEEmail: TDBEdit;
    Label16: TLabel;
    DBERazao: TDBEdit;
    Label2: TLabel;
    DBILogo: TDBImage;
    Label14: TLabel;
    DBEdit1: TDBEdit;
    Label15: TLabel;
    DBEFone2: TDBEdit;
    Label17: TLabel;
    DBEFone3: TDBEdit;
    EscolhePict: TOpenPictureDialog;
    Label18: TLabel;
    Label19: TLabel;
    DBCBTipo: TDBComboBox;
    DBCBAtiva: TDBComboBox;
    DBEFilial: TDBEdit;
    Label20: TLabel;
    XBanner1: TXBanner;
    XBanner2: TXBanner;
    ECidcasou: TEdit;
    EEstCart: TEdit;
    procedure FormShow(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DXBIncluirClick(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure DXBCancelarClick(Sender: TObject);
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure Ativar_campos;
    procedure Desativar_campos;
    procedure DXBEditarClick(Sender: TObject);
    procedure DBILogoClick(Sender: TObject);
    procedure DBERazaoEnter(Sender: TObject);
    procedure DBEFantaziaEnter(Sender: TObject);
    procedure DBEEnderecoEnter(Sender: TObject);
    procedure DBEBairroEnter(Sender: TObject);
    procedure DBECepEnter(Sender: TObject);
    procedure DBEIfalguraEnter(Sender: TObject);
    procedure DBEIeEnter(Sender: TObject);
    procedure CBECnpjEnter(Sender: TObject);
    procedure DBCBTipoEnter(Sender: TObject);
    procedure DBEFone1Enter(Sender: TObject);
    procedure DBEFone2Enter(Sender: TObject);
    procedure DBEFone3Enter(Sender: TObject);
    procedure DBCBAtivaChange(Sender: TObject);
    procedure DBEEmailEnter(Sender: TObject);
    procedure DBCBAtivaEnter(Sender: TObject);
    procedure DBERazaoExit(Sender: TObject);
    procedure DBEFantaziaExit(Sender: TObject);
    procedure DBEEnderecoExit(Sender: TObject);
    procedure DBEBairroExit(Sender: TObject);
    procedure DBECepExit(Sender: TObject);
    procedure DBEIfalguraExit(Sender: TObject);
    procedure DBEIeExit(Sender: TObject);
    procedure CBECnpjExit(Sender: TObject);
    procedure DBCBTipoExit(Sender: TObject);
    procedure DBEFone1Exit(Sender: TObject);
    procedure DBEFone2Exit(Sender: TObject);
    procedure DBEFone3Exit(Sender: TObject);
    procedure DBCBAtivaExit(Sender: TObject);
    procedure DBEEmailExit(Sender: TObject);
    procedure DBEFilialEnter(Sender: TObject);
    procedure DBEFilialExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure ECidcasouExit(Sender: TObject);
    procedure ECidcasouEnter(Sender: TObject);
    procedure DBERazaoChange(Sender: TObject);
    procedure DXBExcluirClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AtualizarCidadeEmpresa;
    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCad_Empresa: TFrmCad_Empresa;

implementation

uses Tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmCad_Empresa.FormShow(Sender: TObject);
begin
  if not DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Open;
  DM_Tabelas.ZQEmpresa.close;
  DM_Tabelas.ZQEmpresa.open;
  AtualizarCidadeEmpresa;
  Desativar_campos;
end;

procedure TFrmCad_Empresa.AtualizarCidadeEmpresa;
var
  LCidade: TField;
  LNomeCidade: Variant;
  LEstado: Variant;
begin
  if (DM_Tabelas = nil) or (not DM_Tabelas.ZQEmpresa.Active) or  DM_Tabelas.ZQEmpresa.IsEmpty then
  begin
    ECidcasou.Clear;
    EEstCart.Clear;
    Exit;
  end;

  if not DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Open;

  ECidcasou.Clear;
  EEstCart.Clear;
  LCidade := DM_Tabelas.ZQEmpresa.FindField('cidade');
  if (LCidade = nil) or LCidade.IsNull then
    Exit;

  { A busca usa a chave física já gravada na empresa. Assim, o primeiro
    registro não depende do ciclo visual de campos calculados. }
  LNomeCidade := DM_Tabelas.ZQCidade.Lookup('idcidade', LCidade.Value,
    'nomecid');
  LEstado := DM_Tabelas.ZQCidade.Lookup('idcidade', LCidade.Value, 'estado');
  if not VarIsNull(LNomeCidade) and not VarIsEmpty(LNomeCidade) then
    ECidcasou.Text := VarToStr(LNomeCidade);
  if not VarIsNull(LEstado) and not VarIsEmpty(LEstado) then
    EEstCart.Text := VarToStr(LEstado);

  { Mantem compatibilidade com consultas que já disponibilizam os lookups. }
  if ECidcasou.Text = '' then
    ECidcasou.Text := DM_Tabelas.ZQEmpresa.FieldByName('nomecidade').AsString;
  if EEstCart.Text = '' then
    EEstCart.Text := DM_Tabelas.ZQEmpresa.FieldByName('estado').AsString;
end;

procedure TFrmCad_Empresa.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmCad_Empresa.DXBIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Empresa','Incluir Conta Bancária','') then exit;
  DM_Tabelas.ZQEmpresa.Insert;
  Ativar_campos;
  DBERazao.setFocus;
end;

procedure TFrmCad_Empresa.DXBGravarClick(Sender: TObject);
begin
  if not Verif_senha('Empresa','Confirmar Inclusão ou Edição (Botão GRAVAR)','Razão: '+DM_Tabelas.ZQEmpresa.FieldByName('razao').AsString) then
   exit;
  DM_Tabelas.ZQEmpresa.Post;
  Desativar_campos;
end;

procedure TFrmCad_Empresa.DXBCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Empresa','Cancelamento de Inclusão ou Edição (Botão CANCELAR) ','Razão: '+DM_Tabelas.ZQEmpresa.FieldByName('razao').AsString) then exit;
  DM_Tabelas.ZQEmpresa.Cancel;
  Desativar_campos;
end;

procedure TFrmCad_Empresa.DXBPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQEmpresa.First;
  AtualizarCidadeEmpresa;
  DXBPrimeiro.Enabled := false;
  DXBAnterior.Enabled := false;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
end;

procedure TFrmCad_Empresa.DXBAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQEmpresa.Prior;
  AtualizarCidadeEmpresa;
  if DM_Tabelas.ZQEmpresa.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
end;

procedure TFrmCad_Empresa.DXBProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQEmpresa.Next;
  AtualizarCidadeEmpresa;
  if DM_Tabelas.ZQEmpresa.RecNo = DM_Tabelas.ZQEmpresa.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
end;

procedure TFrmCad_Empresa.DXBUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQEmpresa.Last;
  AtualizarCidadeEmpresa;
  DXBProximo.Enabled  := false;
  DXBUltimo.Enabled   := false;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
end;

procedure TFrmCad_Empresa.Ativar_campos;
Begin
  DBCBTipo.ReadOnly    := False;
  DBCBAtiva.ReadOnly   := False;
  DBERazao.ReadOnly    := False;
  DBEFantazia.ReadOnly := False;
  DBEEndereco.ReadOnly := False;
  DBEBairro.ReadOnly   := False;
  ECidcasou.ReadOnly   := False;
  DBECep.ReadOnly      := False;
  DBEIfalgura.ReadOnly := False;
  DBEIe.ReadOnly       := False;
  CBECnpj.ReadOnly     := False;
  DBEFone1.ReadOnly    := False;
  DBEFone2.ReadOnly    := False;
  DBEFone3.ReadOnly    := False;
  DBEEmail.ReadOnly    := False;
  DBEFilial.ReadOnly   := False;
  DXBIncluir.Enabled   := False;
  DXBEditar.Enabled    := False;
  DXBExcluir.Enabled   := False;
  DXBPrimeiro.Enabled  := False;
  DXBAnterior.Enabled  := False;
  DXBProximo.Enabled   := False;
  DXBUltimo.Enabled    := False;
  DXBGravar.Enabled    := True;
  DXBCancelar.Enabled  := True;
End;

procedure TFrmCad_Empresa.Desativar_campos;
Begin
  DBCBTipo.ReadOnly    := True;
  DBCBAtiva.ReadOnly   := True;
  DBERazao.ReadOnly    := True;
  DBEFantazia.ReadOnly := True;
  DBEEndereco.ReadOnly := True;
  DBEBairro.ReadOnly   := True;
  ECidcasou.ReadOnly   := True;
  DBECep.ReadOnly      := True;
  DBEIfalgura.ReadOnly := True;
  DBEIe.ReadOnly       := True;
  CBECnpj.ReadOnly     := True;
  DBEFone1.ReadOnly    := True;
  DBEFone2.ReadOnly    := True;
  DBEFone3.ReadOnly    := True;
  DBEEmail.ReadOnly    := True;
  DBEFilial.ReadOnly   := True;
  DXBIncluir.Enabled   := True;
  DXBEditar.Enabled    := True;
  DXBExcluir.Enabled   := True;
  DXBPrimeiro.Enabled  := True;
  DXBAnterior.Enabled  := True;
  DXBProximo.Enabled   := True;
  DXBUltimo.Enabled    := True;
  DXBGravar.Enabled    := False;
  DXBCancelar.Enabled  := False;
End;

procedure TFrmCad_Empresa.DXBEditarClick(Sender: TObject);
begin
  if not Verif_senha('Empresa','Editar','Código: '+DM_Tabelas.ZQEmpresa.FieldByName('codigo').Text+'  Razão: '+DM_Tabelas.ZQEmpresa.FieldByName('razao').AsString) then Exit;
  DM_Tabelas.ZQEmpresa.Edit;
  Ativar_campos;
  DBERazao.setFocus;
end;

procedure TFrmCad_Empresa.DBILogoClick(Sender: TObject);
begin
  if DM_Tabelas.ZQEmpresa.State in [dsEdit, dsInsert] then Begin
//    EscolhePict.InitialDir := 'D:\sistemas\projeto Tabelas';
    EscolhePict.Execute;
    TBlobField(DM_Tabelas.ZQEmpresa.FieldByName('logo')).LoadFromFile(EscolhePict.FileName);
//    DM_Tabelas.ZQEmpresalogo.LoadFromFile(EscolhePict.FileName);
  end;
end;

procedure TFrmCad_Empresa.DBERazaoEnter(Sender: TObject);
begin
DBERazao.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEFantaziaEnter(Sender: TObject);
begin
DBEFantazia.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEEnderecoEnter(Sender: TObject);
begin
DBEEndereco.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEBairroEnter(Sender: TObject);
begin
DBEBairro.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBECepEnter(Sender: TObject);
begin
DBECep.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEIfalguraEnter(Sender: TObject);
begin
DBEIfalgura.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEIeEnter(Sender: TObject);
begin
DBEIe.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.CBECnpjEnter(Sender: TObject);
begin
CBECnpj.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBCBTipoEnter(Sender: TObject);
begin
DBCBTipo.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEFone1Enter(Sender: TObject);
begin
DBEFone1.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEFone2Enter(Sender: TObject);
begin
DBEFone2.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEFone3Enter(Sender: TObject);
begin
DBEFone3.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBCBAtivaChange(Sender: TObject);
begin
DBCBAtiva.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEEmailEnter(Sender: TObject);
begin
DBEEmail.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBCBAtivaEnter(Sender: TObject);
begin
DBCBAtiva.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBERazaoExit(Sender: TObject);
begin
DBERazao.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEFantaziaExit(Sender: TObject);
begin
DBEFantazia.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEEnderecoExit(Sender: TObject);
begin
DBEEndereco.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEBairroExit(Sender: TObject);
begin
DBEBairro.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBECepExit(Sender: TObject);
begin
DBECep.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEIfalguraExit(Sender: TObject);
begin
DBEIfalgura.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEIeExit(Sender: TObject);
begin
DBEIe.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.CBECnpjExit(Sender: TObject);
begin
CBECnpj.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBCBTipoExit(Sender: TObject);
begin
DBCBTipo.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEFone1Exit(Sender: TObject);
begin
DBEFone1.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEFone2Exit(Sender: TObject);
begin
DBEFone2.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEFone3Exit(Sender: TObject);
begin
DBEFone3.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBCBAtivaExit(Sender: TObject);
begin
DBCBAtiva.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEEmailExit(Sender: TObject);
begin
DBEEmail.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.DBEFilialEnter(Sender: TObject);
begin
  DBEFilial.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBEFilialExit(Sender: TObject);
begin
  DBEFilial.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCad_Empresa.ECidcasouExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQEmpresa.State in [DsInsert, DsEdit]) and not(DXBCancelar.Focused or DXBFechar.Focused) Then Begin
    if not Acha_Cidade(ECidcasou.text) Then
      ECidcasou.SetFocus
    else Begin
      ECidcasou.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EEstCart.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQEmpresa.FieldByName('cidade').AsLargeInt := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
//      DM_Tabelas.ZQEmpresauf.Value:=EEstCart.Text;
    end;
  end;
  ECidcasou.Color:= ClWindow;
end;

procedure TFrmCad_Empresa.ECidcasouEnter(Sender: TObject);
begin
ECidcasou.Color:= $00FFF0E1;
end;

procedure TFrmCad_Empresa.DBERazaoChange(Sender: TObject);
begin
  AtualizarCidadeEmpresa;
end;

procedure TFrmCad_Empresa.DXBExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Empresa','Deletar','Código: '+DM_Tabelas.ZQEmpresa.FieldByName('codigo').Text+'  Razão: '+DM_Tabelas.ZQEmpresa.FieldByName('razao').AsString) then Exit;
  if simnao('Confirma a exclusão da empresa ?','SIM') then
  begin
    DM_Tabelas.ZQEmpresa.Cancel;
    DM_Tabelas.ZQEmpresa.Delete;
    Ativar_campos;
    DBERazao.setFocus;
  end;
  DM_Tabelas.ZQEmpresa.Cancel;
end;

procedure TFrmCad_Empresa.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQEmpresa.close;
  FrmCad_Empresa:=nil;
  Action:=CaFree;
end;


procedure TFrmCad_Empresa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
