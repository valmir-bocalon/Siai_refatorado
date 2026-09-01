unit Incorporador;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, DBCtrls, Mask,   dxButton,
  ExtCtrls, XBanner, DB;

type
  TFrm_Incorporador = class(TForm)
    Label32: TLabel;
    XBanner4: TXBanner;
    Label2: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label39: TLabel;
    DBECod: TDBEdit;
    DBENome: TDBEdit;
    DBEDocumento: TDBEdit;
    DBMemo1: TDBMemo;
    DBEEndereco: TDBEdit;
    EUFIncorp: TEdit;
    ECidIncorp: TEdit;
    DBGIncorporador: TDBGrid;
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
    procedure BtFecharClick(Sender: TObject);
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure Ativacampos;
    procedure Desativacampos;
    procedure FormShow(Sender: TObject);
    procedure DBGIncorporadorEnter(Sender: TObject);
    procedure DBGIncorporadorKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure botoes;
    procedure DBGIncorporadorMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure ECidIncorpExit(Sender: TObject);
    procedure DBGIncorporadorDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Incorporador: TFrm_Incorporador;

implementation

uses Tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Incorporador.BtFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Incorporador.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZQIncorporador.First;

end;

procedure TFrm_Incorporador.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZQIncorporador.Prior;

end;

procedure TFrm_Incorporador.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZQIncorporador.Next;
end;

procedure TFrm_Incorporador.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZQIncorporador.Last;
end;

procedure TFrm_Incorporador.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('Incorporadores','Gravar inclus�o ou edi��o','Incorporador: '+DM_Tabelas.ZQIncorporador.FieldByName('idincorporador').Text+' - '+DM_Tabelas.ZQIncorporador.FieldByName('nomeincorp').Value) Then Exit;
  DM_Tabelas.ZQIncorporador.Post;
  Desativacampos;
  DBGIncorporador.SetFocus;
end;

procedure TFrm_Incorporador.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Incorporadores','Cancelar inclus�o ou edi��o','Incorporador: '+DM_Tabelas.ZQIncorporador.FieldByName('idincorporador').Text+' - '+DM_Tabelas.ZQIncorporador.FieldByName('nomeincorp').Value) Then Exit;
  DM_Tabelas.ZQIncorporador.Cancel;
  Desativacampos;
  DBGIncorporador.SetFocus;
end;

procedure TFrm_Incorporador.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Incorporadores','Inclus�o','') Then Exit;
  Ativacampos;
  DBENome.SetFocus;
  DM_Tabelas.ZQIncorporador.Insert;
end;

procedure TFrm_Incorporador.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Incorporadores','Edi��o','Incorporador: '+DM_Tabelas.ZQIncorporador.FieldByName('idincorporador').Text+' - '+DM_Tabelas.ZQIncorporador.FieldByName('nomeincorp').Value) Then Exit;
  Ativacampos;
  DBENome.SetFocus;
  DM_Tabelas.ZQIncorporador.Edit;
end;

procedure TFrm_Incorporador.BTExcluirClick(Sender: TObject);
begin
  DM_Tabelas.ZQIncorporador.Delete;
end;

procedure TFrm_Incorporador.Ativacampos;
Begin
  BtPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BtProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtProcessar.Enabled := False;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BTExcluir.Enabled := False;
  BTPesquisar.Enabled := False;
  BtRelatorio.Enabled := False;
  DBGIncorporador.Enabled := False;
  DBENome.ReadOnly := False;
  DBEDocumento.ReadOnly := False;
  DBEEndereco.ReadOnly := False;
  ECidIncorp.ReadOnly := False;
End;

procedure TFrm_Incorporador.Desativacampos;
Begin
  BtPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BtProximo.Enabled := True;
  BtUltimo.Enabled := True;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtProcessar.Enabled := True;
  BtIncluir.Enabled := True;
  BtEditar.Enabled := True;
  BTExcluir.Enabled := True;
  BTPesquisar.Enabled := True;
  BtRelatorio.Enabled := True;
  DBGIncorporador.Enabled := True;
  DBENome.ReadOnly := True;
  DBEDocumento.ReadOnly := True;
  DBEEndereco.ReadOnly := True;
  ECidIncorp.ReadOnly := True;
End;
procedure TFrm_Incorporador.FormShow(Sender: TObject);
begin
  Desativacampos;
  DBGIncorporador.SetFocus;
end;

procedure TFrm_Incorporador.DBGIncorporadorEnter(Sender: TObject);
begin
  botoes;
end;

procedure TFrm_Incorporador.botoes;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQIncorporador.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQIncorporador.RecNo=DM_Tabelas.ZQIncorporador.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
End;

procedure TFrm_Incorporador.DBGIncorporadorKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  botoes;
end;

procedure TFrm_Incorporador.DBGIncorporadorMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes;
end;

procedure TFrm_Incorporador.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Incorporador.ECidIncorpExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQIncorporador.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ECidIncorp.text) Then
      ECidIncorp.SetFocus
    else Begin
      ECidIncorp.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').Value;
      EUFIncorp.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').Value;
      DM_Tabelas.ZQIncorporador.FieldByName('cidade_inco').Value := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').Value;
      DM_Tabelas.ZQIncorporadorestado.Value := DM_Tabelas.ZQAchaCidade.FieldByName('estado').Value;
    end;
  end;
end;

procedure TFrm_Incorporador.DBGIncorporadorDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQIncorporador.FieldByName('idincorporador').Text Then Begin
    DBGIncorporador.Canvas.Brush.Color :=$006CFFFF;
    DBGIncorporador.Canvas.Font.Color := $00A80000;
    DBGIncorporador.Canvas.Font.Style := [FsBold];
  end;
  DBGIncorporador.DefaultDrawDataCell(Rect, DBGIncorporador.columns[datacol].field, State);
end;


procedure TFrm_Incorporador.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
