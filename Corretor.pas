unit Corretor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, ExtCtrls, XBanner, StdCtrls, Grids, DBGrids,
  DBCtrls, Mask, ComCtrls, TabNotBk, DB, dxCore2;

type
  TFrm_Corretor = class(TForm)
    Label2: TLabel;
    XBanner2: TXBanner;
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
    BtProcessar: TdxButton;
    BTFechar: TdxButton;
    Pag_Corretor: TTabbedNotebook;
    Label1: TLabel;
    DBECod: TDBEdit;
    Label3: TLabel;
    DBENome: TDBEdit;
    Label4: TLabel;
    DBEDoc: TDBEdit;
    Label5: TLabel;
    DBECreci: TDBEdit;
    Label6: TLabel;
    DBETel1: TDBEdit;
    Label7: TLabel;
    DBETel2: TDBEdit;
    Label8: TLabel;
    DBEEndereco: TDBEdit;
    Label9: TLabel;
    DBEBairro: TDBEdit;
    Label10: TLabel;
    DBECep: TDBEdit;
    Label11: TLabel;
    DBEObservacao: TDBMemo;
    XBanner1: TXBanner;
    DBGCorretor: TDBGrid;
    Ecidade: TEdit;
    EEstado: TEdit;
    Label24: TLabel;
    Label12: TLabel;
    Registros: TLabel;
    LReg: TLabel;
    Label13: TLabel;
    DBEdit1: TDBEdit;
    procedure FormShow(Sender: TObject);
    procedure AtivarCampos;
    procedure DesativarCampos;
    procedure EcidadeExit(Sender: TObject);
    procedure Botoes;
    procedure Atualiza;
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
    procedure BTPesquisarClick(Sender: TObject);
    procedure BtRelatorioClick(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure DBGCorretorEnter(Sender: TObject);
    procedure DBGCorretorKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGCorretorMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGCorretorDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBEDocExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Corretor: TFrm_Corretor;

implementation

uses tabelas, funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Corretor.FormShow(Sender: TObject);
begin
  EnsureCorretorLookupFields(DM_Tabelas);
  DM_tabelas.ZQCorretor.Open;
  Pag_Corretor.PageIndex := 0;
  DesativarCampos;
  DBGCorretor.SetFocus;
end;

procedure TFrm_Corretor.AtivarCampos;
begin
  Ecidade.ReadOnly := False;
  DBENome.ReadOnly := False;
  DBEDoc.ReadOnly := False;
  DBECreci.ReadOnly := False;
  DBETel1.ReadOnly := False;
  DBETel2.ReadOnly := False;
  DBEEndereco.ReadOnly := False;
  DBEBairro.ReadOnly := False;
  DBECep.ReadOnly := False;
  DBEObservacao.ReadOnly := False;
  DBGCorretor.Enabled := False;
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
end;

procedure TFrm_Corretor.DesativarCampos;
begin
  Ecidade.ReadOnly := True;
  DBENome.ReadOnly := True;
  DBEDoc.ReadOnly := True;
  DBECreci.ReadOnly := True;
  DBETel1.ReadOnly := True;
  DBETel2.ReadOnly := True;
  DBEEndereco.ReadOnly := True;
  DBEBairro.ReadOnly := True;
  DBECep.ReadOnly := True;
  DBEObservacao.ReadOnly := True;
  DBGCorretor.Enabled := True;
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtProcessar.Enabled := True;
  BtIncluir.Enabled := True;
  BtEditar.Enabled := True;
  BtExcluir.Enabled := True;
  BtPesquisar.Enabled := True;
  BtRelatorio.Enabled := True;
end;
procedure TFrm_Corretor.botoes;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQCorretor.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQCorretor.RecNo=DM_Tabelas.ZQCorretor.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  Atualiza;
End;

procedure TFrm_Corretor.Atualiza;
Begin
  Ecidade.Text := DM_Tabelas.ZQCorretor.FieldByName('nomecidade').AsString;
  EEstado.Text := DM_Tabelas.ZQCorretor.FieldByName('estado').AsString;
  LReg.Caption := inttostr(DM_Tabelas.ZQCorretor.RecordCount);
End;

procedure TFrm_Corretor.EcidadeExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQCorretor.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(Ecidade.text) Then
      Ecidade.SetFocus
    else Begin
      Ecidade.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EEstado.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQCorretor.FieldByName('codcidade').AsInteger := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
end;

procedure TFrm_Corretor.BtPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQCorretor.First;
  Botoes;
end;

procedure TFrm_Corretor.BTAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQCorretor.Prior;
  Botoes;
end;

procedure TFrm_Corretor.BtProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQCorretor.Next;
  Botoes;
end;

procedure TFrm_Corretor.BtUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQCorretor.Last;
  Botoes;
end;

procedure TFrm_Corretor.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('Corretor','Gravar inclusão ou edição','Corretor: '+DM_Tabelas.ZQCorretor.FieldByName('idcorretor').Text+' - '+DM_Tabelas.ZQCorretor.FieldByName('nomecorretor').AsString) Then Exit;
  DM_tabelas.ZQCorretor.Post;
  Desativarcampos;
end;

procedure TFrm_Corretor.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Corretor','Cancelar inclusão ou edição','Corretor: '+DM_Tabelas.ZQCorretor.FieldByName('idcorretor').Text+' - '+DM_Tabelas.ZQCorretor.FieldByName('nomecorretor').AsString) Then Exit;
  DM_tabelas.ZQCorretor.Cancel;
  Desativarcampos;
end;

procedure TFrm_Corretor.BtProcessarClick(Sender: TObject);
begin
//  processar
end;

procedure TFrm_Corretor.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Corretor','Incluir','') Then Exit;
  DM_tabelas.ZQCorretor.Insert;
  Ativarcampos;
  DBENome.SetFocus;
end;

procedure TFrm_Corretor.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Corretor','Editar','Corretor: '+DM_Tabelas.ZQCorretor.FieldByName('idcorretor').Text+' - '+DM_Tabelas.ZQCorretor.FieldByName('nomecorretor').AsString) Then Exit;
  DM_tabelas.ZQCorretor.Edit;
  Ativarcampos;
  DBENome.SetFocus;
end;

procedure TFrm_Corretor.BTExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Corretor','Excluir','Corretor: '+DM_Tabelas.ZQCorretor.FieldByName('idcorretor').Text+' - '+DM_Tabelas.ZQCorretor.FieldByName('nomecorretor').AsString) Then Exit;
  DM_tabelas.ZQCorretor.Delete;
end;

procedure TFrm_Corretor.BTPesquisarClick(Sender: TObject);
begin
  //  Pesquisa
end;

procedure TFrm_Corretor.BtRelatorioClick(Sender: TObject);
begin
  if not Verif_senha('Corretor','Relatório','') Then Exit;
//  Relatorio
end;

procedure TFrm_Corretor.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Corretor.DBGCorretorEnter(Sender: TObject);
begin
  Botoes;
end;

procedure TFrm_Corretor.DBGCorretorKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  Botoes;
end;

procedure TFrm_Corretor.DBGCorretorMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Botoes;
end;

procedure TFrm_Corretor.DBGCorretorDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQCorretor.FieldByName('idcorretor').Text Then Begin
    DBGCorretor.Canvas.Brush.Color :=$006CFFFF;
    DBGCorretor.Canvas.Font.Color := $00A80000;
    DBGCorretor.Canvas.Font.Style := [FsBold];
  end;
  DBGCorretor.DefaultDrawDataCell(Rect, DBGCorretor.columns[datacol].field, State);
end;

procedure TFrm_Corretor.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Corretor.DBEDocExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQCorretor.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then 
    if not Verif_doc(DBEDoc.text,True) Then
      DBEDoc.SetFocus;
end;

procedure TFrm_Corretor.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_tabelas.ZQCorretor.close;
  Frm_Corretor:=nil;
  Action:=CaFree;
end;


procedure TFrm_Corretor.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
